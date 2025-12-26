using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;

// ===== WPF алиасы (чтобы не конфликтовали с Autodesk.Revit.DB.Grid и др.) =====
using Wpf = System.Windows;
using WpfControls = System.Windows.Controls;

namespace Palmsi.RevitTools
{
    /// <summary>
    /// Экспорт 2D сегментов из всех планов на активном листе (хост + видимые Revit-ссылки)
    /// с раздельным слоем «вычитаний» (cutouts) и WPF-окном выбора шахт.
    /// </summary>
    [Transaction(TransactionMode.Manual), Regeneration(RegenerationOption.Manual)]
    public class ExportCurvesCommand : IExternalCommand
    {

        // ---------- базовые настройки ----------
        static readonly bool FORCE_MM = true;
        const int ROUND_DEC = 5;
        const double MIN_SEG_MM = 0.5;
        const double DEDUP_SNAP_MM = 0.5;
        const string OUTPUT_SCALE_MODE = "model"; // "paper" | "unified" | "model"
        const int UNIFIED_SCALE = 200;
        // окно поиска соседних витражных панелей вокруг двери (вдоль стены), мм
        const double CURTAIN_SAMPLE_ALONG_T_MM = 2500.0;   // ~2.5 м в каждую сторону
        // допустимое смещение панелей от центра двери по нормали, мм
        const double CURTAIN_SAMPLE_MAX_N_MM = 1500.0;     // до 1.5 м
        // минимальная толщина панели, при которой считаем, что нашли нормальную пару линий, мм
        const double CURTAIN_MIN_PANEL_THICKNESS_MM = 10.0;
        // допуск, на каком расстоянии витражный сегмент считаем лежащим "на границе помещения" (мм)
        const double ROOM_CURTAIN_DIST_TOL_MM = 150.0;

        // --- OPA: если Revit-граница не доходит до витража ---
        // максимально допустимый "недолёт" (мм), при котором мы заменяем кусок room boundary на витраж
        static readonly double sOpaCurtainReplaceMaxFt = UnitUtils.ConvertToInternalUnits(800.0, UnitTypeId.Millimeters);

        // мостик считаем нужным только если реально есть зазор
        static readonly double sOpaBridgeEpsFt = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters);


        // максимальная длина «ступеньки» границы помещения вдоль витража, которую считаем артефактом (мм)
        const double CURTAIN_STEP_MAX_LEN_MM = 200.0;

        // смещение наружной линии двери по нормали (мм)
        // отрицательное значение сдвигает линию "внутрь" к витражу
        const double DOOR_OUTER_SHIFT_MM = -10.0;
        // --- безопасная замена (XYZ, XYZ) без System.ValueTuple ---
        struct Seg2
        {
            public XYZ A;
            public XYZ B;

            // true, если сегмент принадлежит витражу (панель, импост и т.п.)
            public bool IsCurtain;

            // Id корневой витражной стены (Wall.Id.IntegerValue).
            // 0 — если не смогли определить.
            public int CurtainRootId;

            public Seg2(XYZ a, XYZ b, bool isCurtain = false, int curtainRootId = 0)
            {
                A = a;
                B = b;
                IsCurtain = isCurtain;
                CurtainRootId = curtainRootId;
            }
        }

        class CurtainGraph
        {
            public class Node
            {
                public XYZ P;
                public List<int> Edges = new List<int>();
            }

            public class Edge
            {
                public int A;
                public int B;
                public double Len;
                public XYZ Mid;
            }

            public Dictionary<string, int> NodeIndexByKey = new Dictionary<string, int>();
            public List<Node> Nodes = new List<Node>();
            public List<Edge> Edges = new List<Edge>();
            public XYZ MainDir = new XYZ(1, 0, 0); // главная ось витража (по самой длинной панели)

            public int AddNode(XYZ p)
            {
                string key = ExportCurvesCommand.KeyForPoint(p); // уже есть
                int idx;
                if (!NodeIndexByKey.TryGetValue(key, out idx))
                {
                    idx = Nodes.Count;
                    var n = new Node();
                    n.P = p;
                    Nodes.Add(n);
                    NodeIndexByKey[key] = idx;
                }
                return idx;
            }

            public void AddEdge(int ia, int ib)
            {
                if (ia == ib) return;

                // защита от дублей
                foreach (int eIdx in Nodes[ia].Edges)
                {
                    Edge e = Edges[eIdx];
                    if ((e.A == ia && e.B == ib) || (e.A == ib && e.B == ia))
                        return;
                }

                var edge = new Edge();
                edge.A = ia;
                edge.B = ib;

                var pa = Nodes[ia].P;
                var pb = Nodes[ib].P;

                edge.Len = pa.DistanceTo(pb);
                edge.Mid = new XYZ(0.5 * (pa.X + pb.X), 0.5 * (pa.Y + pb.Y), 0);

                int index = Edges.Count;
                Edges.Add(edge);
                Nodes[ia].Edges.Add(index);
                Nodes[ib].Edges.Add(index);
            }

            public void ComputeMainDirection()
            {
                double maxLen = 0.0;
                XYZ bestDir = null;

                foreach (Edge e in Edges)
                {
                    if (e.Len > maxLen)
                    {
                        maxLen = e.Len;
                        XYZ pa = Nodes[e.A].P;
                        XYZ pb = Nodes[e.B].P;
                        XYZ dir = new XYZ(pb.X - pa.X, pb.Y - pa.Y, 0);

                        XYZ dirN;
                        if (ExportCurvesCommand.TryNormalize2D(dir, out dirN))
                            bestDir = dirN;
                    }
                }

                if (bestDir != null)
                    MainDir = bestDir;
            }
        }

        static string KeyForSeg(Seg2 s) => KeyFor(SnapXY(s.A), SnapXY(s.B));

        // Ключ для узла графа по точке
        static string KeyForPoint(XYZ p)
        {
            var s = SnapXY(new XYZ(p.X, p.Y, 0));
            var inv = CultureInfo.InvariantCulture;
            return s.X.ToString("R", inv) + "|" + s.Y.ToString("R", inv);
        }

        // --- Safe wrappers for GetLinkOverrides (Revit 2024+) ---
        static bool TryGetLinkOverridesViaReflection(View view, ElementId linkElemId, out object rlgs)
        {
            rlgs = null;
            try
            {
                var mi = typeof(View).GetMethod(
                    "GetLinkOverrides",
                    BindingFlags.Public | BindingFlags.Instance,
                    binder: null,
                    types: new[] { typeof(ElementId) },
                    modifiers: null);

                if (mi == null) return false; // API < 2024
                rlgs = mi.Invoke(view, new object[] { linkElemId });
                return rlgs != null;
            }
            catch { return false; }
        }
        // ---------- ПИН ДВЕРЕЙ В ВИТРАЖАХ ----------
        // ---------- ПИН ВСЕХ ДВЕРЕЙ В ХОСТ-ДОКУМЕНТЕ ----------
        // Закрепляет любой экземпляр, который считается "двереподобным"
        // по нашей функции IsDoorLikeInstance (обычные двери, витражные двери,
        // панели витража-двери и т.п.).
        // Закрепляет двереподобные экземпляры и возвращает список тех, у кого состояние изменили с false -> true
        static List<ElementId> PinAllDoors(Document doc)
        {
            var changed = new List<ElementId>();

            var insts = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>();

            foreach (var fi in insts)
            {
                if (fi == null) continue;
                if (!IsDoorLikeInstance(fi))
                    continue;

                try
                {
                    if (!fi.Pinned)
                    {
                        fi.Pinned = true;
                        changed.Add(fi.Id);
                    }
                }
                catch
                {
                    // Если Revit не даёт закрепить элемент (редкие случаи) — пропускаем
                }
            }

            return changed;
        }



        static bool TryReadLinkOverrideFields(object rlgs, out ElementId linkedViewId, out string linkVisMode)
        {
            linkedViewId = ElementId.InvalidElementId;
            linkVisMode = null;
            if (rlgs == null) return false;

            try
            {
                var t = rlgs.GetType();

                var piVis = t.GetProperty("LinkVisibilityType");
                if (piVis != null)
                {
                    var v = piVis.GetValue(rlgs, null);
                    if (v != null) linkVisMode = v.ToString();
                }

                var piLid = t.GetProperty("LinkedViewId");
                if (piLid != null)
                {
                    var obj = piLid.GetValue(rlgs, null);
                    if (obj is ElementId eid) linkedViewId = eid;
                }
                else
                {
                    var miLid = t.GetMethod("GetLinkedViewId", Type.EmptyTypes);
                    if (miLid != null)
                    {
                        var obj = miLid.Invoke(rlgs, null);
                        if (obj is ElementId eid2) linkedViewId = eid2;
                    }
                }
                return true;
            }
            catch { return false; }
        }

        // Проба по instanceId и typeId — что сработает
        static bool TryGetLinkOverridesForInstanceOrType(View hostView, RevitLinkInstance link,
            out ElementId linkedViewId, out string linkVisMode)
        {
            linkedViewId = ElementId.InvalidElementId;
            linkVisMode = null;

            // 1) instanceId
            if (TryGetLinkOverridesViaReflection(hostView, link.Id, out var rlgs1) &&
                TryReadLinkOverrideFields(rlgs1, out linkedViewId, out linkVisMode))
                return true;

            // 2) typeId (часто требуется именно он)
            try
            {
                var typeId = link.GetTypeId();
                if (typeId != null &&
                    TryGetLinkOverridesViaReflection(hostView, typeId, out var rlgs2) &&
                    TryReadLinkOverrideFields(rlgs2, out linkedViewId, out linkVisMode))
                    return true;
            }
            catch { /* ignore */ }

            return false;
        }

        // 2024+: видимые в ВИДЕ ХОСТА элементы из Revit-ссылки (учитывает By Host/By Link/Custom, стадии, фильтры и т.п.)
        // Возвращает элементы из Revit-ссылки, видимые на hostView ровно так,
        // как их видит человек (учитывая "Revit связи → Переопределение видимости/графики").
        // 1) Revit 2024+: используем перегрузку FEC(doc, viewId, linkId) через рефлексию (если доступна).
        // 2) <2024: читаем View.GetLinkOverrides(linkId) и действуем по LinkedViewId/LinkVisibilityType.
        static IEnumerable<Element> GetVisibleFromLinkSmart(
    Document hostDoc, View hostView, RevitLinkInstance link, ElementFilter filter)
        {
            if (hostDoc == null || hostView == null || link == null)
                return Enumerable.Empty<Element>();

            // Лог настроек override'ов (как было)
            try
            {
                if (TryGetLinkOverridesForInstanceOrType(hostView, link, out var linkedViewId, out var linkVisMode))
                    Debug.WriteLine($"[LinkOverrides] mode={linkVisMode}, linkedViewId={linkedViewId.IntegerValue}");
                else
                    Debug.WriteLine("[LinkOverrides] GetLinkOverrides недоступен в этой версии API или overrides не заданы.");
            }
            catch { /* ignore */ }

            // 1) Пытаемся создать FEC(doc, viewId, linkId) через рефлексию (Revit 2024+)
            try
            {
                var ctor = typeof(FilteredElementCollector).GetConstructor(
                    new[] { typeof(Document), typeof(ElementId), typeof(ElementId) });

                if (ctor != null)
                {
                    var fecObj = (FilteredElementCollector)ctor.Invoke(
                        new object[] { hostDoc, hostView.Id, link.Id });

                    return fecObj
                        .WherePasses(filter)
                        .WhereElementIsNotElementType()
                        .ToElements();
                }
            }
            catch (TargetInvocationException tie) when (tie.InnerException != null)
            {
                Debug.WriteLine("[FEC(doc,view,link)] failed: " + tie.InnerException.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[FEC(doc,view,link)] failed: " + ex.Message);
            }

            // 2) Фоллбэк: разрешаем "какой linked view показывать" и обходим по нему (универсально для 20xx)
            var ldoc = link.GetLinkDocument();
            if (ldoc != null)
            {
                var linkedView = TryGetLinkedViewForHost(hostView, link, out var how);
                if (linkedView != null && IsViewOK(ldoc, linkedView))
                {
                    Debug.WriteLine("[LinkElements] fallback via linked view: " + how);
                    return new FilteredElementCollector(ldoc, linkedView.Id)
                        .WherePasses(filter)
                        .WhereElementIsNotElementType()
                        .ToElements();
                }
            }

            // 3) Последний шанс: берём всё из документа ссылки (хуже точность, но безопасно)
            if (ldoc != null)
            {
                Debug.WriteLine("[LinkElements] last fallback: whole link doc");
                return new FilteredElementCollector(ldoc)
                    .WherePasses(filter)
                    .WhereElementIsNotElementType()
                    .ToElements();
            }

            return Enumerable.Empty<Element>();
        }




        // ---------- «строительные» категории ----------
        static readonly BuiltInCategory[] BASE_ALLOWED_BIC = new[]
        {
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_Ceilings,
            BuiltInCategory.OST_Roofs,
            BuiltInCategory.OST_Columns,
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_StructuralFoundation,
            BuiltInCategory.OST_Stairs,
            BuiltInCategory.OST_Ramps,
            BuiltInCategory.OST_StairsRailing,
            BuiltInCategory.OST_Doors,
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_CurtainWallMullions,
            BuiltInCategory.OST_CurtainWallPanels,
            BuiltInCategory.OST_GenericModel,
            BuiltInCategory.OST_ElectricalEquipment,
            BuiltInCategory.OST_MechanicalEquipment
        };

        // --- чёрный список ---
        static readonly HashSet<BuiltInCategory> DENY_BIC = new HashSet<BuiltInCategory>
        {
            BuiltInCategory.OST_Lines,
            BuiltInCategory.OST_SketchLines,
            BuiltInCategory.OST_RoomSeparationLines,
            BuiltInCategory.OST_Grids,
            BuiltInCategory.OST_Levels,
            BuiltInCategory.OST_Topography,
            BuiltInCategory.OST_PointClouds,
            BuiltInCategory.OST_IOS_GeoSite,
        };

        // ---- вычисляется при запуске ----
        static double sMinSegFt;
        static double sKeySnapFt;
        // ---- счётчики вырезанных элементов (поштучно) ----
        static HashSet<string> sCutElementUids; // чтобы один и тот же элемент не считался много раз
        static int sCutElementCount;
        // ---- счётчики дверей (для отладки «зашивки» дверей) ----
        int sDoorLikeSeenHost;     // сколько дверей найдено в хост-модели
        int sDoorLikeSeenLinks;    // сколько дверей найдено в ссылках
        int sDoorLikeNoLocation;   // сколько дверей без LocationPoint
        int sDoorLikeBridged;      // по скольким реально добавлен bridge-сегмент



        // ---- текущий выбор «шахт» (меняется через диалог) ----
        static bool sCut_Enable;                     // «Вырезать шахты: Да/Нет»
        static bool sCut_SystemShafts;               // включать системные шахты (OST_ShaftOpening)
        static bool sCut_Custom;                     // включать пользовательские (категории/семейства)
        static bool sCut_IncludeLinks = true;        // «Учитывать в ссылках»
        static HashSet<BuiltInCategory> sCut_Cats = new HashSet<BuiltInCategory>();
        static HashSet<string> sCut_Fams = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // ---- режим расчёта: ГНС или Общая площадь ----
        // true  = ГНС (как сейчас)
        // false = Общая площадь (OPA)
        static bool sModeGNS = true;
        static bool sModeOPA = false;
        // ---- OPA: из какого параметра читать признак "МОП" ----
        // Пусто/не выбрано => fallback на старую логику (или можно сделать "не МОП")
        static string sOpaMopParamName = "";


        // ---- отладочный лог, выводимый в TaskDialog и в *.debug.txt ----
        static readonly StringBuilder sDebug = new StringBuilder(8192);

        static void DebugAdd(string msg)
        {
            try
            {
                sDebug.AppendLine(msg);
            }
            catch
            {
                // чтобы логгер не мог уронить команду
            }
        }

        // ---------- утилиты ----------
        static double RoundHalfUp(double v, int nd) => Math.Round(v, nd, MidpointRounding.AwayFromZero);
        static bool Nearly(double a, double b, double tol = 1e-9) => Math.Abs(a - b) < tol;
        // Надёжное чтение комментария помещения (учитывает русскую локализацию)
        // Универсальное чтение комментария помещения (работает и в ru, и в en Revit)
        // Универсальное чтение комментария помещения (работает во всех локализациях и версиях API)
        static string GetRoomComment(Element room)
        {
            if (room == null) return "";

            // 1) ищем параметр по Id — наиболее надёжно
            try
            {
                foreach (Parameter p in room.Parameters)
                {
                    if (p.Definition == null) continue;

                    var defName = p.Definition.Name.Trim().ToLowerInvariant();

                    // русская и английская локализации
                    if (defName == "комментарии" || defName == "comments" || defName == "comment")
                    {
                        if (p.StorageType == StorageType.String && p.HasValue)
                        {
                            var s = p.AsString();
                            if (!string.IsNullOrEmpty(s))
                                return s;
                        }
                    }
                }
            }
            catch { }

            // 2) на всякий случай — LookupParameter
            string[] names = { "Комментарии", "Comments", "Comment" };
            foreach (var n in names)
            {
                try
                {
                    var p = room.LookupParameter(n);
                    if (p != null &&
                        p.StorageType == StorageType.String &&
                        p.HasValue)
                    {
                        var s = p.AsString();
                        if (!string.IsNullOrEmpty(s))
                            return s;
                    }
                }
                catch { }
            }

            return "";
        }

        // Чтение текстового параметра "OMDV_назначение" / "OMDV_Purpose" и т.п.
        // Логика такая же, как в GetRoomComment: сначала перебираем Parameters,
        // потом пробуем LookupParameter.
        static string GetRoomPurpose(Element room)
        {
            if (room == null) return "";

            // 1) поиск по Definition.Name (как и в GetRoomComment)
            try
            {
                foreach (Parameter p in room.Parameters)
                {
                    if (p.Definition == null) continue;

                    var defName = p.Definition.Name.Trim();

                    // возможные имена параметра в русской/английской версии шаблона
                    if (defName.Equals("OMDV_назначение", StringComparison.OrdinalIgnoreCase) ||
                        defName.Equals("OMDV_Purpose", StringComparison.OrdinalIgnoreCase))
                    {
                        if (p.StorageType == StorageType.String && p.HasValue)
                        {
                            var s = p.AsString();
                            if (!string.IsNullOrEmpty(s))
                                return s;
                        }
                    }
                }
            }
            catch { }

            // 2) на всякий случай — LookupParameter по тем же именам
            string[] names = { "OMDV_назначение", "OMDV_Purpose" };
            foreach (var n in names)
            {
                try
                {
                    var p = room.LookupParameter(n);
                    if (p != null &&
                        p.StorageType == StorageType.String &&
                        p.HasValue)
                    {
                        var s = p.AsString();
                        if (!string.IsNullOrEmpty(s))
                            return s;
                    }
                }
                catch { }
            }

            return "";
        }

        // Собирает комментарий для JSON из:
        // - стандартного "Комментарии"
        // - OMDV_Назначение
        // - "Назначение" (Occupancy)
        //
        // Если где-то стоит МОП/MOP, гарантированно пишет "МОП".
        static string BuildRoomCommentForJson(Element room)
        {
            if (room == null) return "";

            // 1) то, что уже умеет читать существующий метод
            string baseComment = GetRoomComment(room);
            string omdvPurpose = GetRoomPurpose(room);
            string assign = GetRoomAssignment(room);

            // 2) OPA: МОП читаем из выбранного параметра Room
            // (только свойства помещения, без элементов внутри)
            if (IsModeOPA() && !string.IsNullOrWhiteSpace(sOpaMopParamName))
            {
                string mopSrc = GetRoomParamTextByName(room, sOpaMopParamName);
                if (IsMopMark(mopSrc))
                    return "МОП";
            }
            else
            {
                // старое поведение для GNS (или если параметр не выбран)
                if (IsMopMark(baseComment) || IsMopMark(omdvPurpose) || IsMopMark(assign))
                    return "МОП";
            }



            // 3) иначе аккуратно собираем текст
            var parts = new List<string>();

            if (!string.IsNullOrWhiteSpace(baseComment))
                parts.Add(baseComment.Trim());

            if (!string.IsNullOrWhiteSpace(omdvPurpose))
                parts.Add(omdvPurpose.Trim());

            if (!string.IsNullOrWhiteSpace(assign))
                parts.Add(assign.Trim());

            if (parts.Count == 0)
                return "";

            return string.Join("; ", parts);
        }

        static bool IsMopMark(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;

            string s = value.Trim();

            // быстрый кейс: ровно "МОП"
            if (s.Equals("моп", StringComparison.OrdinalIgnoreCase) ||
                s.Equals("mop", StringComparison.OrdinalIgnoreCase))
                return true;

            // токены по разделителям
            var token = new StringBuilder(16);
            for (int i = 0; i < s.Length; i++)
            {
                char ch = char.ToLowerInvariant(s[i]);

                bool isLetterOrDigit = char.IsLetterOrDigit(ch) || ch == '_';
                if (isLetterOrDigit)
                {
                    token.Append(ch);
                }
                else
                {
                    if (token.Length > 0)
                    {
                        var t = token.ToString();
                        if (t == "моп" || t == "mop") return true;
                        token.Clear();
                    }
                }
            }

            if (token.Length > 0)
            {
                var t = token.ToString();
                if (t == "моп" || t == "mop") return true;
            }

            return false;
        }

        // Чтение текстового параметра "Назначение" / "Purpose" / "Assignment"
        // (обычно это параметр помещения "Назначение", который виден в свойствах Room)
        static string GetRoomAssignment(Element room)
        {
            if (room == null) return "";

            // 1) поиск по Definition.Name
            try
            {
                foreach (Parameter p in room.Parameters)
                {
                    if (p.Definition == null) continue;

                    var defName = p.Definition.Name.Trim().ToLowerInvariant();

                    // возможные варианты имени параметра
                    if (defName == "назначение" ||
                        defName == "purpose" ||
                        defName == "assignment")
                    {
                        if (p.StorageType == StorageType.String && p.HasValue)
                        {
                            var s = p.AsString();
                            if (!string.IsNullOrEmpty(s))
                                return s;
                        }
                    }
                }
            }
            catch { }

            // 2) на всякий случай — LookupParameter
            string[] names = { "Назначение", "Purpose", "Assignment" };
            foreach (var n in names)
            {
                try
                {
                    var p = room.LookupParameter(n);
                    if (p != null &&
                        p.StorageType == StorageType.String &&
                        p.HasValue)
                    {
                        var s = p.AsString();
                        if (!string.IsNullOrEmpty(s))
                            return s;
                    }
                }
                catch { }
            }

            return "";
        }
        // Возвращает строковое значение параметра Room по имени (Definition.Name / LookupParameter).
        static string GetRoomParamTextByName(Element room, string paramName)
        {
            if (room == null) return "";
            if (string.IsNullOrWhiteSpace(paramName)) return "";

            // 1) быстрый путь
            try
            {
                var p = room.LookupParameter(paramName);
                if (p != null && p.StorageType == StorageType.String)
                    return p.AsString() ?? "";
            }
            catch { }

            // 2) универсально — перебор room.Parameters (это именно свойства помещения)
            try
            {
                foreach (Parameter p in room.Parameters)
                {
                    if (p?.Definition == null) continue;
                    if (!string.Equals(p.Definition.Name?.Trim(), paramName.Trim(), StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (p.StorageType == StorageType.String)
                        return p.AsString() ?? "";

                    return "";
                }
            }
            catch { }

            return "";
        }

        // Универсально читает значение параметра помещения по имени (Definition.Name / LookupParameter).
        // Для String вернёт AsString(), для остальных — AsValueString() (если есть).
        static string GetRoomParamValueByName(Element room, string paramName)
        {
            if (room == null) return "";
            if (string.IsNullOrWhiteSpace(paramName)) return "";

            string want = paramName.Trim();

            // 1) Надёжнее: перебор room.Parameters
            try
            {
                foreach (Parameter p in room.Parameters)
                {
                    if (p == null || p.Definition == null) continue;

                    var name = (p.Definition.Name ?? "").Trim();
                    if (!name.Equals(want, StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (!p.HasValue) return "";

                    if (p.StorageType == StorageType.String)
                        return p.AsString() ?? "";

                    // для чисел/да-нет/элементов — обычно AsValueString() даёт то, что видит пользователь
                    var vs = p.AsValueString();
                    if (!string.IsNullOrEmpty(vs)) return vs;

                    // запасной путь (иногда у некоторых параметров AsValueString пустой)
                    try { return p.AsString() ?? ""; } catch { return ""; }
                }
            }
            catch { }

            // 2) Fallback: LookupParameter (на случай, если перебор не нашёл)
            try
            {
                var p2 = room.LookupParameter(want);
                if (p2 != null && p2.HasValue)
                {
                    if (p2.StorageType == StorageType.String)
                        return p2.AsString() ?? "";

                    return p2.AsValueString() ?? "";
                }
            }
            catch { }

            return "";
        }

        // OPA: признак МОП читаем из выбранного пользователем параметра.
        // Если параметр не выбран/не найден — fallback на старую проверку (как было).
        static bool IsRoomMop(Element room)
        {
            // 1) Новый путь: выбранный параметр
            if (IsModeOPA() && !string.IsNullOrWhiteSpace(sOpaMopParamName))
            {
                var v = GetRoomParamValueByName(room, sOpaMopParamName);
                return IsMopMark(v);
            }

            // 2) Fallback: старая логика (как было)
            string baseComment = GetRoomComment(room);
            string omdvPurpose = GetRoomPurpose(room);
            string assign = GetRoomAssignment(room);

            return IsMopMark(baseComment) || IsMopMark(omdvPurpose) || IsMopMark(assign);
        }


        static void CountCutElementOnce(Element el)
        {
            if (el == null) return;
            try
            {
                var uid = el.UniqueId; // устойчивый идентификатор
                if (string.IsNullOrEmpty(uid)) return;
                if (sCutElementUids.Add(uid))
                    sCutElementCount++;
            }
            catch { /* молча */ }
        }

        static string AskJsonPath()
        {
            var dlg = new FileSaveDialog("JSON files (*.json)|*.json");
            dlg.Title = "Сохранить экспорт в JSON";
            dlg.InitialFileName = "export_sheet.json";
            if (dlg.Show() != ItemSelectionDialogResult.Confirmed) return null;
            return ModelPathUtils.ConvertModelPathToUserVisiblePath(dlg.GetSelectedModelPath());
        }

        static (ForgeTypeId uId, string name) GetTargetLengthUnit(Document doc)
        {
            if (FORCE_MM) return (UnitTypeId.Millimeters, "millimeters");
            var fmt = doc.GetUnits().GetFormatOptions(SpecTypeId.Length);
            var uId = fmt.GetUnitTypeId();
            return (uId, uId?.TypeId ?? "length-project-units");
        }

        static XYZ SnapXY(XYZ p)
        {
            if (sKeySnapFt <= 0) return new XYZ(p.X, p.Y, 0);
            return new XYZ(
                Math.Round(p.X / sKeySnapFt) * sKeySnapFt,
                Math.Round(p.Y / sKeySnapFt) * sKeySnapFt,
                0);
        }
        static string KeyFor(XYZ a, XYZ b)
        {
            double ax = a.X, ay = a.Y, bx = b.X, by = b.Y;
            if (ax > bx || (Nearly(ax, bx) && ay > by))
            {
                var tx = ax; var ty = ay;
                ax = bx; ay = by;
                bx = tx; by = ty;
            }
            // используем инвариантную культуру и формат "R" как раньше
            var inv = CultureInfo.InvariantCulture;
            return ax.ToString("R", inv) + "|" + ay.ToString("R", inv) + "|" +
                   bx.ToString("R", inv) + "|" + by.ToString("R", inv);
        }

        // Возвращает true, если вид можно безопасно использовать в FilteredElementCollector(doc, view.Id)
        static bool IsViewOK(Document d, View v)
        {
            if (v == null) return false;

            // Шаблон вида нельзя
            try { if (v.IsTemplate) return false; } catch { /* старые версии API */ }

            // Основная проверка Revit API (Revit 2019+)
            try { return FilteredElementCollector.IsViewValidForElementIteration(d, v.Id); }
            catch
            {
                // Фоллбек на случай очень старых API: оставляем только план-виды
                return (v is ViewPlan) || v.ViewType == ViewType.CeilingPlan || v.ViewType == ViewType.AreaPlan;
            }
        }

        // Возвращает первый НЕ шаблонный план, пригодный для итерации
        static ViewPlan FirstUsablePlan(Document d)
        {
            foreach (var vp in new FilteredElementCollector(d)
                .OfClass(typeof(ViewPlan)).Cast<ViewPlan>())
            {
                try { if (vp.IsTemplate) continue; } catch { /* old API */ }
                try
                {
                    if (!FilteredElementCollector.IsViewValidForElementIteration(d, vp.Id))
                        continue;
                }
                catch { /* old API — OK */ }
                return vp;
            }
            return null;
        }

        // Безопасная проверка "элемент виден в этом виде",
        // работает и с новыми API (через View.IsElementVisibleInView), и со старыми (через IsHidden + Category.Visible)
        static bool IsElementVisibleInViewSafe(Document doc, View view, ElementId elemId)
        {
            if (doc == null || view == null || elemId == null || elemId == ElementId.InvalidElementId)
                return true; // ничего лучше сказать не можем — считаем видимым

            // 1) Пытаемся вызвать статический View.IsElementVisibleInView(Document, ElementId, ElementId)
            try
            {
                var mi = typeof(View).GetMethod(
                    "IsElementVisibleInView",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static,
                    null,
                    new[] { typeof(Document), typeof(ElementId), typeof(ElementId) },
                    null);

                if (mi != null)
                {
                    var res = mi.Invoke(null, new object[] { doc, view.Id, elemId });
                    if (res is bool b)
                        return b;
                }
            }
            catch
            {
                // тихо падаем в фоллбэк
            }

            // 2) Старые версии API — проверяем IsHidden и видимость категории
            Element el = null;
            try { el = doc.GetElement(elemId); } catch { }

            if (el == null)
                return true;

            // Если элемент скрыт в виде — он нам не нужен
            try
            {
                if (el.IsHidden(view))
                    return false;
            }
            catch { }

            // Если категория не видна в этом виде — тоже считаем невидимым
            try
            {
                var cat = el.Category;
                if (cat != null)
                {
                    if (!cat.get_Visible(view))
                        return false;
                }
            }
            catch { }

            // Если ничего не узнали — оставляем как видимый (поведение как "по старому")
            return true;
        }


        // ---------- фильтры и пропуски ----------
        static bool SkipAux(Element el)
        {
            if (el == null) return true;
            if (el is ImportInstance) return true;
            if (el is ReferencePlane) return true;
            if (el is SketchPlane) return true;
            var cat = el.Category; if (cat == null) return true;
            var bic = (BuiltInCategory)cat.Id.IntegerValue;
            if (DENY_BIC.Contains(bic)) return true;
            if (el.ViewSpecific) return true;
            return false;
        }

        static ElementFilter BuildAllowedFilterWithCutouts()
        {
            var ids = new HashSet<ElementId>(BASE_ALLOWED_BIC.Select(b => new ElementId(b)));

            if (sCut_Enable)
            {
                if (sCut_SystemShafts)
                    ids.Add(new ElementId(BuiltInCategory.OST_ShaftOpening));

                if (sCut_Custom && sCut_Cats.Count > 0)
                    foreach (var bic in sCut_Cats) ids.Add(new ElementId(bic));
            }

            return new ElementMulticategoryFilter(ids.ToList());
        }


        static bool ShouldGoToCutouts(Element el, bool fromLink)
        {
            if (!sCut_Enable) return false;
            if (fromLink && !sCut_IncludeLinks) return false;
            if (el == null) return false;

            // системные
            if (sCut_SystemShafts && IsSystemShaft(el))
                return true;

            // пользовательские
            if (!sCut_Custom) return false;          // пользовательские отключены

            var cat = el.Category; if (cat == null) return false;
            var bic = (BuiltInCategory)cat.Id.IntegerValue;

            bool catOk = sCut_Cats.Count == 0 || sCut_Cats.Contains(bic);
            if (!catOk) return false;

            if (sCut_Fams.Count == 0) return true;

            if (el is FamilyInstance fi)
            {
                var fam = fi.Symbol?.Family?.Name ?? fi.Name;
                return fam != null && sCut_Fams.Contains(fam);
            }
            return false;
        }

        // true, если элемент относится к витражу
        static bool IsCurtainElement(Element el)
        {
            if (el == null) return false;

            try
            {
                var cat = el.Category;
                if (cat != null)
                {
                    var bic = (BuiltInCategory)cat.Id.IntegerValue;
                    if (bic == BuiltInCategory.OST_CurtainWallMullions ||
                        bic == BuiltInCategory.OST_CurtainWallPanels)
                        return true;
                }

                if (el is Wall w)
                {
                    try
                    {
                        if (w.WallType != null && w.WallType.Kind == WallKind.Curtain)
                            return true;
                    }
                    catch { }
                }

                // --- ДОБАВЛЕНО: элементы сетки витража ---
                var typeName = el.GetType()?.Name;
                if (string.Equals(typeName, "CurtainGridLine", StringComparison.Ordinal) ||
                    string.Equals(typeName, "CurtainCell", StringComparison.Ordinal))
                    return true;
            }
            catch { }

            return false;
        }

        // Пытаемся по элементу границы найти витражную стену-хост
        // Пытаемся по элементу границы найти витражную стену-хост



        // Более надёжная проверка системной шахты
        static bool IsSystemShaft(Element el)
        {
            if (el == null) return false;
            try
            {
                var cat = el.Category;
                if (cat != null && (BuiltInCategory)cat.Id.IntegerValue == BuiltInCategory.OST_ShaftOpening)
                    return true;

                var tn = el.GetType()?.Name;
                if (string.Equals(tn, "ShaftOpening", StringComparison.Ordinal))
                    return true;
            }
            catch { }
            return false;
        }

        // Попытка получить кривые границы отверстия/шахты
        static bool TryGetOpeningBoundaryCurves(Element el, out IList<Curve> curves)
        {
            curves = null;
            try
            {
                if (el is Opening op)
                {
                    var ca = op.BoundaryCurves;
                    if (ca != null)
                    {
                        curves = ca.Cast<Curve>().ToList();
                        return curves.Count > 0;
                    }
                }

                var t = el.GetType();

                var pi = t.GetProperty("BoundaryCurves");
                if (pi != null)
                {
                    var val = pi.GetValue(el);
                    if (val is CurveArray ca2)
                    {
                        curves = ca2.Cast<Curve>().ToList();
                        return curves.Count > 0;
                    }
                    if (val is IList<Curve> list)
                    {
                        curves = list;
                        return curves.Count > 0;
                    }
                }

                var mi = t.GetMethod("GetBoundaryCurves", Type.EmptyTypes)
                      ?? t.GetMethod("GetBoundary", Type.EmptyTypes)
                      ?? t.GetMethod("GetProfiles", Type.EmptyTypes)
                      ?? t.GetMethod("GetBoundarySegments", Type.EmptyTypes);

                if (mi != null)
                {
                    var v = mi.Invoke(el, null);
                    if (v is CurveArray ca3)
                    {
                        curves = ca3.Cast<Curve>().ToList();
                        return curves.Count > 0;
                    }
                    if (v is IList<Curve> list2)
                    {
                        curves = list2;
                        return curves.Count > 0;
                    }
                    if (v is IList<CurveLoop> loops)
                    {
                        curves = loops.SelectMany(cl => cl).ToList();
                        return curves.Count > 0;
                    }
                }
            }
            catch { }
            return false;
        }

        // Сложить границу шахты в сегменты cutouts (в координатах ВИДА)
        static void CollectShaftOpeningSegments(
            HashSet<string> keys, List<Seg2> segs, Element el,
            Transform T_model_to_view, Transform extraTransform,
            ref int cntSegsAdded)
        {
            if (!TryGetOpeningBoundaryCurves(el, out var crvs) || crvs == null || crvs.Count == 0) return;
            var T_final = Compose(T_model_to_view, extraTransform, null);
            foreach (var c in crvs)
                cntSegsAdded += CurveToSegments2D(keys, segs, c, T_final);
        }

        // ---------- двери: бриджинг ----------
        // Удобное описание экземпляра двери для отладки
        static string DescribeDoorInstance(FamilyInstance fi)
        {
            if (fi == null) return "<null>";
            try
            {
                var docName = fi.Document?.Title ?? "<doc>";
                var catName = fi.Category?.Name ?? "<no cat>";
                int catInt = fi.Category?.Id.IntegerValue ?? 0;
                string bicStr;
                try { bicStr = ((BuiltInCategory)catInt).ToString(); }
                catch { bicStr = catInt.ToString(); }

                var famName = fi.Symbol?.Family?.Name ?? "<no fam>";
                var typeName = fi.Symbol?.Name ?? "<no type>";

                return $"doc='{docName}', id={fi.Id.IntegerValue}, bic={bicStr} ({catName}), fam='{famName}', type='{typeName}'";
            }
            catch
            {
                return "<door info error>";
            }
        }

        // ---------- двери: бриджинг ----------
        class DoorBridgeInfo
        {
            public XYZ Center;        // центр двери в координатах вида
            public XYZ T;             // направление вдоль стены
            public XYZ N;             // нормаль "наружу"
            public double HalfT;      // половина толщины стены
            public FamilyInstance Door; // сам экземпляр двери (для чтения ширины и т.п.)
            public bool IsCurtain;    // true = витражная дверь / витражный контекст
        }



        static bool TryNormalize2D(XYZ v, out XYZ vn)
        {
            double len2 = v.X * v.X + v.Y * v.Y; if (len2 < 1e-18) { vn = null; return false; }
            double k = 1.0 / Math.Sqrt(len2); vn = new XYZ(v.X * k, v.Y * k, 0); return true;
        }
        static double Cross2D(XYZ u, XYZ v) => u.X * v.Y - u.Y * v.X;
        // расстояние от точки p до отрезка AB в 2D (XY)

        static bool ClosestPointOnSegment2D(XYZ p, XYZ a, XYZ b, out XYZ q, out double t)
        {
            var ab = new XYZ(b.X - a.X, b.Y - a.Y, 0);
            double ab2 = ab.X * ab.X + ab.Y * ab.Y;
            if (ab2 < 1e-12)
            {
                q = a; t = 0; return false;
            }

            var ap = new XYZ(p.X - a.X, p.Y - a.Y, 0);
            t = (ap.X * ab.X + ap.Y * ab.Y) / ab2;
            if (t < 0) t = 0;
            else if (t > 1) t = 1;

            q = new XYZ(a.X + ab.X * t, a.Y + ab.Y * t, 0);
            return true;
        }

        static bool ClosestPointOnAnySeg2D(XYZ p, List<Seg2> segs, out XYZ qBest, out double dBest)
        {
            qBest = p;
            dBest = double.MaxValue;
            bool ok = false;

            for (int i = 0; i < segs.Count; ++i)
            {
                var s = segs[i];
                if (!ClosestPointOnSegment2D(p, s.A, s.B, out var q, out var t))
                    continue;

                double dx = p.X - q.X, dy = p.Y - q.Y;
                double d = Math.Sqrt(dx * dx + dy * dy);
                if (d < dBest)
                {
                    dBest = d;
                    qBest = q;
                    ok = true;
                }
            }
            return ok;
        }


        static double DistPointToSegment2D(XYZ p, XYZ a, XYZ b)
        {
            double vx = b.X - a.X;
            double vy = b.Y - a.Y;
            double wx = p.X - a.X;
            double wy = p.Y - a.Y;

            double c1 = vx * wx + vy * wy;
            if (c1 <= 0.0)
            {
                // ближе к A
                return Math.Sqrt(wx * wx + wy * wy);
            }

            double c2 = vx * vx + vy * vy;
            if (c2 <= c1)
            {
                // ближе к B
                double dx = p.X - b.X;
                double dy = p.Y - b.Y;
                return Math.Sqrt(dx * dx + dy * dy);
            }

            double t = c1 / c2;
            double projX = a.X + t * vx;
            double projY = a.Y + t * vy;

            double dxp = p.X - projX;
            double dyp = p.Y - projY;
            return Math.Sqrt(dxp * dxp + dyp * dyp);
        }
        /// <summary>
        /// Убирает «ступеньки» границы помещения, которые висят прямо на витраже.
        /// Работает только с НЕ-витражными сегментами, витраж (roomCurtainSegs) не трогает.
        /// Критерий: оба конца сегмента находятся ближе ROOM_CURTAIN_DIST_TOL_MM
        /// к какому-либо витражному сегменту.
        /// </summary>
        static List<Seg2> RemoveCurtainStepsFromWalls(List<Seg2> allSegs, List<Seg2> roomCurtainSegs)
        {
            if (allSegs == null || allSegs.Count == 0 ||
                roomCurtainSegs == null || roomCurtainSegs.Count == 0)
                return allSegs;

            // допуск расстояния от концов «ступеньки» до витража
            double maxDistFt = UnitUtils.ConvertToInternalUnits(
                ROOM_CURTAIN_DIST_TOL_MM, UnitTypeId.Millimeters);

            // ключи витражных сегментов — их никогда не удаляем
            var curtainKeys = new HashSet<string>();
            foreach (var cs in roomCurtainSegs)
                curtainKeys.Add(KeyForSeg(cs));

            var result = new List<Seg2>(allSegs.Count);

            foreach (var s in allSegs)
            {
                // 1) Сегменты самого витража вообще не трогаем
                if (curtainKeys.Contains(KeyForSeg(s)))
                {
                    result.Add(s);
                    continue;
                }

                // 2) Считаем минимальное расстояние от каждого конца до витража
                double bestDistA = double.PositiveInfinity;
                double bestDistB = double.PositiveInfinity;

                foreach (var cs in roomCurtainSegs)
                {
                    double da = DistPointToSegment2D(s.A, cs.A, cs.B);
                    if (da < bestDistA) bestDistA = da;

                    double db = DistPointToSegment2D(s.B, cs.A, cs.B);
                    if (db < bestDistB) bestDistB = db;
                }

                // 3) Оба конца «сидят» почти на линии витража → считаем артефактом и выкидываем
                if (bestDistA < maxDistFt && bestDistB < maxDistFt)
                {
                    // это как раз твоя диагональная ступенька 100 мм
                    continue;
                }

                // 4) Всё остальное оставляем как есть
                result.Add(s);
            }

            return result;
        }


        // Определяем, что экземпляр "двереподобный"
        // Определяем, что экземпляр "двереподобный"
        static bool IsDoorLikeInstance(FamilyInstance fi)
        {
            if (fi == null) return false;
            Debug.WriteLine("[DoorLike?] " + DescribeDoorInstance(fi));
            // Имена для анализа (в т.ч. «Полуторная витражная дверь» и подобные)
            var famName = fi.Symbol?.Family?.Name ?? string.Empty;
            var typeName = fi.Symbol?.Name ?? string.Empty;
            var fullName = (famName + " " + typeName).Trim();

            if (!string.IsNullOrEmpty(fullName))
            {
                // Явная проверка на витражные двери
                if (fullName.IndexOf("полуторная витражная дверь", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    fullName.IndexOf("витражная дверь", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Debug.WriteLine("[DoorBridge] treat as curtain-door by name: " + fullName);
                    return true;
                }
            }

            // 1) Категория самого экземпляра
            var instCatId = fi.Category?.Id.IntegerValue ?? -1;
            if (instCatId == (int)BuiltInCategory.OST_Doors)
                return true;

            // 2) Категория семейства
            var famCatId = fi.Symbol?.Family?.FamilyCategory?.Id.IntegerValue ?? -1;
            if (famCatId == (int)BuiltInCategory.OST_Doors)
                return true;

            // 3) Категория типа (на всякий случай)
            var symCatId = fi.Symbol?.Category?.Id.IntegerValue ?? -1;
            if (symCatId == (int)BuiltInCategory.OST_Doors)
                return true;

            // 4) Подстраховка по имени (для витражных панелей-дверей и прочих изощрений)
            var name = (fi.Symbol?.Family?.Name + " " + fi.Name) ?? string.Empty;
            if (name.IndexOf("door", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            if (name.IndexOf("двер", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            return false;
        }


        // Определяет, относится ли дверь к витражному контексту
        static bool IsCurtainDoorContext(FamilyInstance fi)
        {
            if (fi == null) return false;

            // 1) Дверь стоит в витражной стене
            if (fi.Host is Wall w)
            {
                try
                {
                    if (w.WallType.Kind == WallKind.Curtain)
                        return true;
                }
                catch { }
            }

            // 2) Имя семейства содержит "витраж"
            var famName = fi.Symbol?.Family?.Name ?? "";
            var typeName = fi.Symbol?.Name ?? "";
            var nm = (famName + " " + typeName).ToLowerInvariant();

            if (nm.Contains("витраж"))
                return true;

            // Иначе — обычная дверь
            return false;
        }

        // Попытка получить реальную ширину двери (в футах)
        static bool TryGetDoorWidthFt(FamilyInstance fi, out double widthFt)
        {
            widthFt = 0.0;
            if (fi == null) return false;

            Parameter p = null;

            // 1) Стандартный параметр DOOR_WIDTH
            try
            {
                p = fi.get_Parameter(BuiltInParameter.DOOR_WIDTH)
                    ?? fi.Symbol?.get_Parameter(BuiltInParameter.DOOR_WIDTH);
            }
            catch { }

            if (p != null && p.StorageType == StorageType.Double && p.HasValue)
            {
                widthFt = p.AsDouble();
                if (widthFt > 1e-6) return true;
            }

            // 2) Популярные имена параметров ("Width", "Ширина", "B")
            string[] names = { "Width", "Ширина", "B" };
            foreach (var n in names)
            {
                try
                {
                    p = fi.LookupParameter(n) ?? fi.Symbol?.LookupParameter(n);
                }
                catch { p = null; }

                if (p != null && p.StorageType == StorageType.Double && p.HasValue)
                {
                    widthFt = p.AsDouble();
                    if (widthFt > 1e-6) return true;
                }
            }

            // 3) Фоллбэк: по bounding box двери вдоль локальной оси X
            try
            {
                var bb = fi.get_BoundingBox(null);
                var tr = fi.GetTransform();
                if (bb != null && tr != null)
                {
                    var inv = tr.Inverse;

                    XYZ[] pts =
                    {
                        bb.Min,
                        new XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
                        new XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
                        new XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
                        new XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z),
                        new XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
                        new XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
                        bb.Max
                    };

                    double min = double.PositiveInfinity;
                    double max = double.NegativeInfinity;

                    foreach (var p3 in pts)
                    {
                        var lp = inv.OfPoint(p3); // координаты в системе семейства
                        double s = lp.X;          // вдоль локальной оси X (ширина проёма)
                        if (s < min) min = s;
                        if (s > max) max = s;
                    }

                    widthFt = max - min;
                    if (widthFt > 1e-6) return true;
                }
            }
            catch { }

            return false;
        }


        // Центр и ориентация двери (в т.ч. витражной) в координатах модели документа
        // Центр и ориентация двери (в т.ч. витражной) в координатах модели документа
        static bool TryGetDoorCenterAndOrientation(
            FamilyInstance fi,
            Document doc,
            View viewOrNull,         // может быть null (для элементов в ссылке)
            out XYZ centerM,
            out XYZ tM,
            out XYZ nM,
            out double halfThickness)
        {
            centerM = null;
            tM = null;
            nM = null;
            halfThickness = 0.0;

            if (fi == null || doc == null)
                return false;

            // заранее определяем, витражный ли контекст
            bool curtainContext = IsCurtainDoorContext(fi);

            // --- 1. Базовый центр двери (как раньше) ---
            if (fi.Location is LocationPoint lp)
            {
                centerM = lp.Point;
            }
            else if (fi.Location is LocationCurve lc && lc.Curve != null)
            {
                centerM = lc.Curve.Evaluate(0.5, true);
            }
            else
            {
                BoundingBoxXYZ bb = null;
                try { bb = fi.get_BoundingBox(viewOrNull); } catch { }
                if (bb == null)
                    try { bb = fi.get_BoundingBox(null); } catch { }

                if (bb != null)
                    centerM = 0.5 * (bb.Min + bb.Max);
            }

            if (centerM == null)
                return false;

            // Вид, в котором будем искать ближайшую стену
            View wallSearchView = viewOrNull ?? FirstUsablePlan(doc);

            // Хост-стена (для обычных и витражных дверей)
            Wall hostWall = fi.Host as Wall;
            if (hostWall == null && wallSearchView != null)
                hostWall = FindNearestWallForDoor(doc, wallSearchView, centerM);
            if (hostWall == null)
                hostWall = FindNearestWallForDoor(doc, null, centerM);

            // ===== ВАРИАНТ 3: специальная обработка витражных дверей =====
            // Для витража мы всегда доверяем геометрии стены, а не FacingOrientation
            if (curtainContext && hostWall != null)
            {
                // 1) проецируем центр на линию стены, чтобы попасть ровно на витраж
                try
                {
                    if (hostWall.Location is LocationCurve wlc && wlc.Curve != null)
                    {
                        var pr = wlc.Curve.Project(centerM);
                        if (pr != null)
                            centerM = pr.XYZPoint;
                    }
                }
                catch { /* безопасный фоллбек, оставляем исходный centerM */ }

                // 2) берём T и N только из стены
                if (!GetWallDirsAtPoint(hostWall, centerM, out tM, out nM))
                    return false;

                halfThickness = 0.5 * hostWall.Width;
                centerM = new XYZ(centerM.X, centerM.Y, 0.0);
                return true;
            }

            // ===== Дальше — старая логика для Обычных дверей =====

            // --- 2. Нормаль к проёму по FacingOrientation ---
            XYZ n = null;
            try
            {
                var fn = fi.FacingOrientation;
                if (fn != null)
                {
                    // если ориентация почти вертикальная → не подходит
                    if (Math.Abs(fn.Z) > 0.7)
                        n = null;
                    else
                        n = new XYZ(fn.X, fn.Y, 0.0);
                }
            }
            catch
            {
                n = null;
            }

            // если FacingOrientation не даёт вменяемый вектор → пробуем стену-хост
            if (n == null || (n.X * n.X + n.Y * n.Y) < 1e-12)
            {
                if (hostWall == null)
                    return false;

                if (!GetWallDirsAtPoint(hostWall, centerM, out tM, out nM))
                    return false;

                halfThickness = 0.5 * hostWall.Width;
                centerM = new XYZ(centerM.X, centerM.Y, 0.0);
                return true;
            }

            // --- 3. Обычный путь: нормаль из двери (FacingOrientation) ---
            if (!TryNormalize2D(n, out var nXY))
                return false;

            var tXY = new XYZ(-nXY.Y, nXY.X, 0.0);

            if (hostWall != null)
                halfThickness = 0.5 * hostWall.Width;
            else
                // запасной вариант: толщина 100 мм
                halfThickness = 0.5 * UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);

            centerM = new XYZ(centerM.X, centerM.Y, 0.0);
            tM = tXY;
            nM = nXY;
            return true;
        }



        static (bool ok, XYZ pNeg, XYZ pPos) NearestHitsAlongLine(XYZ q0, XYZ T, List<Seg2> segs)
        {
            if (T.X * T.X + T.Y * T.Y < 1e-18) return (false, null, null);
            double bestPos = double.PositiveInfinity, bestNeg = double.NegativeInfinity;
            XYZ hitPos = null, hitNeg = null;

            foreach (var seg in segs)
            {
                var a = new XYZ(seg.A.X, seg.A.Y, 0);
                var b = new XYZ(seg.B.X, seg.B.Y, 0);
                var S = b - a;
                double den = Cross2D(T, S);
                if (Math.Abs(den) < 1e-12) continue;

                var w = a - q0;
                double lambda = Cross2D(w, S) / den;
                double mu = Cross2D(w, T) / den;

                if (mu < -1e-9 || mu > 1 + 1e-9) continue;

                if (lambda > 1e-9 && lambda < bestPos) { bestPos = lambda; hitPos = q0 + T.Multiply(lambda); }
                if (lambda < -1e-9 && lambda > bestNeg) { bestNeg = lambda; hitNeg = q0 + T.Multiply(lambda); }
            }
            return (hitPos != null && hitNeg != null, hitNeg, hitPos);
        }
        // --- «магнит» центра двери к ближайшей линии стены/витража по нормали N ---
        static bool TryProjectDoorCenterToNearestSeg(DoorBridgeInfo d, List<Seg2> segs, out XYZ correctedCenter)
        {
            correctedCenter = d.Center;
            double bestAbsDist = double.PositiveInfinity;
            bool found = false;

            foreach (var s in segs)
            {
                // направление сегмента в плоскости XY
                var dir = new XYZ(s.B.X - s.A.X, s.B.Y - s.A.Y, 0);
                double len2 = dir.X * dir.X + dir.Y * dir.Y;
                if (len2 < 1e-18) continue;

                double len = Math.Sqrt(len2);
                var dirN = new XYZ(dir.X / len, dir.Y / len, 0);

                // интересуют только сегменты, почти параллельные оси двери T
                double dot = Math.Abs(dirN.X * d.T.X + dirN.Y * d.T.Y);
                if (dot < 0.9) // 0.9 ~ 25°
                    continue;

                // смещение от центра двери до сегмента вдоль нормали N
                var v = new XYZ(s.A.X - d.Center.X, s.A.Y - d.Center.Y, 0);
                double dist = v.X * d.N.X + v.Y * d.N.Y; // signed distance по нормали

                double abs = Math.Abs(dist);
                if (abs < bestAbsDist)
                {
                    bestAbsDist = abs;
                    correctedCenter = new XYZ(
                        d.Center.X + d.N.X * dist,
                        d.Center.Y + d.N.Y * dist,
                        0);
                    found = true;
                }
            }

            return found;
        }
        // Пытаемся считать положение граней витражной панели по уже собранным сегментам.
        // Возвращает два оффсета от центра двери вдоль нормали N:
        // innerOffset < 0  — внутренняя грань панели
        // outerOffset > 0  — наружная грань панели
        static bool TryGetCurtainPanelOffsetsFromSegments(
            DoorBridgeInfo d,
            List<Seg2> segs,
            out double innerOffset,
            out double outerOffset)
        {
            innerOffset = 0.0;
            outerOffset = 0.0;

            if (segs == null || segs.Count == 0)
                return false;

            // пороги в футах
            double maxAlongT = UnitUtils.ConvertToInternalUnits(
                CURTAIN_SAMPLE_ALONG_T_MM, UnitTypeId.Millimeters);
            double maxAlongN = UnitUtils.ConvertToInternalUnits(
                CURTAIN_SAMPLE_MAX_N_MM, UnitTypeId.Millimeters);
            double minThickness = UnitUtils.ConvertToInternalUnits(
                CURTAIN_MIN_PANEL_THICKNESS_MM, UnitTypeId.Millimeters);

            var pos = new List<double>(); // оффсеты > 0 (с одной стороны витража)
            var neg = new List<double>(); // оффсеты < 0 (с другой стороны)

            foreach (var s in segs)
            {
                var a = s.A;
                var b = s.B;

                // направление сегмента в XY
                var dir = new XYZ(b.X - a.X, b.Y - a.Y, 0);
                double len2 = dir.X * dir.X + dir.Y * dir.Y;
                if (len2 < 1e-12) continue;

                double len = Math.Sqrt(len2);
                var dirN = new XYZ(dir.X / len, dir.Y / len, 0);

                // интересуют только сегменты, почти параллельные оси двери T
                double dotT = Math.Abs(dirN.X * d.T.X + dirN.Y * d.T.Y);
                if (dotT < 0.95) // > ~18°
                    continue;

                // вектор от центра двери до концов сегмента
                var va = a - d.Center;
                var vb = b - d.Center;

                // координаты вдоль оси двери (T)
                double ta = va.X * d.T.X + va.Y * d.T.Y;
                double tb = vb.X * d.T.X + vb.Y * d.T.Y;

                // сегмент полностью слишком далеко по T — не наш сосед
                if (ta > maxAlongT && tb > maxAlongT) continue;
                if (ta < -maxAlongT && tb < -maxAlongT) continue;

                // смещение вдоль нормали двери (N) — для параллельных сегментов оно почти константа
                double na = va.X * d.N.X + va.Y * d.N.Y;

                if (Math.Abs(na) > maxAlongN)
                    continue;

                if (na > 0) pos.Add(na);
                else neg.Add(na);
            }

            if (pos.Count == 0 || neg.Count == 0)
                return false;

            // берём самые "наружные" грани с каждой стороны
            outerOffset = pos.Max(); // > 0
            innerOffset = neg.Min(); // < 0

            // если толщина вышла подозрительно маленькой — лучше не использовать
            if (outerOffset - innerOffset < minThickness)
                return false;

            return true;
        }


        static bool GetWallDirsAtPoint(Wall wall, XYZ atPoint, out XYZ tXY, out XYZ nXY)
        {
            tXY = null; nXY = null;
            var lc = wall?.Location as LocationCurve; if (lc == null) return false;
            var crv = lc.Curve; var pr = crv.Project(atPoint); if (pr == null) return false;

            XYZ t = null;
            try { t = crv.ComputeDerivatives(pr.Parameter, true)?.BasisX; } catch { }
            if (t == null || (t.X * t.X + t.Y * t.Y + t.Z * t.Z) < 1e-18)
            {
                try { t = crv.GetEndPoint(1) - crv.GetEndPoint(0); } catch { }
            }
            if (t == null) return false;

            if (!TryNormalize2D(new XYZ(t.X, t.Y, 0), out tXY)) return false;
            nXY = new XYZ(-tXY.Y, tXY.X, 0);
            return true;
        }
        static Wall FindNearestWallForDoor(Document doc, View view, XYZ p, double searchRft = 30.0)
        {
            Wall best = null; double bestD = double.PositiveInfinity;
            foreach (var w in new FilteredElementCollector(doc, view.Id).OfClass(typeof(Wall)).Cast<Wall>())
            {
                var lc = w.Location as LocationCurve; if (lc == null) continue;
                var pr = lc.Curve.Project(p); if (pr == null) continue;
                var q = pr.XYZPoint;
                double d = Math.Sqrt((q.X - p.X) * (q.X - p.X) + (q.Y - p.Y) * (q.Y - p.Y));
                if (d < bestD) { bestD = d; best = w; }
            }
            return (best != null && bestD <= searchRft) ? best : null;
        }

        // ---------- трансформы ----------
        static Transform Compose(Transform tModelToView, Transform extra, Transform instT)
        {
            var I = Transform.Identity;
            var a = extra ?? I;
            var b = tModelToView ?? I;
            var c = instT ?? I;
            return b.Multiply(a).Multiply(c);
        }

        static Transform BuildViewToOutTransform(View v, Viewport vp, XYZ centerView, string mode, int unifiedScale)
        {
            if (v == null || vp == null) return Transform.Identity;
            double ang = 0.0;
            switch (vp.Rotation)
            {
                case ViewportRotation.None:
                    ang = 0.0;
                    break;
                case ViewportRotation.Clockwise:
                    ang = -Math.PI / 2.0;
                    break;
                case ViewportRotation.Counterclockwise:
                    ang = Math.PI / 2.0;
                    break;
                default:
                    ang = Math.PI;
                    break;
            }

            var Rvp = Transform.CreateRotation(XYZ.BasisZ, ang);
            var cSheet = vp.GetBoxCenter();             // бумажные ft
            int scaleDen = Math.Max(1, v.Scale);

            double k; XYZ origin;
            if (mode.Equals("paper", StringComparison.OrdinalIgnoreCase))
            {
                k = 1.0 / scaleDen; origin = cSheet - Rvp.OfVector(centerView.Multiply(k));
            }
            else if (mode.Equals("unified", StringComparison.OrdinalIgnoreCase))
            {
                int den = Math.Max(1, unifiedScale);
                k = 1.0 / den; origin = cSheet - Rvp.OfVector(centerView.Multiply(k));
            }
            else // "model"
            {
                // размеры оставляем в модельных (k=1),
                // но раскладываем виды как на листе: смещение листа переводим в модель через Scale
                k = 1.0;

                double den = Math.Max(1, v.Scale);
                XYZ cSheetModel = cSheet.Multiply(den); // paper ft -> model ft

                // хотим, чтобы centerView попал в "центр вьюпорта на листе", но уже в модельных единицах
                origin = cSheetModel - Rvp.OfVector(centerView);
            }


            var T = Transform.Identity;
            T.BasisX = Rvp.OfVector(new XYZ(k, 0, 0));
            T.BasisY = Rvp.OfVector(new XYZ(0, k, 0));
            T.BasisZ = XYZ.BasisZ;
            T.Origin = origin;
            return T;
        }

        static XYZ GetViewportCenterInViewCoords(View v)
        {
            try
            {
                var miGetMgr = typeof(View).GetMethod("GetAnnotationCropRegionShapeManager", Type.EmptyTypes);
                var mgrObj = miGetMgr?.Invoke(v, null);

                var piAnnoActive = typeof(View).GetProperty("AnnotationCropActive");
                bool annoActive = (piAnnoActive != null) && (bool)(piAnnoActive.GetValue(v, null) ?? false);

                if (annoActive && mgrObj != null)
                {
                    var tMgr = mgrObj.GetType();
                    var miShape =
                        tMgr.GetMethod("GetRegionShape", Type.EmptyTypes) ??
                        tMgr.GetMethod("GetAnnotationCropShape", Type.EmptyTypes) ??
                        tMgr.GetMethod("GetCropRegionShape", Type.EmptyTypes);

                    var loops = miShape?.Invoke(mgrObj, null) as System.Collections.Generic.IList<CurveLoop>;
                    if (loops != null && loops.Count > 0)
                    {
                        double minx = double.PositiveInfinity, miny = double.PositiveInfinity;
                        double maxx = double.NegativeInfinity, maxy = double.NegativeInfinity;

                        foreach (var cl in loops)
                        {
                            foreach (Curve c in cl)
                            {
                                var pts = c.Tessellate();
                                if (pts == null) continue;
                                foreach (var p in pts)
                                {
                                    if (p.X < minx) minx = p.X; if (p.Y < miny) miny = p.Y;
                                    if (p.X > maxx) maxx = p.X; if (p.Y > maxy) maxy = p.Y;
                                }
                            }
                        }
                        return new XYZ(0.5 * (minx + maxx), 0.5 * (miny + maxy), 0);
                    }
                }
            }
            catch { /* тихий фоллбэк */ }

            var cb = v?.CropBox;
            return (cb != null) ? 0.5 * (cb.Min + cb.Max) : XYZ.Zero;
        }

        // ---------- геометрия ----------
        static bool TryAddSeg2D(
            HashSet<string> keys,
            List<Seg2> segs,
            XYZ A,
            XYZ B,
            bool isCurtain = false,
            int curtainRootId = 0)
        {
            if (A == null || B == null) return false;
            if (A.DistanceTo(B) < sMinSegFt) return false;

            string k = KeyFor(A, B);
            if (!keys.Add(k)) return false;

            segs.Add(new Seg2(A, B, isCurtain, curtainRootId));
            return true;
        }





        static void AddSeg2D(
            HashSet<string> keys,
            List<Seg2> segs,
            XYZ a,
            XYZ b,
            ref int cntSegsAdded,
            bool isCurtain = false,
            int curtainRootId = 0)
        {
            if (a.DistanceTo(b) < sMinSegFt) return;

            var A = SnapXY(new XYZ(a.X, a.Y, 0));
            var B = SnapXY(new XYZ(b.X, b.Y, 0));

            if (Nearly(A.X, B.X) && Nearly(A.Y, B.Y)) return;

            var k = KeyFor(A, B);
            if (keys.Add(k))
            {
                segs.Add(new Seg2(A, B, isCurtain, curtainRootId));
                cntSegsAdded++;
            }
        }




        static int CurveToSegments2D(
    HashSet<string> keys,
    List<Seg2> segs,
    Curve c,
    Transform T_final,
    bool isCurtain = false,
    int curtainRootId = 0)
        {
            if (c == null) return 0;

            int cntSegsAdded = 0;

            if (c is Line line)
            {
                var a = T_final.OfPoint(line.GetEndPoint(0));
                var b = T_final.OfPoint(line.GetEndPoint(1));

                if (TryAddSeg2D(keys, segs, a, b, isCurtain, curtainRootId))
                    cntSegsAdded++;
            }
            else
            {
                var tess = c.Tessellate();
                if (tess != null && tess.Count >= 2)
                {
                    for (int i = 0; i + 1 < tess.Count; ++i)
                    {
                        var a = T_final.OfPoint(tess[i]);
                        var b = T_final.OfPoint(tess[i + 1]);

                        if (TryAddSeg2D(keys, segs, a, b, isCurtain, curtainRootId))
                            cntSegsAdded++;
                    }
                }
            }

            return cntSegsAdded;
        }


        static int PolylineToSegments2D(
    HashSet<string> keys,
    List<Seg2> segs,
    PolyLine pl,
    Transform T_final,
    bool isCurtain = false,
    int curtainRootId = 0)
        {
            if (pl == null) return 0;

            int cntSegsAdded = 0;
            var pts = pl.GetCoordinates();
            if (pts == null || pts.Count < 2) return 0;

            for (int i = 0; i + 1 < pts.Count; ++i)
            {
                var a = T_final.OfPoint(pts[i]);
                var b = T_final.OfPoint(pts[i + 1]);

                if (TryAddSeg2D(keys, segs, a, b, isCurtain, curtainRootId))
                    cntSegsAdded++;
            }

            return cntSegsAdded;
        }


        static void CollectFromElementGeometry(
    HashSet<string> keys,
    List<Seg2> segs,
    Element el,
    Options opts,
    Transform T_model_to_view,
    Transform extraTransform,
    bool markAsCurtain,
    ref int cntModelCat,
    ref int cntBBoxNull,
    ref int cntGeomNull,
    ref int cntCurves,
    ref int cntSolids,
    ref int cntSegsAdded)
        {
            try
            {
                var cat = el.Category;
                if (cat == null || cat.CategoryType != CategoryType.Model)
                    return;
            }
            catch
            {
                return;
            }

            cntModelCat++;
            // для витражей запоминаем корневую стену (один Id для всех панелей/импостов этой стены)
            int curtainRootId = 0;
            if (markAsCurtain)
            {
                try
                {
                    if (el is Wall w &&
                        w.WallType != null &&
                        w.WallType.Kind == WallKind.Curtain)
                    {
                        // сама витражная стена
                        curtainRootId = w.Id.IntegerValue;
                    }
                    else if (el is FamilyInstance fi)
                    {
                        Wall hostWall = null;

                        if (fi.Host is Wall hw)
                            hostWall = hw;
                        else if (fi.SuperComponent is FamilyInstance sfi && sfi.Host is Wall shw)
                            hostWall = shw;

                        if (hostWall != null &&
                            hostWall.WallType != null &&
                            hostWall.WallType.Kind == WallKind.Curtain)
                        {
                            // панели, импосты, витражные двери и т.п. получают тот же Id, что и стена-хост
                            curtainRootId = hostWall.Id.IntegerValue;
                        }
                    }
                }
                catch
                {
                    curtainRootId = 0;
                }
            }



            try
            {
                if (el.get_BoundingBox(opts.View) == null)
                {
                    cntBBoxNull++;
                    return;
                }
            }
            catch
            {
                cntBBoxNull++;
                return;
            }

            GeometryElement ge = null;
            try { ge = el.get_Geometry(opts); }
            catch { ge = null; }

            if (ge == null)
            {
                cntGeomNull++;
                return;
            }

            foreach (var obj in ge)
            {
                switch (obj)
                {
                    case Curve c:
                        {
                            cntCurves++;
                            var T_final = Compose(T_model_to_view, extraTransform, null);
                            cntSegsAdded += CurveToSegments2D(
                                keys,
                                segs,
                                c,
                                T_final,
                                markAsCurtain,
                                curtainRootId);
                            break;
                        }


                    case PolyLine pl:
                        {
                            cntCurves++;
                            var T_final = Compose(T_model_to_view, extraTransform, null);
                            cntSegsAdded += PolylineToSegments2D(
                                keys,
                                segs,
                                pl,
                                T_final,
                                markAsCurtain,
                                curtainRootId);
                            break;
                        }


                    case GeometryInstance gi:
                        {
                            GeometryElement inst = gi.GetSymbolGeometry();
                            Transform instT = gi.Transform;

                            if (inst == null)
                            {
                                inst = gi.GetInstanceGeometry();
                                instT = Transform.Identity;
                            }

                            if (inst == null) break;

                            foreach (var sub in inst)
                            {
                                var T_final = Compose(T_model_to_view, extraTransform, instT);

                                if (sub is Curve c2)
                                {
                                    cntCurves++;
                                    cntSegsAdded += CurveToSegments2D(
                                        keys,
                                        segs,
                                        c2,
                                        T_final,
                                        markAsCurtain,
                                        curtainRootId);
                                }
                                else if (sub is PolyLine pl2)
                                {
                                    cntCurves++;
                                    cntSegsAdded += PolylineToSegments2D(
                                        keys,
                                        segs,
                                        pl2,
                                        T_final,
                                        markAsCurtain,
                                        curtainRootId);
                                }
                                else if (sub is Solid s2 && s2.Edges != null && s2.Edges.Size > 0)
                                {
                                    cntSolids++;
                                    foreach (Edge e in s2.Edges)
                                    {
                                        try
                                        {
                                            cntSegsAdded += CurveToSegments2D(
                                                keys,
                                                segs,
                                                e.AsCurve(),
                                                T_final,
                                                markAsCurtain,
                                                curtainRootId);
                                        }
                                        catch { }
                                    }
                                }
                            }
                            break;
                        }

                    case Solid s:
                        {
                            if (s.Edges != null && s.Edges.Size > 0)
                            {
                                cntSolids++;
                                var T_final = Compose(T_model_to_view, extraTransform, null);
                                foreach (Edge e in s.Edges)
                                {
                                    try
                                    {
                                        // ВАЖНО: передаём curtainRootId, а не только флаг markAsCurtain
                                        cntSegsAdded += CurveToSegments2D(
                                            keys,
                                            segs,
                                            e.AsCurve(),
                                            T_final,
                                            markAsCurtain,
                                            curtainRootId);
                                    }
                                    catch { }
                                }
                            }
                            break;
                        }

                }
            }
        }

        // ---------- JSON (НЕ МЕНЯЛ) ----------
        // Описание помещения для JSON
        class RoomPayload
        {
            public string id;
            public string name;
            public string number;
            public string comment;

            // Основные петли помещения (как сейчас)
            public List<List<XYZ>> loops;

            // Новое: пути по витражам для этого помещения:
            // ключ = CurtainRootId, значение = список точек пути в координатах ВИДА (как loops до трансформации в лист)
            public Dictionary<int, List<XYZ>> curtainPaths;
        }

        // ---------- JSON ----------
        struct GroupPayload
        {
            public string id;
            public string name;
            public string source;
            public List<Seg2> segs;
            public List<Seg2> cutouts;
            public List<RoomPayload> rooms;
        }

        // делаем доступ к флагу режима из статических методов
        internal static bool IsModeOPA() => sModeOPA;

        static void WriteJsonMulti(
            string path, List<GroupPayload> groups, string sheetName, string unitName,
            IEnumerable<string> cutCatNames, IEnumerable<string> cutFamNames, bool includeLinks, bool systemShafts, bool cutEnabled)
        {
            var inv = CultureInfo.InvariantCulture;
            var sb = new StringBuilder(1 << 20);
            sb.Append("{\n  \"groups\": [\n");

            for (int gi = 0; gi < groups.Count; gi++)
            {
                var g = groups[gi];
                sb.Append("    {\n");
                sb.Append("      \"id\": \"").Append(g.id.Replace("\"", "\\\"")).Append("\",\n");
                sb.Append("      \"name\": \"").Append(g.name.Replace("\"", "\\\"")).Append("\",\n");
                sb.Append("      \"tags\": { \"source\": \"").Append(g.source.Replace("\"", "\\\"")).Append("\" },\n");

                sb.Append("      \"segments\": [\n");
                for (int i = 0; i < g.segs.Count; i++)
                {
                    var a = g.segs[i].A; var b = g.segs[i].B;
                    sb.Append("        [[")
                      .Append(a.X.ToString(inv)).Append(", ").Append(a.Y.ToString(inv)).Append("], [")
                      .Append(b.X.ToString(inv)).Append(", ").Append(b.Y.ToString(inv)).Append("]]");
                    if (i + 1 < g.segs.Count) sb.Append(',');
                    sb.Append('\n');
                }
                sb.Append("      ],\n");

                sb.Append("      \"cutouts\": [\n");
                for (int i = 0; i < g.cutouts.Count; i++)
                {
                    var a = g.cutouts[i].A; var b = g.cutouts[i].B;
                    sb.Append("        [[")
                      .Append(a.X.ToString(inv)).Append(", ").Append(a.Y.ToString(inv)).Append("], [")
                      .Append(b.X.ToString(inv)).Append(", ").Append(b.Y.ToString(inv)).Append("]]");
                    if (i + 1 < g.cutouts.Count) sb.Append(',');
                    sb.Append('\n');
                }
                sb.Append("      ],\n");

                sb.Append("      \"rooms\": [\n");
                if (g.rooms != null && g.rooms.Count > 0)
                {
                    for (int ri = 0; ri < g.rooms.Count; ri++)
                    {
                        var r = g.rooms[ri];
                        sb.Append("        {\n");
                        sb.Append("          \"id\": \"").Append((r.id ?? "").Replace("\"", "\\\"")).Append("\",\n");
                        sb.Append("          \"name\": \"").Append((r.name ?? "").Replace("\"", "\\\"")).Append("\",\n");
                        sb.Append("          \"number\": \"").Append((r.number ?? "").Replace("\"", "\\\"")).Append("\",\n");
                        sb.Append("          \"comment\": \"").Append((r.comment ?? "").Replace("\"", "\\\"")).Append("\",\n");

                        // loops
                        sb.Append("          \"loops\": [\n");

                        if (r.loops != null && r.loops.Count > 0)
                        {
                            for (int li = 0; li < r.loops.Count; li++)
                            {
                                var loop = r.loops[li];
                                sb.Append("            [\n");
                                for (int pi = 0; pi < loop.Count; pi++)
                                {
                                    var p = loop[pi];
                                    sb.Append("              [")
                                      .Append(p.X.ToString(inv)).Append(", ")
                                      .Append(p.Y.ToString(inv)).Append("]");
                                    if (pi + 1 < loop.Count) sb.Append(',');
                                    sb.Append('\n');
                                }
                                sb.Append("            ]");
                                if (li + 1 < r.loops.Count) sb.Append(',');
                                sb.Append('\n');
                            }
                        }

                        sb.Append("          ],\n");

                        // НОВОЕ: curtain_paths (по витражам)
                        sb.Append("          \"curtain_paths\": {\n");
                        if (r.curtainPaths != null && r.curtainPaths.Count > 0)
                        {
                            int ci = 0;
                            foreach (var kvp in r.curtainPaths)
                            {
                                int curtainId = kvp.Key;
                                var pts = kvp.Value ?? new List<XYZ>();

                                sb.Append("            \"").Append(curtainId.ToString(inv)).Append("\": [\n");
                                for (int pi = 0; pi < pts.Count; pi++)
                                {
                                    var p = pts[pi];
                                    sb.Append("              [")
                                      .Append(p.X.ToString(inv)).Append(", ")
                                      .Append(p.Y.ToString(inv)).Append("]");
                                    if (pi + 1 < pts.Count) sb.Append(',');
                                    sb.Append('\n');
                                }
                                sb.Append("            ]");
                                if (++ci < r.curtainPaths.Count) sb.Append(',');
                                sb.Append('\n');
                            }
                        }
                        sb.Append("          }\n");

                        sb.Append("        }");

                        if (ri + 1 < g.rooms.Count) sb.Append(',');
                        sb.Append('\n');
                    }
                }
                sb.Append("      ]\n");

                sb.Append("    }");

                if (gi + 1 < groups.Count) sb.Append(',');
                sb.Append('\n');
            }

            sb.Append("  ],\n  \"meta\": {\n");
            sb.Append("    \"source\": \"Revit Sheet Plans 2D\",\n");
            sb.Append("    \"sheet\": \"").Append(sheetName?.Replace("\"", "\\\"") ?? "").Append("\",\n");
            sb.Append("    \"units_length\": \"").Append(unitName).Append("\",\n");
            sb.Append("    \"calc_mode\": \"").Append(IsModeOPA() ? "OPA" : "GNS").Append("\",\n");
            sb.Append("    \"cutout_filters\": {\n");

            sb.Append("      \"enabled\": ").Append(cutEnabled ? "true" : "false").Append(",\n");
            sb.Append("      \"mode\": \"").Append(systemShafts ? "system" : "custom").Append("\",\n");
            sb.Append("      \"include_links\": ").Append(includeLinks ? "true" : "false").Append(",\n");
            sb.Append("      \"categories\": [")
                .Append(string.Join(", ", cutCatNames.Select(n => $"\"{n}\"")))
                .Append("],\n");
            sb.Append("      \"families\": [")
              .Append(string.Join(", ", cutFamNames.Select(n => $"\"{n.Replace("\"", "\\\"")}\"")))
              .Append("]\n");
            sb.Append("    }\n");
            sb.Append("  }\n");
            sb.Append("}\n");

            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(false));
        }

        // ---------- LINK VIEW RESOLVE ----------
        // ---------- LINK VIEW RESOLVE ----------
        static View TryGetLinkedViewForHost(View hostView, RevitLinkInstance link, out string how)
        {
            how = "none";

            if (hostView == null || link == null)
            {
                how = "args-null";
                return null;
            }

            var ldoc = link.GetLinkDocument();
            if (ldoc == null)
            {
                how = "link-doc-null";
                return null;
            }

            ElementId linkedViewId = ElementId.InvalidElementId;
            string linkVisMode = null; // "ByHostView", "ByLinkView", "Custom", ...

            // Пытаемся вытащить RevitLinkGraphicsSettings через View.GetLinkOverrides(link.Id)
            try
            {
                var miGetLinkOverrides = typeof(View).GetMethod("GetLinkOverrides", new Type[] { typeof(ElementId) });
                var rlgs = miGetLinkOverrides != null ? miGetLinkOverrides.Invoke(hostView, new object[] { link.Id }) : null;

                if (rlgs != null)
                {
                    var tGs = rlgs.GetType();

                    // 1) Режим отображения (ByHostView / ByLinkView / Custom) — соответствует радио-кнопкам в "Revit связи"
                    var piVis = tGs.GetProperty("LinkVisibilityType");
                    if (piVis != null)
                    {
                        var visObj = piVis.GetValue(rlgs, null);
                        if (visObj != null)
                            linkVisMode = visObj.ToString(); // "ByHostView", "ByLinkView", "Custom", ...
                    }

                    // 2) Привязанный Linked View (если задан)
                    var piLid = tGs.GetProperty("LinkedViewId");
                    if (piLid != null)
                    {
                        var lidObj = piLid.GetValue(rlgs, null);
                        if (lidObj is ElementId eid)
                            linkedViewId = eid;
                    }
                    else
                    {
                        // Для совместимости — старый стиль через метод GetLinkedViewId()
                        var miLid = tGs.GetMethod("GetLinkedViewId", Type.EmptyTypes);
                        if (miLid != null)
                        {
                            var lidObj2 = miLid.Invoke(rlgs, null);
                            if (lidObj2 is ElementId eid2)
                                linkedViewId = eid2;
                        }
                    }
                }
            }
            catch
            {
                // тихий fallback
            }

            // 1) Если в "Revit связи" выбран конкретный Linked View
            //    (Display Settings = "По связанному виду" или "Пользовательский" с указанным видом)
            if (linkedViewId != ElementId.InvalidElementId)
            {
                try
                {
                    var vLinked = ldoc.GetElement(linkedViewId) as View;
                    if (vLinked != null && IsViewOK(ldoc, vLinked))
                    {
                        if (!string.IsNullOrEmpty(linkVisMode))
                            how = "overrides-" + linkVisMode; // например "overrides-ByLinkView"
                        else
                            how = "overrides-linked-view";
                        return vLinked;
                    }
                }
                catch
                {
                    // пойдём в fallback ниже
                }
            }

            // 2) Режим "По основному виду" (ByHostView) — Revit графику берёт из хоста,
            //    но API не даёт прямого "вида линка" для этого случая.
            //    Делаем best-effort: сначала ищем план с тем же именем, что и hostView.
            try
            {
                var vByName = new FilteredElementCollector(ldoc)
                    .OfClass(typeof(ViewPlan))
                    .Cast<ViewPlan>()
                    .FirstOrDefault(vp =>
                    {
                        try { if (vp.IsTemplate) return false; } catch { }
                        return string.Equals(vp.Name, hostView.Name, StringComparison.OrdinalIgnoreCase)
                               && IsViewOK(ldoc, vp);
                    });

                if (vByName != null)
                {
                    how = string.IsNullOrEmpty(linkVisMode)
                        ? "name-match"
                        : "name-match-" + linkVisMode; // например "name-match-ByHostView"
                    return vByName;
                }
            }
            catch
            {
                // игнорируем, пойдём в финальный fallback
            }

            // 3) Финальный fallback — первый пригодный план из связанного файла
            var any = FirstUsablePlan(ldoc);
            if (any != null)
            {
                how = string.IsNullOrEmpty(linkVisMode)
                    ? "first-usable-plan-fallback"
                    : "fallback-" + linkVisMode;
                return any;
            }

            how = "no-plan-found";
            return null;
        }



        // ---------- PREFS ----------
        class CutoutPrefs
        {
            public bool Enabled = true;
            public bool SystemShafts = false;
            public bool CustomEnabled = true;
            public bool IncludeLinks = true;
            public List<int> Categories = new List<int>();
            public List<string> Families = new List<string>();

            // "GNS" – ГНС, "OPA" – Общая площадь
            public string ExportMode = "GNS";
            // Имя параметра Room, из которого читаем признак МОП в режиме OPA
            public string MopParamName = "";

        }


        static string PrefsPath()
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Palmsi");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "ExportCurves.settings.ini");
        }
        static CutoutPrefs LoadPrefs()
        {
            var pr = new CutoutPrefs();
            try
            {
                var p = PrefsPath(); if (!File.Exists(p)) return pr;
                var lines = File.ReadAllLines(p, Encoding.UTF8);
                string Get(string key) => lines.Select(s => s.Split(new[] { '=' }, 2))
                                               .FirstOrDefault(a => a.Length == 2 && a[0] == key)?[1];
                pr.Enabled = Get("Enabled") == "1";
                pr.SystemShafts = Get("SystemShafts") == "1";
                pr.IncludeLinks = Get("IncludeLinks") == "1";
                var ce = Get("CustomEnabled");
                if (ce == "0" || ce == "1")
                    pr.CustomEnabled = (ce == "1");

                var cats = Get("Categories");
                if (!string.IsNullOrEmpty(cats))
                    pr.Categories = cats.Split(';').Select(s => { int x; return int.TryParse(s, out x) ? x : (int?)null; })
                                        .Where(x => x.HasValue).Select(x => x.Value).ToList();
                var fams = Get("Families");
                if (!string.IsNullOrEmpty(fams))
                    pr.Families = fams.Split(';').Where(s => !string.IsNullOrWhiteSpace(s)).ToList();

                var mode = Get("ExportMode");
                if (!string.IsNullOrEmpty(mode))
                    pr.ExportMode = mode;
                var mop = Get("MopParamName");
                if (!string.IsNullOrEmpty(mop))
                    pr.MopParamName = mop;

            }
            catch { }
            return pr;
        }
        static void SavePrefs(CutoutPrefs pr)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("Enabled=" + (pr.Enabled ? "1" : "0"));
                sb.AppendLine("SystemShafts=" + (pr.SystemShafts ? "1" : "0"));
                sb.AppendLine("IncludeLinks=" + (pr.IncludeLinks ? "1" : "0"));
                sb.AppendLine("Categories=" + string.Join(";", pr.Categories ?? new List<int>()));
                sb.AppendLine("Families=" + string.Join(";", pr.Families ?? new List<string>()));
                sb.AppendLine("CustomEnabled=" + (pr.CustomEnabled ? "1" : "0"));
                sb.AppendLine("ExportMode=" + (pr.ExportMode ?? "GNS"));
                sb.AppendLine("MopParamName=" + (pr.MopParamName ?? ""));


                File.WriteAllText(PrefsPath(), sb.ToString(), new UTF8Encoding(false));
            }
            catch { }
        }

        // ---------- WPF-окно без XAML ----------

        class CutoutDialogWindow : Wpf.Window
        {
            // Публичные геттеры для чтения выбора
            public bool EnableCutouts => _rbYes.IsChecked == true;
            public bool UseSystemShafts => _cbSystem.IsChecked == true;
            public bool UseCustom => _cbCustom.IsChecked == true;
            public bool IncludeLinks => _cbLinks.IsChecked == true;

            // Режим экспорта: "GNS" или "OPA"
            public string ExportMode
            {
                get
                {
                    if (_rbModeOPA != null && _rbModeOPA.IsChecked == true)
                        return "OPA";
                    return "GNS";
                }
            }


            // Категория считается выбранной, если на ней стоит галочка ИЛИ выбрано хотя бы одно семейство внутри неё
            public List<BuiltInCategory> SelectedCategories
            {
                get
                {
                    var set = new HashSet<BuiltInCategory>();
                    foreach (var kv in _catCheckByBic)
                        if (kv.Value.IsChecked == true) set.Add(kv.Key);

                    foreach (var kv in _famChosen)
                        if (kv.Value != null && kv.Value.Count > 0) set.Add(kv.Key);

                    return set.ToList();
                }
            }

            public List<string> SelectedFamilies =>
                _famChosen.Values.SelectMany(s => s).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            // --- визуальные элементы ---
            // --- визуальные элементы ---
            private readonly WpfControls.RadioButton _rbNo, _rbYes;

            // новый блок: радиокнопки режима
            private readonly WpfControls.RadioButton _rbModeGNS;
            private readonly WpfControls.RadioButton _rbModeOPA;
            private readonly WpfControls.ComboBox _cbMopParam;

            // выбранное имя параметра Room для МОП
            public string MopParamName
            {
                get
                {
                    return _cbMopParam?.SelectedItem as string ?? "";
                }
            }


            private readonly WpfControls.CheckBox _cbSystem, _cbCustom;
            private readonly WpfControls.CheckBox _cbLinks;

            private readonly WpfControls.ListBox _lbCats, _lbFams;
            private readonly WpfControls.GroupBox _gbCats, _gbFams;

            // --- вспомогательные структуры для синхронизации ---
            private readonly List<WpfControls.CheckBox> _catChecks = new List<WpfControls.CheckBox>();
            private readonly List<WpfControls.CheckBox> _famChecks = new List<WpfControls.CheckBox>();

            private BuiltInCategory? _currentCatBic = null; // Какая категория сейчас выбрана для просмотра семейств

            // Чекбокс категории по её BIC
            private readonly Dictionary<BuiltInCategory, WpfControls.CheckBox> _catCheckByBic
                = new Dictionary<BuiltInCategory, WpfControls.CheckBox>();

            // Глобально выбранные семейства по категориям
            private readonly Dictionary<BuiltInCategory, HashSet<string>> _famChosen
                = new Dictionary<BuiltInCategory, HashSet<string>>();

            public class CatItem { public string Display; public BuiltInCategory Bic; public override string ToString() => Display; }
            public class FamItem { public string Name; public override string ToString() => Name; }

            private readonly Dictionary<BuiltInCategory, List<string>> _map;

            public CutoutDialogWindow(Dictionary<BuiltInCategory, List<string>> cats, List<string> roomParams, CutoutPrefs prev)

            {

                Title = "Экспорт в JSON";
                WindowStartupLocation = Wpf.WindowStartupLocation.CenterOwner;
                ResizeMode = Wpf.ResizeMode.NoResize;
                Width = 620; Height = 420;

                _map = cats;

                // --- корневая сетка ---
                var root = new WpfControls.Grid { Margin = new Wpf.Thickness(16) };

                // 0: режим (ГНС / Общая площадь)
                // 1: выбор параметра МОП
                // 2: вопрос "Вырезать шахты?"
                // 3: чекбоксы режимов шахт
                // 4: списки категорий/семейств
                // 5: кнопки
                root.RowDefinitions.Add(new WpfControls.RowDefinition { Height = Wpf.GridLength.Auto });
                root.RowDefinitions.Add(new WpfControls.RowDefinition { Height = Wpf.GridLength.Auto });
                root.RowDefinitions.Add(new WpfControls.RowDefinition { Height = Wpf.GridLength.Auto });
                root.RowDefinitions.Add(new WpfControls.RowDefinition { Height = Wpf.GridLength.Auto });
                root.RowDefinitions.Add(new WpfControls.RowDefinition { Height = new Wpf.GridLength(1, Wpf.GridUnitType.Star) });
                root.RowDefinitions.Add(new WpfControls.RowDefinition { Height = Wpf.GridLength.Auto });

                // --- Режим экспорта: ГНС / Общая площадь ---
                var modePanel = new WpfControls.StackPanel
                {
                    Orientation = WpfControls.Orientation.Horizontal,
                    Margin = new Wpf.Thickness(0, 0, 0, 8)
                };

                modePanel.Children.Add(new WpfControls.TextBlock
                {
                    Text = "Режим:",
                    FontSize = 16,
                    Margin = new Wpf.Thickness(0, 0, 12, 0)
                });

                var rbModeGNS = new WpfControls.RadioButton
                {
                    Content = "ГНС",
                    Margin = new Wpf.Thickness(0, 0, 12, 0)
                };

                var rbModeOPA = new WpfControls.RadioButton
                {
                    Content = "Общая площадь"
                };

                // по умолчанию из настроек
                if (prev != null && prev.ExportMode == "OPA")
                    rbModeOPA.IsChecked = true;
                else
                    rbModeGNS.IsChecked = true;

                modePanel.Children.Add(rbModeGNS);
                modePanel.Children.Add(rbModeOPA);

                WpfControls.Grid.SetRow(modePanel, 0);
                root.Children.Add(modePanel);

                _rbModeGNS = rbModeGNS;
                _rbModeOPA = rbModeOPA;

                modePanel.Children.Add(new WpfControls.TextBlock
                {
                    Text = "   МОП параметр:",
                    Margin = new Wpf.Thickness(16, 0, 8, 0),
                    VerticalAlignment = Wpf.VerticalAlignment.Center
                });

                _cbMopParam = new WpfControls.ComboBox
                {
                    Width = 200,
                    MinWidth = 200,
                    IsEditable = false,
                    Margin = new Wpf.Thickness(0, 0, 0, 0)
                };

                // наполняем список (это ИМЕНА параметров Room)
                if (roomParams != null)
                {
                    foreach (var n in roomParams)
                        _cbMopParam.Items.Add(n);
                }


                // дефолт: из настроек, иначе "Комментарии/Comments", иначе первый
                string want = prev?.MopParamName ?? "";
                int pick = -1;

                if (!string.IsNullOrWhiteSpace(want))
                {
                    for (int i = 0; i < _cbMopParam.Items.Count; i++)
                        if (string.Equals(_cbMopParam.Items[i] as string, want, StringComparison.OrdinalIgnoreCase))
                        { pick = i; break; }
                }

                if (pick < 0)
                {
                    for (int i = 0; i < _cbMopParam.Items.Count; i++)
                    {
                        var s = (_cbMopParam.Items[i] as string) ?? "";
                        if (s.Equals("Комментарии", StringComparison.OrdinalIgnoreCase) ||
                            s.Equals("Comments", StringComparison.OrdinalIgnoreCase))
                        { pick = i; break; }
                    }
                }

                if (pick < 0 && _cbMopParam.Items.Count > 0) pick = 0;
                _cbMopParam.SelectedIndex = pick;

                // Логика доступности: параметр МОП нужен только в режиме OPA
                _cbMopParam.IsEnabled = (rbModeOPA.IsChecked == true);

                rbModeOPA.Checked += (s, e) => { _cbMopParam.IsEnabled = true; };
                rbModeGNS.Checked += (s, e) => { _cbMopParam.IsEnabled = false; };

                modePanel.Children.Add(_cbMopParam);

                // Q1: Вырезать шахты?
                var q = new WpfControls.StackPanel { Orientation = WpfControls.Orientation.Horizontal, Margin = new Wpf.Thickness(0, 0, 0, 8) };
                q.Children.Add(new WpfControls.TextBlock { Text = "Вырезать шахты?", FontSize = 16, Margin = new Wpf.Thickness(0, 0, 12, 0) });
                _rbNo = new WpfControls.RadioButton { Content = "Нет", Margin = new Wpf.Thickness(0, 0, 12, 0) };
                _rbYes = new WpfControls.RadioButton { Content = "Да" };
                q.Children.Add(_rbNo); q.Children.Add(_rbYes);
                WpfControls.Grid.SetRow(q, 2); root.Children.Add(q);

                // Режимы + ссылки
                var modes = new WpfControls.StackPanel { Orientation = WpfControls.Orientation.Horizontal, Margin = new Wpf.Thickness(0, 0, 0, 8) };
                _cbSystem = new WpfControls.CheckBox { Content = "Системные шахты", Margin = new Wpf.Thickness(0, 0, 18, 0) };
                _cbCustom = new WpfControls.CheckBox { Content = "Пользовательские (катег./сем.)" };
                _cbLinks = new WpfControls.CheckBox { Content = "Учитывать в ссылках", Margin = new Wpf.Thickness(24, 0, 0, 0) };
                modes.Children.Add(_cbSystem); modes.Children.Add(_cbCustom); modes.Children.Add(_cbLinks);
                WpfControls.Grid.SetRow(modes, 3); root.Children.Add(modes);

                // Две колонки списков
                var grids = new WpfControls.Grid { Margin = new Wpf.Thickness(0, 0, 0, 8) };
                grids.ColumnDefinitions.Add(new WpfControls.ColumnDefinition { Width = new Wpf.GridLength(1, Wpf.GridUnitType.Star) });
                grids.ColumnDefinitions.Add(new WpfControls.ColumnDefinition { Width = new Wpf.GridLength(1, Wpf.GridUnitType.Star) });

                _gbCats = new WpfControls.GroupBox { Header = "Категории (можно несколько):", Margin = new Wpf.Thickness(0, 0, 8, 0) };
                _lbCats = new WpfControls.ListBox { Height = 240, SelectionMode = WpfControls.SelectionMode.Single };
                _gbCats.Content = _lbCats; WpfControls.Grid.SetColumn(_gbCats, 0); grids.Children.Add(_gbCats);

                _gbFams = new WpfControls.GroupBox { Header = "Семейства (можно несколько):", Margin = new Wpf.Thickness(8, 0, 0, 0) };
                _lbFams = new WpfControls.ListBox { Height = 240, SelectionMode = WpfControls.SelectionMode.Extended };
                _gbFams.Content = _lbFams; WpfControls.Grid.SetColumn(_gbFams, 1); grids.Children.Add(_gbFams);

                WpfControls.Grid.SetRow(grids, 4); root.Children.Add(grids);

                // Кнопки
                var btns = new WpfControls.StackPanel { Orientation = WpfControls.Orientation.Horizontal, HorizontalAlignment = Wpf.HorizontalAlignment.Right };
                var ok = new WpfControls.Button { Content = "Экспорт .json", IsDefault = true, Padding = new Wpf.Thickness(12, 4, 12, 4), Margin = new Wpf.Thickness(0, 0, 8, 0) };
                var cancel = new WpfControls.Button { Content = "Отмена", IsCancel = true, Padding = new Wpf.Thickness(12, 4, 12, 4) };
                ok.Click += (s, e) => { DialogResult = true; Close(); };
                btns.Children.Add(ok); btns.Children.Add(cancel);
                WpfControls.Grid.SetRow(btns, 5); root.Children.Add(btns);

                Content = root;

                // --- наполняем список категорий ---
                foreach (var kv in cats.OrderBy(k => k.Key.ToString()))
                {
                    var item = new CatItem { Display = kv.Key.ToString().Replace("OST_", ""), Bic = kv.Key };

                    // Хранилище выбранных семейств по этой категории
                    _famChosen[item.Bic] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    var cbCat = new WpfControls.CheckBox
                    {
                        Content = item.Display,
                        Tag = item,
                        Margin = new Wpf.Thickness(2, 1, 2, 1),
                        IsChecked = (prev.Categories?.Contains((int)item.Bic) ?? false) // восстановление галочки
                    };

                    _catChecks.Add(cbCat);
                    _catCheckByBic[item.Bic] = cbCat;
                    _lbCats.Items.Add(cbCat);

                    // Сняли галочку с категории → снять все семейства этой категории
                    cbCat.Unchecked += (s, e) =>
                    {
                        var catBic = ((CatItem)cbCat.Tag).Bic;

                        if (_famChosen.TryGetValue(catBic, out var set))
                            set.Clear();

                        if (_currentCatBic == catBic)
                        {
                            // Снять все галки на правом списке (визуально)
                            foreach (var fcb in _famChecks)
                                fcb.IsChecked = false;
                        }
                    };

                    // По ТЗ при установке галочки на категории автоматически семейства не трогаем
                    cbCat.Checked += (s, e) => { /* ничего не делаем */ };
                }

                // Переключение выбора категории (левая колонка) — просто меняем _currentCatBic и обновляем список семейств
                _lbCats.SelectionChanged += (s, e) =>
                {
                    var sel = _lbCats.SelectedItem as WpfControls.CheckBox;
                    _currentCatBic = (sel?.Tag as CatItem)?.Bic;
                    RefreshFamilies();
                };

                // Выставим дефолтную «текущую» категорию
                if (_lbCats.Items.Count > 0 && _lbCats.SelectedIndex < 0)
                    _lbCats.SelectedIndex = 0;
                var sel2 = _lbCats.SelectedItem as WpfControls.CheckBox;
                _currentCatBic = (sel2?.Tag as CatItem)?.Bic;

                // Раскладываем сохранённые семейства по своим категориям
                if (prev.Families != null && prev.Families.Count > 0)
                {
                    var wantFamilies = new HashSet<string>(prev.Families, StringComparer.OrdinalIgnoreCase);
                    foreach (var kv2 in cats)
                    {
                        var bic = kv2.Key;
                        if (!_famChosen.ContainsKey(bic))
                            _famChosen[bic] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        foreach (var fam in kv2.Value ?? new List<string>())
                            if (wantFamilies.Contains(fam))
                                _famChosen[bic].Add(fam);
                    }
                }

                // restore верхние переключатели
                _rbYes.IsChecked = prev.Enabled;
                _rbNo.IsChecked = !prev.Enabled;
                _cbSystem.IsChecked = prev.SystemShafts;
                _cbCustom.IsChecked = prev.CustomEnabled;
                _cbLinks.IsChecked = prev.IncludeLinks;

                // включение/отключение блоков
                _rbYes.Checked += (s, e) => ToggleAll(true);
                _rbNo.Checked += (s, e) => ToggleAll(false);
                _cbCustom.Checked += (s, e) => ToggleCustomSelectors(true);
                _cbCustom.Unchecked += (s, e) => ToggleCustomSelectors(false);

                ToggleAll(_rbYes.IsChecked == true);
                ToggleCustomSelectors(_cbCustom.IsChecked == true);

                // Заполнить правый список для текущей категории и применить сохранённые галочки семейств
                RefreshFamilies();

                if (prev.Families != null && prev.Families.Count > 0)
                {
                    foreach (var fcb in _famChecks)
                    {
                        var it = fcb.Tag as FamItem;
                        if (it != null && prev.Families.Contains(it.Name))
                            fcb.IsChecked = true;
                    }
                }
            }

            // --- приватные методы-помощники (без локальных функций, совместимо с C# 7.3) ---

            private void ToggleCustomSelectors(bool on)
            {
                _gbCats.IsEnabled = on; _lbCats.IsEnabled = on;
                _gbFams.IsEnabled = on; _lbFams.IsEnabled = on;
                _gbCats.Opacity = on ? 1.0 : 0.5;
                _gbFams.Opacity = on ? 1.0 : 0.5;
            }

            private void ToggleAll(bool on)
            {
                _cbSystem.IsEnabled = on;
                _cbCustom.IsEnabled = on;
                _cbLinks.IsEnabled = on;
                ToggleCustomSelectors(on && (_cbCustom.IsChecked == true));
            }

            private void RefreshFamilies()
            {
                _lbFams.Items.Clear();
                _famChecks.Clear();

                if (_currentCatBic == null) return;

                var bic = _currentCatBic.Value;

                if (!_map.TryGetValue(bic, out var famList) || famList == null || famList.Count == 0)
                    return;

                // уже выбранные семейства по этой категории
                if (!_famChosen.TryGetValue(bic, out var chosen))
                    chosen = _famChosen[bic] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var f in famList
                        .Where(s => !string.IsNullOrWhiteSpace(s))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(x => x))
                {
                    var item = new FamItem { Name = f };
                    var cbFam = new WpfControls.CheckBox
                    {
                        Content = item.Name,
                        Tag = item,
                        Margin = new Wpf.Thickness(2, 1, 2, 1),
                        IsChecked = chosen.Contains(f)
                    };

                    // Семейство включили → добавить в chosen и гарантировать галочку категории
                    cbFam.Checked += (s, e) =>
                    {
                        chosen.Add(item.Name);
                        if (_catCheckByBic.TryGetValue(bic, out var catCb) && catCb.IsChecked != true)
                            catCb.IsChecked = true;
                    };

                    // Семейство сняли → убрать из chosen; если стало 0, снять галочку категории
                    cbFam.Unchecked += (s, e) =>
                    {
                        chosen.Remove(item.Name);
                        if (chosen.Count == 0)
                        {
                            if (_catCheckByBic.TryGetValue(bic, out var catCb) && catCb.IsChecked == true)
                                catCb.IsChecked = false;
                        }
                    };

                    _famChecks.Add(cbFam);
                    _lbFams.Items.Add(cbFam);
                }
            }
        }







        // ---------- сбор справочника категорий/семейств ----------
        static Dictionary<BuiltInCategory, List<string>> BuildCatalogForSheet(Document doc, ViewSheet sheet)
        {
            var dict = new Dictionary<BuiltInCategory, List<string>>();
            var allowedFilter = new ElementMulticategoryFilter(
                BASE_ALLOWED_BIC.Select(b => new ElementId(b)).ToList()
            );

            Action<BuiltInCategory, string> add = (bic, fam) =>
            {
                List<string> list;
                if (!dict.TryGetValue(bic, out list)) { list = new List<string>(); dict[bic] = list; }
                if (!string.IsNullOrWhiteSpace(fam) && !list.Contains(fam, StringComparer.OrdinalIgnoreCase))
                    list.Add(fam);
            };


            foreach (ElementId vpId in sheet.GetAllViewports())
            {
                var vp = doc.GetElement(vpId) as Viewport; if (vp == null) continue;
                var v = doc.GetElement(vp.ViewId) as View; if (v == null) continue;
                if (!IsViewOK(doc, v)) continue;   // <— пропускаем легенды, драфтинги, шаблоны и пр

                foreach (var el in new FilteredElementCollector(doc, v.Id)
                                    .WherePasses(allowedFilter)
                                    .WhereElementIsNotElementType())
                {
                    if (el.Category == null) continue;
                    var bic = (BuiltInCategory)el.Category.Id.IntegerValue;
                    if (DENY_BIC.Contains(bic)) continue;
                    var fam = (el as FamilyInstance) != null ? ((FamilyInstance)el).Symbol?.Family?.Name : null;
                    add(bic, fam);
                }


                // LINKS
                // LINKS (Revit 2024+: «как видно в виде хоста v»)
                var links = new FilteredElementCollector(doc, v.Id)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .ToList();

                foreach (var link in links)
                {
                    try
                    {
                        if (!IsElementVisibleInViewSafe(doc, v, link.Id))
                            continue;
                    }
                    catch { }

                    // СТАЛО:
                    var elems = GetVisibleFromLinkSmart(doc, v, link, allowedFilter);


                    foreach (var el in elems)
                    {
                        if (el.Category == null) continue;
                        var bic = (BuiltInCategory)el.Category.Id.IntegerValue;
                        if (DENY_BIC.Contains(bic)) continue;
                        var fam = (el as FamilyInstance) != null ? ((FamilyInstance)el).Symbol?.Family?.Name : null;
                        add(bic, fam);
                    }
                }


            }

            if (!dict.ContainsKey(BuiltInCategory.OST_ShaftOpening))
                dict[BuiltInCategory.OST_ShaftOpening] = new List<string>();

            return dict;
        }

        static Room GetAnyPlacedRoom(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Cast<Room>()
                .FirstOrDefault(r =>
                {
                    try { return r != null && r.Area > 0; }
                    catch { return false; }
                });
        }

        // Берём параметры "как есть" из любого помещения.
        // По умолчанию оставляем только String, т.к. IsMopMark() ищет "МОП/MOP" в тексте.
        static List<string> BuildRoomParamCatalogSimple(Document doc)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var room = GetAnyPlacedRoom(doc);
            if (room != null)
            {
                try
                {
                    foreach (Parameter p in room.Parameters)
                    {
                        if (p?.Definition == null) continue;

                        // Оставляем только строковые, чтобы МОП определять по тексту ("МОП"/"MOP")
                        if (p.StorageType != StorageType.String) continue;

                        var n = (p.Definition.Name ?? "").Trim();
                        if (n.Length == 0) continue;

                        names.Add(n);
                    }
                }
                catch { /* ignore */ }
            }

            // Fallback, если вообще не нашли ни одного строкового параметра
            if (names.Count == 0)
            {
                names.Add("Комментарии");
                names.Add("Comments");
            }

            return names.OrderBy(x => x).ToList();
        }

        static List<List<XYZ>> BuildLoopsFromSegs(List<Seg2> segs)
        {
            var loops = new List<List<XYZ>>();
            if (segs == null || segs.Count == 0)
                return loops;

            int n = segs.Count;
            var used = new bool[n];

            // точка -> список индексов сегментов, которые к ней примыкают
            var adj = new Dictionary<string, List<int>>();

            for (int i = 0; i < n; ++i)
            {
                var kA = KeyForPoint(segs[i].A);
                var kB = KeyForPoint(segs[i].B);

                if (!adj.TryGetValue(kA, out var la)) { la = new List<int>(); adj[kA] = la; }
                la.Add(i);

                if (!adj.TryGetValue(kB, out var lb)) { lb = new List<int>(); adj[kB] = lb; }
                lb.Add(i);
            }

            for (int i = 0; i < n; ++i)
            {
                if (used[i]) continue;

                var loop = new List<XYZ>();

                int segIndex = i;
                used[segIndex] = true;

                XYZ start = segs[segIndex].A;
                XYZ current = segs[segIndex].B;

                loop.Add(start);
                loop.Add(current);

                string startKey = KeyForPoint(start);

                for (; ; )
                {
                    string curKey = KeyForPoint(current);
                    if (!adj.TryGetValue(curKey, out var incident))
                        break;

                    int nextIndex = -1;
                    bool forward = true;

                    foreach (var idx in incident)
                    {
                        if (used[idx]) continue;
                        var s = segs[idx];

                        if (KeyForPoint(s.A) == curKey)
                        {
                            nextIndex = idx;
                            forward = true;
                            break;
                        }
                        if (KeyForPoint(s.B) == curKey)
                        {
                            nextIndex = idx;
                            forward = false;
                            break;
                        }
                    }

                    if (nextIndex < 0)
                        break;

                    used[nextIndex] = true;
                    var sNext = segs[nextIndex];
                    var nextPoint = forward ? sNext.B : sNext.A;
                    string nextKey = KeyForPoint(nextPoint);

                    loop.Add(nextPoint);

                    if (nextKey == startKey)
                        break;

                    current = nextPoint;
                }

                if (loop.Count >= 3)
                    loops.Add(loop);
            }

            return loops;
        }


        class PGNode
        {
            public XYZ P;
            public List<int> Out = new List<int>(); // индексы half-edge, выходящих из узла
        }

        class PGHalfEdge
        {
            public int From;
            public int To;
            public int Twin;     // индекс обратного half-edge
            public double Angle; // atan2 на From
            public bool Used;    // для обхода faces
        }

        static double Dot2D(XYZ a, XYZ b) => a.X * b.X + a.Y * b.Y;

        static bool BBoxOverlaps2D(XYZ a, XYZ b, XYZ c, XYZ d, double eps)
        {
            double minAx = Math.Min(a.X, b.X) - eps;
            double maxAx = Math.Max(a.X, b.X) + eps;
            double minAy = Math.Min(a.Y, b.Y) - eps;
            double maxAy = Math.Max(a.Y, b.Y) + eps;

            double minBx = Math.Min(c.X, d.X) - eps;
            double maxBx = Math.Max(c.X, d.X) + eps;
            double minBy = Math.Min(c.Y, d.Y) - eps;
            double maxBy = Math.Max(c.Y, d.Y) + eps;

            if (maxAx < minBx || maxBx < minAx) return false;
            if (maxAy < minBy || maxBy < minAy) return false;
            return true;
        }

        static bool IsPointOnSeg2D(XYZ p, XYZ a, XYZ b, double eps)
        {
            // коллинеарность + попадание в bbox
            XYZ ab = new XYZ(b.X - a.X, b.Y - a.Y, 0);
            XYZ ap = new XYZ(p.X - a.X, p.Y - a.Y, 0);
            double cross = Cross2D(ab, ap);
            if (Math.Abs(cross) > eps) return false;

            double minx = Math.Min(a.X, b.X) - eps;
            double maxx = Math.Max(a.X, b.X) + eps;
            double miny = Math.Min(a.Y, b.Y) - eps;
            double maxy = Math.Max(a.Y, b.Y) + eps;

            return (p.X >= minx && p.X <= maxx && p.Y >= miny && p.Y <= maxy);
        }

        static bool TryProjectParamOnSeg2D(XYZ p, XYZ a, XYZ b, double eps, out double t)
        {
            t = 0.0;
            XYZ ab = new XYZ(b.X - a.X, b.Y - a.Y, 0);
            double len2 = ab.X * ab.X + ab.Y * ab.Y;
            if (len2 < 1e-18) return false;

            XYZ ap = new XYZ(p.X - a.X, p.Y - a.Y, 0);
            t = Dot2D(ap, ab) / len2;
            if (t < -eps || t > 1.0 + eps) return false;
            return true;
        }

        static bool SegmentIntersect2D(XYZ a, XYZ b, XYZ c, XYZ d, double eps, out XYZ ip, out double ta, out double tb)
        {
            ip = null; ta = 0; tb = 0;

            XYZ r = new XYZ(b.X - a.X, b.Y - a.Y, 0);
            XYZ s = new XYZ(d.X - c.X, d.Y - c.Y, 0);
            double den = Cross2D(r, s);
            if (Math.Abs(den) < eps) return false; // параллельны (в т.ч. коллинеарны)

            XYZ ca = new XYZ(c.X - a.X, c.Y - a.Y, 0);
            ta = Cross2D(ca, s) / den;
            tb = Cross2D(ca, r) / den;

            if (ta < -eps || ta > 1.0 + eps) return false;
            if (tb < -eps || tb > 1.0 + eps) return false;

            double x = a.X + r.X * ta;
            double y = a.Y + r.Y * ta;
            ip = new XYZ(x, y, 0);
            return true;
        }

        static void AddSplitParam(List<double> pars, double t, double eps)
        {
            // дедуп по параметру
            for (int i = 0; i < pars.Count; i++)
                if (Math.Abs(pars[i] - t) <= eps)
                    return;
            pars.Add(t);
        }

        static List<Seg2> PlanarizeSegments2D(List<Seg2> input, double epsFt)
        {
            var result = new List<Seg2>();
            if (input == null || input.Count == 0) return result;

            int n = input.Count;

            // для каждого сегмента — список параметров разреза [0..1]
            var split = new List<double>[n];
            for (int i = 0; i < n; i++)
            {
                split[i] = new List<double>(8);
                split[i].Add(0.0);
                split[i].Add(1.0);
            }

            // попарно находим пересечения + накладки (коллинеарные overlap/T-junction)
            for (int i = 0; i < n; i++)
            {
                XYZ a = input[i].A; XYZ b = input[i].B;

                for (int j = i + 1; j < n; j++)
                {
                    XYZ c = input[j].A; XYZ d = input[j].B;

                    if (!BBoxOverlaps2D(a, b, c, d, epsFt))
                        continue;

                    // 1) обычное пересечение (непараллельные)
                    if (SegmentIntersect2D(a, b, c, d, epsFt, out var ip, out var ta, out var tb))
                    {
                        AddSplitParam(split[i], ta, 1e-9);
                        AddSplitParam(split[j], tb, 1e-9);
                        continue;
                    }

                    // 2) параллельные — проверяем коллинеарность и разрезаем по концам (overlap/T)
                    // коллинеарность: точки c и d лежат на линии ab
                    XYZ ab = new XYZ(b.X - a.X, b.Y - a.Y, 0);
                    double lenAb = Math.Sqrt(ab.X * ab.X + ab.Y * ab.Y);
                    if (lenAb < 1e-12) continue;

                    // расстояние через cross/|ab|
                    double distC = Math.Abs(Cross2D(ab, new XYZ(c.X - a.X, c.Y - a.Y, 0))) / lenAb;
                    double distD = Math.Abs(Cross2D(ab, new XYZ(d.X - a.X, d.Y - a.Y, 0))) / lenAb;
                    if (distC > epsFt || distD > epsFt)
                        continue; // параллельны, но не на одной прямой

                    // разрезаем i по концам j, если концы j попадают на i
                    if (TryProjectParamOnSeg2D(c, a, b, epsFt, out double tiC))
                        AddSplitParam(split[i], tiC, 1e-9);
                    if (TryProjectParamOnSeg2D(d, a, b, epsFt, out double tiD))
                        AddSplitParam(split[i], tiD, 1e-9);

                    // разрезаем j по концам i, если концы i попадают на j
                    if (TryProjectParamOnSeg2D(a, c, d, epsFt, out double tjA))
                        AddSplitParam(split[j], tjA, 1e-9);
                    if (TryProjectParamOnSeg2D(b, c, d, epsFt, out double tjB))
                        AddSplitParam(split[j], tjB, 1e-9);
                }
            }

            // собираем разрезанные под-сегменты
            var keys = new HashSet<string>();
            for (int i = 0; i < n; i++)
            {
                var s = input[i];
                XYZ a = s.A; XYZ b = s.B;

                XYZ ab = new XYZ(b.X - a.X, b.Y - a.Y, 0);
                double len = Math.Sqrt(ab.X * ab.X + ab.Y * ab.Y);
                if (len < 1e-12) continue;

                var pars = split[i];
                pars.Sort();

                for (int k = 0; k + 1 < pars.Count; k++)
                {
                    double t0 = pars[k];
                    double t1 = pars[k + 1];
                    if (t1 <= t0 + 1e-12) continue;

                    XYZ p0 = new XYZ(a.X + ab.X * t0, a.Y + ab.Y * t0, 0);
                    XYZ p1 = new XYZ(a.X + ab.X * t1, a.Y + ab.Y * t1, 0);

                    // SNAP + фильтр длины + дедуп
                    var P0 = SnapXY(p0);
                    var P1 = SnapXY(p1);
                    if (P0.DistanceTo(P1) < sMinSegFt) continue;

                    TryAddSeg2D(keys, result, P0, P1, s.IsCurtain, s.CurtainRootId);
                }
            }

            return result;
        }

        static int PG_AddNode(Dictionary<string, int> map, List<PGNode> nodes, XYZ p)
        {
            string key = KeyForPoint(p);
            if (map.TryGetValue(key, out int idx))
                return idx;

            idx = nodes.Count;
            nodes.Add(new PGNode { P = p });
            map[key] = idx;
            return idx;
        }

        static void PG_AddUndirectedEdge(
            List<PGNode> nodes, List<PGHalfEdge> hes,
            int a, int b)
        {
            if (a == b) return;

            // создаём 2 half-edge
            int i0 = hes.Count;
            int i1 = i0 + 1;

            XYZ pa = nodes[a].P;
            XYZ pb = nodes[b].P;

            double ang0 = Math.Atan2(pb.Y - pa.Y, pb.X - pa.X);
            double ang1 = Math.Atan2(pa.Y - pb.Y, pa.X - pb.X);

            hes.Add(new PGHalfEdge { From = a, To = b, Twin = i1, Angle = ang0, Used = false });
            hes.Add(new PGHalfEdge { From = b, To = a, Twin = i0, Angle = ang1, Used = false });

            nodes[a].Out.Add(i0);
            nodes[b].Out.Add(i1);
        }

        static void PG_SortOutgoing(List<PGNode> nodes, List<PGHalfEdge> hes)
        {
            foreach (var n in nodes)
            {
                n.Out.Sort((i, j) => hes[i].Angle.CompareTo(hes[j].Angle));
            }
        }

        static double PolySignedArea2D(List<XYZ> poly)
        {
            if (poly == null || poly.Count < 3) return 0.0;
            double a = 0.0;
            for (int i = 0; i < poly.Count; i++)
            {
                int j = (i + 1) % poly.Count;
                a += poly[i].X * poly[j].Y - poly[j].X * poly[i].Y;
            }
            return 0.5 * a;
        }

        static bool PointInPoly2D_Inclusive(XYZ p, List<XYZ> poly, double eps)
        {
            if (poly == null || poly.Count < 3) return false;

            // на ребре => inside
            for (int i = 0; i < poly.Count; i++)
            {
                int j = (i + 1) % poly.Count;
                if (IsPointOnSeg2D(p, poly[i], poly[j], eps))
                    return true;
            }

            // ray casting
            bool inside = false;
            for (int i = 0, j = poly.Count - 1; i < poly.Count; j = i++)
            {
                double xi = poly[i].X, yi = poly[i].Y;
                double xj = poly[j].X, yj = poly[j].Y;

                bool intersect = ((yi > p.Y) != (yj > p.Y)) &&
                                 (p.X < (xj - xi) * (p.Y - yi) / (yj - yi + 1e-30) + xi);
                if (intersect) inside = !inside;
            }
            return inside;
        }

        static List<List<XYZ>> ExtractFacesFromPlanarSegs(List<Seg2> planarSegs, double epsFt)
        {
            var faces = new List<List<XYZ>>();
            if (planarSegs == null || planarSegs.Count == 0) return faces;

            var nodes = new List<PGNode>();
            var map = new Dictionary<string, int>();
            var hes = new List<PGHalfEdge>();

            // добавляем рёбра
            // (дедуп по ключу, чтобы не плодить одинаковые)
            var edgeKey = new HashSet<string>();
            foreach (var s in planarSegs)
            {
                var A = SnapXY(s.A);
                var B = SnapXY(s.B);
                if (A.DistanceTo(B) < sMinSegFt) continue;

                string k = KeyFor(A, B);
                if (!edgeKey.Add(k)) continue;

                int ia = PG_AddNode(map, nodes, A);
                int ib = PG_AddNode(map, nodes, B);
                PG_AddUndirectedEdge(nodes, hes, ia, ib);
            }

            if (hes.Count == 0) return faces;

            PG_SortOutgoing(nodes, hes);

            // обход half-edge -> face (держим грань слева)
            int safetyLimit = hes.Count * 4;

            for (int heStart = 0; heStart < hes.Count; heStart++)
            {
                if (hes[heStart].Used) continue;

                var poly = new List<XYZ>(64);
                int he = heStart;
                int iter = 0;

                while (true)
                {
                    if (iter++ > safetyLimit) break;

                    var cur = hes[he];
                    if (cur.Used) break;

                    hes[he].Used = true;

                    if (poly.Count == 0)
                        poly.Add(nodes[cur.From].P);

                    poly.Add(nodes[cur.To].P);

                    int v = cur.To;
                    int twin = cur.Twin;

                    // найти позицию twin в исходящих v
                    var outList = nodes[v].Out;
                    int deg = outList.Count;
                    if (deg == 0) break;

                    int posTwin = -1;
                    for (int k = 0; k < deg; k++)
                    {
                        if (outList[k] == twin)
                        {
                            posTwin = k;
                            break;
                        }
                    }
                    if (posTwin < 0) break;

                    // next = ребро, которое идёт "слева": берём предыдущее (clockwise) от twin в CCW списке
                    int posNext = (posTwin - 1 + deg) % deg;
                    he = outList[posNext];

                    if (he == heStart)
                        break;
                }

                // убрать повтор последней точки, если совпала с первой
                if (poly.Count >= 2)
                {
                    string k0 = KeyForPoint(poly[0]);
                    string kN = KeyForPoint(poly[poly.Count - 1]);
                    if (k0 == kN)
                        poly.RemoveAt(poly.Count - 1);
                }

                if (poly.Count < 3) continue;

                double area = PolySignedArea2D(poly);
                if (Math.Abs(area) < 1e-9) continue;

                faces.Add(poly);
            }

            return faces;
        }
        static List<XYZ> PickMinFaceByPoint(List<List<XYZ>> faces, XYZ pTest, double epsFt)
        {
            if (faces == null || faces.Count == 0 || pTest == null) return null;

            List<XYZ> best = null;
            double bestArea = double.PositiveInfinity;

            foreach (var f in faces)
            {
                if (f == null || f.Count < 3) continue;
                if (!PointInPoly2D_Inclusive(pTest, f, epsFt)) continue;

                double a = Math.Abs(PolySignedArea2D(f));
                if (a < 1e-9) continue;

                if (a < bestArea)
                {
                    bestArea = a;
                    best = f;
                }
            }

            return best;
        }

        static bool TryGetRoomTestPointView(Room room, Transform T_model_to_view, out XYZ pV)
        {
            pV = null;
            if (room == null || T_model_to_view == null) return false;

            XYZ pM = null;

            try
            {
                if (room.Location is LocationPoint lp)
                    pM = lp.Point;
                else if (room.Location is LocationCurve lc && lc.Curve != null)
                    pM = lc.Curve.Evaluate(0.5, true);
            }
            catch { pM = null; }

            if (pM == null)
            {
                try
                {
                    var bb = room.get_BoundingBox(null);
                    if (bb != null) pM = 0.5 * (bb.Min + bb.Max);
                }
                catch { pM = null; }
            }

            if (pM == null) return false;

            var pv = T_model_to_view.OfPoint(pM);
            pV = new XYZ(pv.X, pv.Y, 0);
            return true;
        }

        static List<Seg2> SegsFromLoops(List<List<XYZ>> loops)
        {
            var segs = new List<Seg2>();
            if (loops == null) return segs;

            foreach (var loop in loops)
            {
                if (loop == null || loop.Count < 2) continue;
                int n = loop.Count;
                for (int i = 0; i < n; i++)
                {
                    var a = loop[i];
                    var b = loop[(i + 1) % n];
                    segs.Add(new Seg2(new XYZ(a.X, a.Y, 0), new XYZ(b.X, b.Y, 0), false, 0));
                }
            }
            return segs;
        }

        static HashSet<int> FindCurtainIdsTouchingRoom(
            List<Seg2> roomSegs,
            Dictionary<int, List<Seg2>> curtainSegsViewFt,
            double tolFt)
        {
            var ids = new HashSet<int>();
            if (roomSegs == null || roomSegs.Count == 0) return ids;
            if (curtainSegsViewFt == null || curtainSegsViewFt.Count == 0) return ids;

            for (int ci = 0; ci < roomSegs.Count; ci++)
            {
                // (оптимизация: bbox или grid можно позже, пока достаточно)
            }

            foreach (var kv in curtainSegsViewFt)
            {
                int curtainId = kv.Key;
                var list = kv.Value;
                if (list == null || list.Count == 0) continue;

                bool touch = false;

                foreach (var cs in list)
                {
                    var mid = new XYZ(0.5 * (cs.A.X + cs.B.X), 0.5 * (cs.A.Y + cs.B.Y), 0);

                    int bestIdx = -1;
                    double bestDist = double.PositiveInfinity;

                    for (int i = 0; i < roomSegs.Count; i++)
                    {
                        double d = DistPointToSegment2D(mid, roomSegs[i].A, roomSegs[i].B);
                        if (d < bestDist)
                        {
                            bestDist = d;
                            bestIdx = i;
                        }
                    }

                    if (bestIdx < 0 || bestDist > tolFt)
                        continue;

                    // параллельность (чтобы не схватить перпендикулярные)
                    var rs = roomSegs[bestIdx];
                    var rsDir = new XYZ(rs.B.X - rs.A.X, rs.B.Y - rs.A.Y, 0);
                    var csDir = new XYZ(cs.B.X - cs.A.X, cs.B.Y - cs.A.Y, 0);

                    if (!TryNormalize2D(rsDir, out var rsN)) continue;
                    if (!TryNormalize2D(csDir, out var csN)) continue;

                    double dot = Math.Abs(rsN.X * csN.X + rsN.Y * csN.Y);
                    if (dot < 0.85) // ослабил до 0.85, чтобы не отваливалось на “шуме”
                        continue;

                    touch = true;
                    break;
                }

                if (touch) ids.Add(curtainId);
            }

            return ids;
        }
        // --- OPA: корневой Id витражной стены для элемента границы помещения ---
        static int GetCurtainRootIdForElement(Element el)
        {
            if (el == null) return 0;

            try
            {
                if (el is Wall w)
                {
                    if (w.WallType != null && w.WallType.Kind == WallKind.Curtain)
                        return w.Id.IntegerValue;
                    return 0;
                }

                if (el is FamilyInstance fi)
                {
                    Wall hostWall = null;

                    if (fi.Host is Wall hw)
                        hostWall = hw;
                    else if (fi.SuperComponent is FamilyInstance sfi && sfi.Host is Wall shw)
                        hostWall = shw;

                    if (hostWall != null &&
                        hostWall.WallType != null &&
                        hostWall.WallType.Kind == WallKind.Curtain)
                        return hostWall.Id.IntegerValue;
                }
            }
            catch { }

            return 0;
        }

        // --- OPA: читаем boundary от Revit (полная петля как видит Revit) + отдельно сегменты НЕ-витража для графа ---
        static bool TryGetRoomBoundaryDataView(
            Room room,
            Document doc,
            SpatialElementBoundaryOptions bopts,
            Transform T_model_to_view,
            out List<List<XYZ>> loopsAllViewFt,
            out List<Seg2> segsNonCurtainViewFt,
            out HashSet<int> curtainIdsHint)
        {
            loopsAllViewFt = new List<List<XYZ>>();
            segsNonCurtainViewFt = new List<Seg2>();
            curtainIdsHint = new HashSet<int>();

            if (room == null || doc == null || bopts == null || T_model_to_view == null)
                return false;

            IList<IList<BoundarySegment>> loops = null;
            try { loops = room.GetBoundarySegments(bopts); }
            catch { loops = null; }

            if (loops == null || loops.Count == 0)
                return false;

            var keys = new HashSet<string>();

            foreach (var loop in loops)
            {
                if (loop == null || loop.Count == 0) continue;

                var ptsLoop = new List<XYZ>(128);
                XYZ lastP = null;

                foreach (var bs in loop)
                {
                    if (bs == null) continue;

                    Element boundEl = null;
                    try
                    {
                        if (bs.ElementId != null && bs.ElementId != ElementId.InvalidElementId)
                            boundEl = doc.GetElement(bs.ElementId);
                    }
                    catch { boundEl = null; }

                    bool isCurtainBound = IsCurtainElement(boundEl);
                    if (isCurtainBound)
                    {
                        int rid = GetCurtainRootIdForElement(boundEl);
                        if (rid != 0) curtainIdsHint.Add(rid);
                    }

                    Curve c = null;
                    try { c = bs.GetCurve(); } catch { c = null; }
                    if (c == null) continue;

                    // Тесселируем кривую boundary в точки (в порядке обхода)
                    IList<XYZ> tess = null;
                    if (c is Line ln)
                    {
                        tess = new List<XYZ> { ln.GetEndPoint(0), ln.GetEndPoint(1) };
                    }
                    else
                    {
                        try { tess = c.Tessellate(); }
                        catch { tess = null; }
                    }

                    if (tess == null || tess.Count < 2) continue;

                    // 1) Полная петля "как дал Revit" (для fallback/не-витражных)
                    for (int i = 0; i < tess.Count; i++)
                    {
                        var pv = T_model_to_view.OfPoint(tess[i]);
                        var p2 = SnapXY(new XYZ(pv.X, pv.Y, 0));

                        if (lastP != null && KeyForPoint(lastP) == KeyForPoint(p2))
                            continue;

                        ptsLoop.Add(p2);
                        lastP = p2;
                    }

                    // 2) Для графа: берём только НЕ-витражные сегменты boundary
                    if (!isCurtainBound)
                    {
                        for (int i = 0; i + 1 < tess.Count; i++)
                        {
                            var a = T_model_to_view.OfPoint(tess[i]);
                            var b = T_model_to_view.OfPoint(tess[i + 1]);
                            TryAddSeg2D(keys, segsNonCurtainViewFt, a, b, false, 0);
                        }
                    }
                }

                // убрать замыкание, если последняя == первая
                if (ptsLoop.Count >= 2 && KeyForPoint(ptsLoop[0]) == KeyForPoint(ptsLoop[ptsLoop.Count - 1]))
                    ptsLoop.RemoveAt(ptsLoop.Count - 1);

                if (ptsLoop.Count >= 3)
                    loopsAllViewFt.Add(ptsLoop);
            }

            return loopsAllViewFt.Count > 0 || segsNonCurtainViewFt.Count > 0;
        }

        // --- OPA: если boundary "рисует стенку вместо витража" — выкидываем сегменты boundary, которые близко и параллельно витражу ---
        static List<Seg2> RemoveRoomSegsNearCurtainParallel(
            List<Seg2> roomSegs,
            List<Seg2> curtainSegs,
            double tolFt)
        {
            if (roomSegs == null || roomSegs.Count == 0) return roomSegs;
            if (curtainSegs == null || curtainSegs.Count == 0) return roomSegs;

            var res = new List<Seg2>(roomSegs.Count);

            foreach (var rs in roomSegs)
            {
                var mid = new XYZ(0.5 * (rs.A.X + rs.B.X), 0.5 * (rs.A.Y + rs.B.Y), 0);

                var rsDir = new XYZ(rs.B.X - rs.A.X, rs.B.Y - rs.A.Y, 0);
                if (!TryNormalize2D(rsDir, out var rsN))
                {
                    res.Add(rs);
                    continue;
                }

                bool remove = false;

                foreach (var cs in curtainSegs)
                {
                    // близость
                    double d = DistPointToSegment2D(mid, cs.A, cs.B);
                    if (d > tolFt) continue;

                    // параллельность
                    var csDir = new XYZ(cs.B.X - cs.A.X, cs.B.Y - cs.A.Y, 0);
                    if (!TryNormalize2D(csDir, out var csN)) continue;

                    double dot = Math.Abs(rsN.X * csN.X + rsN.Y * csN.Y);
                    if (dot < 0.85) continue;

                    remove = true;
                    break;
                }

                if (!remove)
                    res.Add(rs);
            }

            return res;
        }

        // --- OPA: мостики от "открытых концов" room boundary до ближайшего сегмента витража ---
        static List<Seg2> BuildBridgesFromOpenEndsToCurtain(
            List<Seg2> roomSegs,
            List<Seg2> curtainSegs,
            double maxBridgeFt)
        {
            var bridges = new List<Seg2>();
            if (roomSegs == null || roomSegs.Count == 0) return bridges;
            if (curtainSegs == null || curtainSegs.Count == 0) return bridges;

            // считаем степени вершин (чтобы найти "разрывы" после удаления сегментов вдоль витража)
            var deg = new Dictionary<string, int>();
            var ptByKey = new Dictionary<string, XYZ>();

            void AddDeg(XYZ p)
            {
                string k = KeyForPoint(p);
                ptByKey[k] = p;
                if (!deg.ContainsKey(k)) deg[k] = 1;
                else deg[k] = deg[k] + 1;
            }

            foreach (var s in roomSegs)
            {
                AddDeg(s.A);
                AddDeg(s.B);
            }

            var keys = new HashSet<string>(); // дедуп мостиков

            foreach (var kv in deg)
            {
                if (kv.Value != 1) continue; // интересуют только "открытые" концы

                var p = ptByKey[kv.Key];

                XYZ qBest = null;
                double dBest = double.PositiveInfinity;

                foreach (var cs in curtainSegs)
                {
                    if (!ClosestPointOnSegment2D(p, cs.A, cs.B, out var q, out var t))
                        continue;

                    double dx = p.X - q.X, dy = p.Y - q.Y;
                    double d = Math.Sqrt(dx * dx + dy * dy);

                    if (d < dBest)
                    {
                        dBest = d;
                        qBest = q;
                    }
                }

                if (qBest == null) continue;
                if (dBest > maxBridgeFt) continue;
                if (dBest < sOpaBridgeEpsFt) continue; // почти ноль — не нужен

                var A = SnapXY(p);
                var B = SnapXY(qBest);
                if (A.DistanceTo(B) < sMinSegFt) continue;

                string kseg = KeyFor(A, B);
                if (!keys.Add(kseg)) continue;

                bridges.Add(new Seg2(A, B, false, 0));
            }

            return bridges;
        }

        // --- OPA: главный метод "room boundary (без витража) + полный витраж + мостики" -> минимальная грань с точкой ---
        static List<List<XYZ>> TryBuildRoomLoopOPA_BoundaryPlusCurtain(
            Room room,
            List<Seg2> roomBoundaryNonCurtainSegs,
            HashSet<int> curtainIdsHint,
            Dictionary<int, List<Seg2>> curtainSegsViewFt,
            double epsFt,
            Transform T_model_to_view)
        {
            if (room == null || curtainSegsViewFt == null || curtainSegsViewFt.Count == 0)
                return null;

            var ids = new HashSet<int>();
            if (curtainIdsHint != null) ids.UnionWith(curtainIdsHint);

            // ГЛАВНОЕ: допуск должен быть большим (иначе при "недолёте" boundary до витража мы не найдём витраж)
            double tolFt = sOpaCurtainReplaceMaxFt;
            if (roomBoundaryNonCurtainSegs != null && roomBoundaryNonCurtainSegs.Count > 0)
                ids.UnionWith(FindCurtainIdsTouchingRoom(roomBoundaryNonCurtainSegs, curtainSegsViewFt, tolFt));

            if (ids.Count == 0)
                return null; // витражей рядом нет — этот метод не нужен

            // собираем ПОЛНЫЕ сегменты витражей (как требование)
            var roomCurtainSegs = new List<Seg2>(2048);
            foreach (int id in ids)
            {
                if (curtainSegsViewFt.TryGetValue(id, out var list) && list != null && list.Count > 0)
                    roomCurtainSegs.AddRange(list);
            }
            if (roomCurtainSegs.Count == 0)
                return null;

            var roomSegs = roomBoundaryNonCurtainSegs ?? new List<Seg2>();
            if (roomSegs.Count == 0)
                return null;

            // 1) выкидываем сегменты boundary, которые Revit "рисует вместо витража" (иначе будет неправильная минимальная грань)
            var roomSegsCut = RemoveRoomSegsNearCurtainParallel(roomSegs, roomCurtainSegs, tolFt);

            // 2) дополнительно чистим "ступеньки" (у тебя уже есть готовая функция)
            roomSegsCut = RemoveCurtainStepsFromWalls(roomSegsCut, roomCurtainSegs);

            // 3) мостики (чтобы замкнуть граф между концами boundary и линией витража)
            var bridges = BuildBridgesFromOpenEndsToCurtain(roomSegsCut, roomCurtainSegs, tolFt);

            // 4) финальный набор сегментов для графа
            var all = new List<Seg2>(roomSegsCut.Count + roomCurtainSegs.Count + bridges.Count);
            all.AddRange(roomSegsCut);
            all.AddRange(roomCurtainSegs);
            all.AddRange(bridges);

            var planar = PlanarizeSegments2D(all, epsFt);
            if (planar == null || planar.Count == 0) return null;

            var faces = ExtractFacesFromPlanarSegs(planar, epsFt);
            if (faces == null || faces.Count == 0) return null;

            if (!TryGetRoomTestPointView(room, T_model_to_view, out var pTest))
                return null;

            var best = PickMinFaceByPoint(faces, pTest, epsFt);
            if (best == null || best.Count < 3) return null;

            return new List<List<XYZ>> { best };
        }

        static List<List<XYZ>> BuildRoomLoopsPlanarGraphForOPA(
            Room room,
            List<List<XYZ>> roomLoopsViewFt,
            Dictionary<int, List<Seg2>> curtainSegsViewFt,
            Transform T_model_to_view)
        {
            // 1) сегменты границы помещения
            var roomSegs = SegsFromLoops(roomLoopsViewFt);
            if (roomSegs.Count == 0) return roomLoopsViewFt;

            // 2) какие витражи реально касаются этого помещения
            double tolFt = UnitUtils.ConvertToInternalUnits(ROOM_CURTAIN_DIST_TOL_MM, UnitTypeId.Millimeters);
            var curtainIds = FindCurtainIdsTouchingRoom(roomSegs, curtainSegsViewFt, tolFt);

            if (curtainIds.Count == 0)
                return roomLoopsViewFt; // витражей рядом нет — оставляем как есть

            // 3) полный набор сегментов для планарного графа
            var all = new List<Seg2>(roomSegs.Count + 1024);
            all.AddRange(roomSegs);

            foreach (int id in curtainIds)
            {
                if (curtainSegsViewFt.TryGetValue(id, out var list) && list != null && list.Count > 0)
                    all.AddRange(list);
            }

            // 4) планаризация
            double epsFt = Math.Max(sKeySnapFt, UnitUtils.ConvertToInternalUnits(0.5, UnitTypeId.Millimeters));
            var planar = PlanarizeSegments2D(all, epsFt);
            if (planar == null || planar.Count == 0)
                return roomLoopsViewFt;

            // 5) faces
            var faces = ExtractFacesFromPlanarSegs(planar, epsFt);
            if (faces == null || faces.Count == 0)
                return roomLoopsViewFt;

            // 6) тестовая точка помещения
            if (!TryGetRoomTestPointView(room, T_model_to_view, out var pTest))
                return roomLoopsViewFt;

            // 7) выбираем минимальную по площади face, содержащую точку
            List<XYZ> best = null;
            double bestArea = double.PositiveInfinity;

            foreach (var f in faces)
            {
                if (!PointInPoly2D_Inclusive(pTest, f, epsFt))
                    continue;

                double a = Math.Abs(PolySignedArea2D(f));
                if (a < 1e-9) continue;

                if (a < bestArea)
                {
                    bestArea = a;
                    best = f;
                }
            }

            if (best == null || best.Count < 3)
                return roomLoopsViewFt;

            return new List<List<XYZ>> { best };
        }

        // Строим контуры помещения по solid'у комнаты через SpatialElementGeometryCalculator
        // 1) Берём solid комнаты
        // 2) Находим минимальный Z (низ помещения)
        // 3) Собираем все рёбра solid'а, которые лежат почти в этой плоскости
        // 4) Проецируем в координаты вида (T_model_to_view) и собираем петли
        static List<List<XYZ>> BuildRoomLoopsBySolid(
            Room room,
            SpatialElementGeometryCalculator calc,
            Transform T_model_to_view)
        {
            var result = new List<List<XYZ>>();
            if (room == null || calc == null || T_model_to_view == null)
                return result;

            SpatialElementGeometryResults geom;
            try
            {
                geom = calc.CalculateSpatialElementGeometry(room);
            }
            catch
            {
                return result;
            }

            if (geom == null)
                return result;

            var solid = geom.GetGeometry() as Solid;
            if (solid == null || solid.Edges == null || solid.Edges.Size == 0)
                return result;

            // Находим самый нижний Z по всем вершинам
            double zMin = double.PositiveInfinity;
            foreach (Edge e in solid.Edges)
            {
                IList<XYZ> pts = e.Tessellate();
                if (pts == null) continue;
                foreach (var p in pts)
                    if (p.Z < zMin) zMin = p.Z;
            }

            if (double.IsPositiveInfinity(zMin))
                return result;

            // Плоскость "пола" и допуск по Z
            double tolZ = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters); // ~1 мм
            double zPlane = zMin + tolZ * 0.5;

            var segs = new List<Seg2>();
            var keys = new HashSet<string>();

            // Собираем только те отрезки рёбер, которые лежат внизу помещения
            foreach (Edge e in solid.Edges)
            {
                IList<XYZ> pts = e.Tessellate();
                if (pts == null || pts.Count < 2) continue;

                for (int i = 0; i + 1 < pts.Count; ++i)
                {
                    var p = pts[i];
                    var q = pts[i + 1];

                    if (Math.Abs(p.Z - zPlane) > tolZ || Math.Abs(q.Z - zPlane) > tolZ)
                        continue;

                    // в координаты вида
                    var pV = T_model_to_view.OfPoint(p);
                    var qV = T_model_to_view.OfPoint(q);

                    // в 2D-сегменты (SnapXY, отсечение по длине, дедупликация уже внутри)
                    TryAddSeg2D(keys, segs, pV, qV, false, 0);
                }
            }

            // Из собранных сегментов строим замкнутые петли
            result = BuildLoopsFromSegs(segs);
            return result;
        }


        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            var uiapp = data.Application;
            var doc = uiapp.ActiveUIDocument != null ? uiapp.ActiveUIDocument.Document : null;
            if (doc == null) return Result.Failed;

            var active = doc.ActiveView;
            var sheet = active as ViewSheet;
            if (sheet == null)
            {
                TaskDialog.Show("Экспорт в JSON", "Открой лист и запусти команду с ЛИСТА.");
                return Result.Cancelled;
            }
            sDebug.Clear();
            DebugAdd("=== ExportCurvesCommand started ===");
            DebugAdd("Sheet: " + sheet.Name);
            // --- ВСЁ ДАЛЬНЕЙШЕЕ ВНУТРИ TransactionGroup, ЧТОБЫ В КОНЦЕ ОТКАТИТЬ ПИНЫ ---
            using (var tg = new TransactionGroup(doc, "Export Curves JSON (temp pin doors)"))
            {
                tg.Start();

                // список временно закреплённых дверей
                List<ElementId> doorsPinnedForExport = null;
                int pinnedDoorsTotal = 0;

                try
                {
                    // --- ПЕРЕД ЭКСПОРТОМ: временно закрепляем двереподобные экземпляры ---
                    using (var t = new Transaction(doc, "Pin all doors for export"))
                    {
                        if (t.Start() == TransactionStatus.Started)
                        {
                            doorsPinnedForExport = PinAllDoors(doc);
                            pinnedDoorsTotal = doorsPinnedForExport?.Count ?? 0;
                            t.Commit();
                        }
                    }

                    var tu = GetTargetLengthUnit(doc);
                    var targetUnitId = tu.uId;
                    var unitName = tu.name;

                    sMinSegFt = UnitUtils.ConvertToInternalUnits(MIN_SEG_MM, UnitTypeId.Millimeters);
                    sKeySnapFt = UnitUtils.ConvertToInternalUnits(DEDUP_SNAP_MM, UnitTypeId.Millimeters);

                    var catalog = BuildCatalogForSheet(doc, sheet);

                    var roomParamNames = BuildRoomParamCatalogSimple(doc);


                    var prev = LoadPrefs();
                    var win = new CutoutDialogWindow(catalog, roomParamNames, prev);


                    try { WindowInteropHelperEx.AttachToRevitMainWindow(win); } catch { }
                    var dlgRes = win.ShowDialog();
                    if (dlgRes != true)
                    {
                        tg.RollBack();
                        return Result.Cancelled;
                    }

                    var chosen = new CutoutPrefs
                    {
                        Enabled = win.EnableCutouts,
                        SystemShafts = win.UseSystemShafts,
                        CustomEnabled = win.UseCustom,
                        IncludeLinks = win.IncludeLinks,
                        Categories = win.UseCustom ? win.SelectedCategories.Select(b => (int)b).ToList() : new List<int>(),
                        Families = win.UseCustom ? win.SelectedFamilies : new List<string>(),
                        ExportMode = win.ExportMode,
                        MopParamName = win.MopParamName
                    };

                    SavePrefs(chosen);

                    // применить в рантайме
                    sCut_Enable = chosen.Enabled;
                    sCut_SystemShafts = chosen.SystemShafts;
                    sCut_Custom = chosen.CustomEnabled;
                    sCut_IncludeLinks = chosen.IncludeLinks;
                    sCut_Cats = new HashSet<BuiltInCategory>(chosen.Categories.Select(i => (BuiltInCategory)i));
                    sCut_Fams = new HashSet<string>(chosen.Families ?? new List<string>(), StringComparer.OrdinalIgnoreCase);

                    // режим расчёта: ГНС / Общая площадь
                    sModeGNS = (chosen.ExportMode == "GNS");
                    sModeOPA = !sModeGNS;
                    sOpaMopParamName = chosen.MopParamName ?? "";

                    // имена для meta
                    var metaCatNames = sCut_SystemShafts
                        ? new List<string> { "ShaftOpening" }
                        : (sCut_Cats.Count == 0 ? new List<string>() : sCut_Cats.Select(b => b.ToString().Replace("OST_", "")).ToList());
                    var metaFamNames = sCut_Fams.ToList();

                    // 3) путь сохранения
                    var outPath = AskJsonPath();
                    if (string.IsNullOrEmpty(outPath))
                    {
                        tg.RollBack();
                        return Result.Cancelled;
                    }

                    // сброс счётчиков шахт/вырезов
                    sCutElementUids = new HashSet<string>();
                    sCutElementCount = 0;

                    // сброс счётчиков дверей (бриджинг)
                    sDoorLikeSeenHost = 0;
                    sDoorLikeSeenLinks = 0;
                    sDoorLikeNoLocation = 0;
                    sDoorLikeBridged = 0;

                    var outGroups = new List<GroupPayload>();

                    // 4) Основная обработка
                    var vpIds = sheet.GetAllViewports();

                    int doorHostTotal = 0;
                    int doorLinkTotal = 0;
                    int doorNoLocation = 0;
                    int doorBridgedTotal = 0;

                    foreach (ElementId vpId in vpIds)
                    {
                        var vp = doc.GetElement(vpId) as Viewport;
                        if (vp == null) continue;
                        var v = doc.GetElement(vp.ViewId) as View;
                        DebugAdd($"--- Viewport {vpId.IntegerValue}, View='{v.Name}' ---");

                        if (v == null) continue;
                        if (!IsViewOK(doc, v)) continue;

                        if (!(v is ViewPlan) && v.ViewType != ViewType.CeilingPlan && v.ViewType != ViewType.AreaPlan)
                            continue;

                        // помещения для режима «Общая площадь»
                        List<RoomPayload> roomPayloads = null;

                        // Модель → XY вида
                        var T_model_to_view = (v.CropBox != null ? v.CropBox.Transform : Transform.Identity).Inverse;

                        var keysLocalHost = new HashSet<string>();
                        var segsViewFtHost = new List<Seg2>();
                        var keysLocalCut = new HashSet<string>();
                        var segsViewFtCut = new List<Seg2>();
                        Dictionary<int, List<Seg2>> curtainSegsViewFt = null; // все сегменты витражей по CurtainRootId
                        double opaEpsFt = Math.Max(sKeySnapFt, UnitUtils.ConvertToInternalUnits(0.5, UnitTypeId.Millimeters));


                        var doorReqs = new List<DoorBridgeInfo>(); // заявки на «зашивку» дверей (только host)
                        var allowedFilter = BuildAllowedFilterWithCutouts();

                        // ---------- ХОСТ ----------
                        var hostOpts = new Options { View = v, ComputeReferences = false, IncludeNonVisibleObjects = false };

                        var hostElems = new FilteredElementCollector(doc, v.Id)
                            .WherePasses(allowedFilter)
                            .WhereElementIsNotElementType()
                            .ToElements();

                        int cntModelCat = 0, cntBBoxNull = 0, cntGeomNull = 0, cntCurves = 0, cntSolids = 0, cntSegsAdded = 0;

                        foreach (var el in hostElems)
                        {
                            if (SkipAux(el)) continue;

                            // Системные шахты → cutouts
                            if (sCut_Enable && sCut_SystemShafts && IsSystemShaft(el))
                            {
                                CountCutElementOnce(el);
                                CollectShaftOpeningSegments(
                                    keysLocalCut, segsViewFtCut, el,
                                    T_model_to_view, null, ref cntSegsAdded);
                                continue;
                            }

                            bool isCut = ShouldGoToCutouts(el, false);

                            if (isCut) CountCutElementOnce(el);

                            var keysTarget = isCut ? keysLocalCut : keysLocalHost;
                            var segsTarget = isCut ? segsViewFtCut : segsViewFtHost;

                            bool isCurtainElement = IsCurtainElement(el);

                            // двери — «бриджинг»
                            if (el is FamilyInstance dfi && IsDoorLikeInstance(dfi))
                            {
                                // витражная дверь
                                if (IsCurtainDoorContext(dfi))
                                {
                                    if (!TryGetDoorCenterAndOrientation(
                                            dfi, doc, v,
                                            out var centerM, out var tM, out var nM, out var halfT))
                                    {
                                        doorNoLocation++;
                                    }
                                    else
                                    {
                                        var cV = T_model_to_view.OfPoint(centerM);
                                        var tV = T_model_to_view.OfVector(tM);
                                        var nV = T_model_to_view.OfVector(nM);

                                        if (TryNormalize2D(new XYZ(tV.X, tV.Y, 0), out var t2))
                                        {
                                            if (!TryNormalize2D(new XYZ(nV.X, nV.Y, 0), out var n2))
                                                n2 = new XYZ(-t2.Y, t2.X, 0);

                                            doorReqs.Add(new DoorBridgeInfo
                                            {
                                                Center = new XYZ(cV.X, cV.Y, 0),
                                                T = t2,
                                                N = n2,
                                                HalfT = halfT,
                                                Door = dfi,
                                                IsCurtain = true
                                            });

                                            continue;
                                        }
                                    }
                                }
                                // обычная дверь
                                else
                                {
                                    var lp = dfi.Location as LocationPoint;
                                    if (lp != null)
                                    {
                                        var wall = dfi.Host as Wall ?? FindNearestWallForDoor(doc, v, lp.Point);
                                        if (wall != null && GetWallDirsAtPoint(wall, lp.Point, out var tM, out var nM))
                                        {
                                            var cV = T_model_to_view.OfPoint(lp.Point);
                                            var tV = T_model_to_view.OfVector(tM);
                                            var nV = T_model_to_view.OfVector(nM);

                                            if (TryNormalize2D(new XYZ(tV.X, tV.Y, 0), out var t2))
                                            {
                                                if (!TryNormalize2D(new XYZ(nV.X, nV.Y, 0), out var n2))
                                                    n2 = new XYZ(-t2.Y, t2.X, 0);

                                                doorReqs.Add(new DoorBridgeInfo
                                                {
                                                    Center = new XYZ(cV.X, cV.Y, 0),
                                                    T = t2,
                                                    N = n2,
                                                    HalfT = 0.5 * wall.Width,
                                                    IsCurtain = false
                                                });

                                                continue;
                                            }
                                        }
                                    }
                                }
                            }

                            // остальная геометрия
                            CollectFromElementGeometry(
                                keysTarget,
                                segsTarget,
                                el,
                                hostOpts,
                                T_model_to_view,
                                null,
                                isCurtainElement && !isCut,
                                ref cntModelCat,
                                ref cntBBoxNull,
                                ref cntGeomNull,
                                ref cntCurves,
                                ref cntSolids,
                                ref cntSegsAdded);
                        }

                        // ---------- ССЫЛКИ ----------
                        var links = new FilteredElementCollector(doc, v.Id)
                            .OfClass(typeof(RevitLinkInstance))
                            .Cast<RevitLinkInstance>()
                            .ToList();

                        foreach (var link in links)
                        {
                            try
                            {
                                if (!IsElementVisibleInViewSafe(doc, v, link.Id))
                                    continue;
                            }
                            catch { }

                            var ldoc = link.GetLinkDocument();
                            if (ldoc == null) continue;

                            var linkElems = GetVisibleFromLinkSmart(doc, v, link, allowedFilter);

                            var linkOpts = new Options
                            {
                                ComputeReferences = false,
                                IncludeNonVisibleObjects = false,
                                DetailLevel = v.DetailLevel
                            };

                            var T_link_to_host = link.GetTotalTransform() ?? Transform.Identity;

                            int lCntModelCat = 0, lCntBBoxNull = 0, lCntGeomNull = 0, lCntCurves = 0, lCntSolids = 0, lCntSegsAdded = 0;

                            foreach (var lel in linkElems)
                            {
                                if (SkipAux(lel)) continue;

                                if (sCut_Enable && sCut_SystemShafts && IsSystemShaft(lel))
                                {
                                    CountCutElementOnce(lel);
                                    CollectShaftOpeningSegments(
                                        keysLocalCut, segsViewFtCut, lel,
                                        T_model_to_view, T_link_to_host,
                                        ref lCntSegsAdded);
                                    continue;
                                }

                                bool isCutL = ShouldGoToCutouts(lel, true);
                                if (isCutL) CountCutElementOnce(lel);

                                var keysTargetL = isCutL ? keysLocalCut : keysLocalHost;
                                var segsTargetL = isCutL ? segsViewFtCut : segsViewFtHost;

                                bool isCurtainElementL = IsCurtainElement(lel);

                                if (lel is FamilyInstance ldfi && IsDoorLikeInstance(ldfi))
                                {
                                    if (IsCurtainDoorContext(ldfi))
                                    {
                                        if (TryGetDoorCenterAndOrientation(
                                                ldfi, ldoc, null,
                                                out var centerM, out var tM, out var nM, out var halfT))
                                        {
                                            var cV = T_model_to_view.OfPoint(T_link_to_host.OfPoint(centerM));
                                            var tV = T_model_to_view.OfVector(T_link_to_host.OfVector(tM));
                                            var nV = T_model_to_view.OfVector(T_link_to_host.OfVector(nM));

                                            if (TryNormalize2D(new XYZ(tV.X, tV.Y, 0), out var t2))
                                            {
                                                if (!TryNormalize2D(new XYZ(nV.X, nV.Y, 0), out var n2))
                                                    n2 = new XYZ(-t2.Y, t2.X, 0);

                                                doorReqs.Add(new DoorBridgeInfo
                                                {
                                                    Center = new XYZ(cV.X, cV.Y, 0),
                                                    T = t2,
                                                    N = n2,
                                                    HalfT = halfT,
                                                    IsCurtain = true
                                                });

                                                continue;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        var lp = ldfi.Location as LocationPoint;
                                        if (lp != null)
                                        {
                                            var linkView = FirstUsablePlan(ldoc);
                                            var wall = ldfi.Host as Wall ?? FindNearestWallForDoor(ldoc, linkView, lp.Point);

                                            if (wall != null && GetWallDirsAtPoint(wall, lp.Point, out var tM, out var nM))
                                            {
                                                var cV = T_model_to_view.OfPoint(T_link_to_host.OfPoint(lp.Point));
                                                var tV = T_model_to_view.OfVector(T_link_to_host.OfVector(tM));
                                                var nV = T_model_to_view.OfVector(T_link_to_host.OfVector(nM));

                                                if (TryNormalize2D(new XYZ(tV.X, tV.Y, 0), out var t2))
                                                {
                                                    if (!TryNormalize2D(new XYZ(nV.X, nV.Y, 0), out var n2))
                                                        n2 = new XYZ(-t2.Y, t2.X, 0);

                                                    doorReqs.Add(new DoorBridgeInfo
                                                    {
                                                        Center = new XYZ(cV.X, cV.Y, 0),
                                                        T = t2,
                                                        N = n2,
                                                        HalfT = 0.5 * wall.Width,
                                                        IsCurtain = false
                                                    });

                                                    continue;
                                                }
                                            }
                                        }
                                    }
                                }

                                CollectFromElementGeometry(
                                    keysTargetL,
                                    segsTargetL,
                                    lel,
                                    linkOpts,
                                    T_model_to_view,
                                    T_link_to_host,
                                    isCurtainElementL && !isCutL,
                                    ref lCntModelCat,
                                    ref lCntBBoxNull,
                                    ref lCntGeomNull,
                                    ref lCntCurves,
                                    ref lCntSolids,
                                    ref lCntSegsAdded);
                            }
                        }



                        // --- «зашивка» дверей (в координатах вида, только host) ---
                        if (doorReqs.Count > 0 && segsViewFtHost.Count > 0)
                        {
                            int addedByDoors = 0;

                            var curtainSegs = segsViewFtHost
                                .Where(s => s.IsCurtain)
                                .ToList();

                            foreach (var d in doorReqs)
                            {
                                if (!d.IsCurtain)
                                {
                                    var qOuter = new XYZ(d.Center.X + d.N.X * d.HalfT,
                                                         d.Center.Y + d.N.Y * d.HalfT, 0);
                                    var qInner = new XYZ(d.Center.X - d.N.X * d.HalfT,
                                                         d.Center.Y - d.N.Y * d.HalfT, 0);

                                    var hitO = NearestHitsAlongLine(qOuter, d.T, curtainSegs);
                                    var hitI = NearestHitsAlongLine(qInner, d.T, curtainSegs);

                                    if (!hitO.ok || !hitI.ok)
                                    {
                                        var hc = NearestHitsAlongLine(d.Center, d.T, segsViewFtHost);
                                        if (hc.ok)
                                        {
                                            var pNegO = new XYZ(hc.pNeg.X + d.N.X * d.HalfT,
                                                                hc.pNeg.Y + d.N.Y * d.HalfT, 0);
                                            var pPosO = new XYZ(hc.pPos.X + d.N.X * d.HalfT,
                                                                hc.pPos.Y + d.N.Y * d.HalfT, 0);
                                            var pNegI = new XYZ(hc.pNeg.X - d.N.X * d.HalfT,
                                                                hc.pNeg.Y - d.N.Y * d.HalfT, 0);
                                            var pPosI = new XYZ(hc.pPos.X - d.N.X * d.HalfT,
                                                                hc.pPos.Y - d.N.Y * d.HalfT, 0);

                                            AddSeg2D(keysLocalHost, segsViewFtHost, pNegO, pPosO, ref addedByDoors);
                                            AddSeg2D(keysLocalHost, segsViewFtHost, pNegI, pPosI, ref addedByDoors);
                                            continue;
                                        }
                                    }

                                    if (hitO.ok)
                                        AddSeg2D(keysLocalHost, segsViewFtHost, hitO.pNeg, hitO.pPos, ref addedByDoors);
                                    if (hitI.ok)
                                        AddSeg2D(keysLocalHost, segsViewFtHost, hitI.pNeg, hitI.pPos, ref addedByDoors);

                                    continue;
                                }
                                else
                                {
                                    if (TryGetCurtainPanelOffsetsFromSegments(d, curtainSegs, out double innerOff, out double outerOff))
                                    {
                                        var qOuter = d.Center + d.N.Multiply(outerOff);
                                        var hitO = NearestHitsAlongLine(qOuter, d.T, segsViewFtHost);

                                        var qInner = d.Center + d.N.Multiply(innerOff);
                                        var hitI = NearestHitsAlongLine(qInner, d.T, segsViewFtHost);

                                        if (!hitO.ok || !hitI.ok)
                                        {
                                            var hc = NearestHitsAlongLine(d.Center, d.T, curtainSegs);
                                            if (hc.ok)
                                            {
                                                if (!hitO.ok)
                                                {
                                                    var pNegO = hc.pNeg + d.N.Multiply(outerOff);
                                                    var pPosO = hc.pPos + d.N.Multiply(outerOff);
                                                    AddSeg2D(keysLocalHost, segsViewFtHost, pNegO, pPosO, ref addedByDoors);
                                                }

                                                if (!hitI.ok)
                                                {
                                                    var pNegI = hc.pNeg + d.N.Multiply(innerOff);
                                                    var pPosI = hc.pPos + d.N.Multiply(innerOff);
                                                    AddSeg2D(keysLocalHost, segsViewFtHost, pNegI, pPosI, ref addedByDoors);
                                                }

                                                if (!hitO.ok && !hitI.ok)
                                                    continue;
                                            }
                                        }

                                        if (hitO.ok)
                                            AddSeg2D(keysLocalHost, segsViewFtHost, hitO.pNeg, hitO.pPos, ref addedByDoors);

                                        if (hitI.ok)
                                            AddSeg2D(keysLocalHost, segsViewFtHost, hitI.pNeg, hitI.pPos, ref addedByDoors);

                                        continue;
                                    }

                                    double shiftOuterFt = UnitUtils.ConvertToInternalUnits(
                                        DOOR_OUTER_SHIFT_MM, UnitTypeId.Millimeters);

                                    var qOuterOld = d.Center + d.N.Multiply(d.HalfT + shiftOuterFt);
                                    var qInnerOld = d.Center - d.N.Multiply(d.HalfT);

                                    var hitOold = NearestHitsAlongLine(qOuterOld, d.T, segsViewFtHost);
                                    var hitIold = NearestHitsAlongLine(qInnerOld, d.T, segsViewFtHost);

                                    if (!hitOold.ok || !hitIold.ok)
                                    {
                                        var hc = NearestHitsAlongLine(d.Center, d.T, segsViewFtHost);
                                        if (hc.ok)
                                        {
                                            var pNegO = hc.pNeg + d.N.Multiply(d.HalfT + shiftOuterFt);
                                            var pPosO = hc.pPos + d.N.Multiply(d.HalfT + shiftOuterFt);

                                            var pNegI = hc.pNeg - d.N.Multiply(d.HalfT);
                                            var pPosI = hc.pPos - d.N.Multiply(d.HalfT);

                                            AddSeg2D(keysLocalHost, segsViewFtHost, pNegO, pPosO, ref addedByDoors);
                                            AddSeg2D(keysLocalHost, segsViewFtHost, pNegI, pPosI, ref addedByDoors);
                                            continue;
                                        }
                                    }

                                    if (hitOold.ok)
                                    {
                                        var pNegO = hitOold.pNeg + d.N.Multiply(shiftOuterFt);
                                        var pPosO = hitOold.pPos + d.N.Multiply(shiftOuterFt);
                                        AddSeg2D(keysLocalHost, segsViewFtHost, pNegO, pPosO, ref addedByDoors);
                                    }

                                    if (hitIold.ok)
                                    {
                                        AddSeg2D(keysLocalHost, segsViewFtHost, hitIold.pNeg, hitIold.pPos, ref addedByDoors);
                                    }
                                }
                            }
                        }
                        // --- ПРОСТО: собираем все сегменты витражей по CurtainRootId (для вывода 1:1) ---
                        if (sModeOPA && segsViewFtHost.Count > 0)
                        {
                            curtainSegsViewFt = new Dictionary<int, List<Seg2>>();

                            foreach (var s in segsViewFtHost)
                            {
                                // интересуют только сегменты, помеченные как vitrazh
                                if (!s.IsCurtain || s.CurtainRootId == 0)
                                    continue;

                                if (!curtainSegsViewFt.TryGetValue(s.CurtainRootId, out var list))
                                {
                                    list = new List<Seg2>();
                                    curtainSegsViewFt[s.CurtainRootId] = list;
                                }

                                // добавляем сегменты РОВНО в том порядке, как они стоят в общем списке
                                list.Add(s);
                            }
                        }
                       
                        // ---------- ПОМЕЩЕНИЯ (Rooms) ДЛЯ РЕЖИМА ОПА ----------
                        if (sModeOPA)
                        {
                            roomPayloads = new List<RoomPayload>();

                            // Один калькулятор на вид
                            var geomOpts = new SpatialElementBoundaryOptions();
                            // по желанию можешь настроить:
                            // geomOpts.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;
                            // geomOpts.StoreFreeBoundaryFaces = false;

                            var spatialCalc = new SpatialElementGeometryCalculator(doc, geomOpts);

                            var roomCollector = new FilteredElementCollector(doc, v.Id)
                                .OfCategory(BuiltInCategory.OST_Rooms)
                                .WhereElementIsNotElementType()
                                .Cast<Room>();

                            foreach (var room in roomCollector)
                            {
                                // пропускаем помещения нулевой площади
                                try
                                {
                                    if (room.Area <= 0)
                                        continue;
                                }
                                catch
                                {
                                    // если Revit не дал прочитать Area – просто пробуем дальше
                                }
                                if (IsRoomMop(room))
                                {
                                    DebugAdd($"[OPA] Skip MOP room: id={room.Id.IntegerValue}, num='{room.Number}', name='{room.Name}', " +
                                             $"{sOpaMopParamName}='{GetRoomParamValueByName(room, sOpaMopParamName)}'");
                                    continue;
                                }
                                List<List<XYZ>> loopsForRoom = null;

                                // boundary options — именно Room boundary от Revit
                                var bopts = new SpatialElementBoundaryOptions();
                                // если нужно именно по внутренней отделке:
                                bopts.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;

                                if (!TryGetRoomBoundaryDataView(room, doc, bopts, T_model_to_view,
                                        out var roomLoopsRevitViewFt,
                                        out var roomSegsNonCurtainViewFt,
                                        out var curtainIdsHint))
                                {
                                    continue;
                                }

                                // 1) Если рядом есть витражи — строим граф только из (boundary без витража) + (полный витраж)
                                loopsForRoom = TryBuildRoomLoopOPA_BoundaryPlusCurtain(
                                    room,
                                    roomSegsNonCurtainViewFt,
                                    curtainIdsHint,
                                    curtainSegsViewFt,
                                    opaEpsFt,
                                    T_model_to_view);

                                // 2) Если витражей нет ИЛИ граф не смог собрать корректную грань — берём boundary как дал Revit
                                if (loopsForRoom == null || loopsForRoom.Count == 0)
                                {
                                    if (roomLoopsRevitViewFt != null && roomLoopsRevitViewFt.Count > 0)
                                        loopsForRoom = roomLoopsRevitViewFt;
                                }

                                // 3) Последний фоллбек (как было) — по solid
                                if (loopsForRoom == null || loopsForRoom.Count == 0)
                                {
                                    var roomLoopsView = BuildRoomLoopsBySolid(room, spatialCalc, T_model_to_view);
                                    if (roomLoopsView == null || roomLoopsView.Count == 0)
                                        continue;

                                    loopsForRoom = roomLoopsView;
                                }

                                roomPayloads.Add(rp);

                            }

                        }


                        // --- удалить из segments дубликаты cutouts ---
                        if (sCut_Enable && segsViewFtCut.Count > 0 && segsViewFtHost.Count > 0)
                        {
                            var cutKeys = new HashSet<string>();
                            foreach (var seg in segsViewFtCut)
                                cutKeys.Add(KeyFor(SnapXY(seg.A), SnapXY(seg.B)));

                            segsViewFtHost = segsViewFtHost
                                .Where(seg => !cutKeys.Contains(KeyFor(SnapXY(seg.A), SnapXY(seg.B))))
                                .ToList();
                        }

                        // Центр размещения вьюпорта
                        var centerView = GetViewportCenterInViewCoords(v);

                        // Вид → выходное пространство
                        var Tv2s = BuildViewToOutTransform(v, vp, centerView, OUTPUT_SCALE_MODE, UNIFIED_SCALE);

                        // Перевод в конечные единицы
                        var segsViewOutHost = new List<Seg2>(segsViewFtHost.Count);
                        var segsViewOutCut = new List<Seg2>(segsViewFtCut.Count);
                        Func<double, double> FromFt = ft => UnitUtils.ConvertFromInternalUnits(ft, targetUnitId);

                        foreach (var seg in segsViewFtHost)
                        {
                            var As = Tv2s.OfPoint(seg.A);
                            var Bs = Tv2s.OfPoint(seg.B);
                            var a2 = new XYZ(RoundHalfUp(FromFt(As.X), ROUND_DEC), RoundHalfUp(FromFt(As.Y), ROUND_DEC), 0);
                            var b2 = new XYZ(RoundHalfUp(FromFt(Bs.X), ROUND_DEC), RoundHalfUp(FromFt(Bs.Y), ROUND_DEC), 0);
                            segsViewOutHost.Add(new Seg2(a2, b2));
                        }
                        foreach (var seg in segsViewFtCut)
                        {
                            var As = Tv2s.OfPoint(seg.A);
                            var Bs = Tv2s.OfPoint(seg.B);
                            var a2 = new XYZ(RoundHalfUp(FromFt(As.X), ROUND_DEC), RoundHalfUp(FromFt(As.Y), ROUND_DEC), 0);
                            var b2 = new XYZ(RoundHalfUp(FromFt(Bs.X), ROUND_DEC), RoundHalfUp(FromFt(Bs.Y), ROUND_DEC), 0);
                            segsViewOutCut.Add(new Seg2(a2, b2));
                        }
                        // НОВОЕ: отдельные сегменты витражей в выходных координатах (для групп по витражам)
                        Dictionary<int, List<Seg2>> curtainSegsOut = null;
                        if (curtainSegsViewFt != null && curtainSegsViewFt.Count > 0)
                        {
                            curtainSegsOut = new Dictionary<int, List<Seg2>>();

                            foreach (var kv in curtainSegsViewFt)
                            {
                                int curtainId = kv.Key;
                                var listIn = kv.Value;
                                var listOut = new List<Seg2>(listIn.Count);

                                foreach (var seg in listIn)
                                {
                                    var As = Tv2s.OfPoint(seg.A);
                                    var Bs = Tv2s.OfPoint(seg.B);

                                    var a2 = new XYZ(
                                        RoundHalfUp(FromFt(As.X), ROUND_DEC),
                                        RoundHalfUp(FromFt(As.Y), ROUND_DEC),
                                        0);
                                    var b2 = new XYZ(
                                        RoundHalfUp(FromFt(Bs.X), ROUND_DEC),
                                        RoundHalfUp(FromFt(Bs.Y), ROUND_DEC),
                                        0);

                                    listOut.Add(new Seg2(a2, b2));
                                }

                                curtainSegsOut[curtainId] = listOut;
                            }
                        }

                        // Преобразуем контуры помещений
                        if (roomPayloads != null)
                        {
                            foreach (var rp in roomPayloads)
                            {
                                var newLoops = new List<List<XYZ>>();

                                if (rp.loops != null)
                                {
                                    foreach (var loop in rp.loops)
                                    {
                                        if (loop == null || loop.Count == 0) continue;

                                        var newLoop = new List<XYZ>();
                                        foreach (var p in loop)
                                        {
                                            var ps = Tv2s.OfPoint(p);
                                            var p2 = new XYZ(
                                                RoundHalfUp(FromFt(ps.X), ROUND_DEC),
                                                RoundHalfUp(FromFt(ps.Y), ROUND_DEC),
                                                0);
                                            newLoop.Add(p2);
                                        }
                                        if (newLoop.Count > 0)
                                            newLoops.Add(newLoop);
                                    }
                                }

                                rp.loops = newLoops;



                            }
                        }

                        outGroups.Add(new GroupPayload
                        {
                            id = "VP_" + vpId.IntegerValue,
                            name = v.Name,
                            source = "host",
                            segs = segsViewOutHost,
                            cutouts = segsViewOutCut,
                            rooms = roomPayloads ?? new List<RoomPayload>()
                        });
                        // НОВОЕ: отдельные группы по каждому витражу (все сегменты витража 1:1)
                        if (sModeOPA && curtainSegsOut != null && curtainSegsOut.Count > 0)
                        {
                            foreach (var kv in curtainSegsOut)
                            {
                                int curtainId = kv.Key;
                                var segsForCurtain = kv.Value ?? new List<Seg2>();

                                outGroups.Add(new GroupPayload
                                {
                                    id = $"VP_{vpId.IntegerValue}_CURTAIN_{curtainId}",
                                    name = v.Name + $" (Curtain {curtainId})",
                                    source = "curtain",
                                    segs = segsForCurtain,
                                    cutouts = new List<Seg2>(),
                                    rooms = new List<RoomPayload>()
                                });
                            }
                        }

                    } // foreach viewport

                    // запись JSON
                    try
                    {
                        WriteJsonMulti(
                            outPath,
                            outGroups,
                            sheet.Name,
                            unitName,
                            metaCatNames,
                            metaFamNames,
                            sCut_IncludeLinks,
                            sCut_SystemShafts,
                            sCut_Enable
                        );
                    }
                    catch (Exception ex)
                    {
                        tg.RollBack();
                        TaskDialog.Show("Экспорт в JSON", "Ошибка записи JSON:\n" + ex.Message);
                        return Result.Failed;
                    }

                    // --- ПОСЛЕ ЭКСПОРТА: откатить все временные изменения в документе ---
                    // (в том числе транзакцию "Pin all doors for export")
                    tg.RollBack();

                    // пишем полный лог в файл рядом с JSON
                    string debugPath = null;
                    try
                    {
                        debugPath = Path.ChangeExtension(outPath, ".debug.txt");
                        File.WriteAllText(debugPath, sDebug.ToString(), new UTF8Encoding(false));
                    }
                    catch
                    {
                        debugPath = null;
                    }

                    // обрезаем лог для диалога, чтобы не переборщить
                    string debugForDialog = sDebug.ToString();
                    if (!string.IsNullOrEmpty(debugForDialog) && debugForDialog.Length > 2000)
                    {
                        debugForDialog = debugForDialog.Substring(0, 2000) + "\n...(обрезано)";
                    }

                    string dlgText =
                        "Готово ✅\n\n" +
                        "Лист: " + sheet.Name + "\n" +
                        "Групп (видов): " + outGroups.Count + "\n" +
                        "Шахт (вырезанных элементов): " + sCutElementCount + "\n" +
                        "Дверей (хост): " + doorHostTotal + "\n" +
                        "Дверей (ссылки): " + doorLinkTotal + "\n" +
                        "Дверей без точки/ориентации: " + doorNoLocation + "\n" +
                        "Успешно \"зашитых\" дверей: " + doorBridgedTotal + "\n\n" +
                        "Временно закреплено дверей (хост): " + pinnedDoorsTotal + "\n\n" +
                        "Файл сохранён:\n" + outPath;

                    if (!string.IsNullOrEmpty(debugPath))
                    {
                        dlgText += "\n\nDebug-лог сохранён:\n" + debugPath;
                    }

                    if (!string.IsNullOrEmpty(debugForDialog))
                    {
                        dlgText += "\n\n--- DEBUG (первые 2000 символов) ---\n" + debugForDialog;
                    }

                    TaskDialog.Show("Export Curves (JSON) – Экспорт в JSON", dlgText);


                    return Result.Succeeded;


                }
                catch (Exception ex)
                {
                    tg.RollBack();
                    TaskDialog.Show("Экспорт в JSON", "Ошибка:\n" + ex.Message);
                    return Result.Failed;
                }
            }
        }
    }

    // Привязка WPF-окна к главному окну Revit (без XAML)
    internal static class WindowInteropHelperEx
    {
        public static void AttachToRevitMainWindow(Wpf.Window w)
        {
            try
            {
                var helper = new System.Windows.Interop.WindowInteropHelper(w);
                helper.Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            }
            catch { }
        }
    }
}