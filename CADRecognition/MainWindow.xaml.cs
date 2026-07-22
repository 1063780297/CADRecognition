using Microsoft.Win32;
using netDxf;
using netDxf.Entities;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using ComboBox = System.Windows.Controls.ComboBox;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;
using Line = netDxf.Entities.Line;
using WinOpenFileDialog = Microsoft.Win32.OpenFileDialog;
using WpfBrushes = System.Windows.Media.Brushes;
using WpfColor = System.Windows.Media.Color;
using WpfEllipse = System.Windows.Shapes.Ellipse;
using WpfLine = System.Windows.Shapes.Line;
using WpfRectangle = System.Windows.Shapes.Rectangle;
using CADImport;
using HslCommunication.ModBus;
using Path = System.Windows.Shapes.Path;
using Brushes = System.Windows.Media.Brushes;
using FormsFolderBrowserDialog = System.Windows.Forms.FolderBrowserDialog;
using System.Text;

namespace CADRecognition
{
    internal static class DxfSafe
    {
        public static IEnumerable<Line> Lines(DxfDocument doc) => doc?.Entities?.Lines ?? Enumerable.Empty<Line>();
        public static IEnumerable<Circle> Circles(DxfDocument doc) => doc?.Entities?.Circles ?? Enumerable.Empty<Circle>();
        public static IEnumerable<Arc> Arcs(DxfDocument doc) => doc?.Entities?.Arcs ?? Enumerable.Empty<Arc>();
        public static IEnumerable<Polyline2D> Polylines2D(DxfDocument doc) => doc?.Entities?.Polylines2D ?? Enumerable.Empty<Polyline2D>();
        public static IEnumerable<Insert> Inserts(DxfDocument doc) => doc?.Entities?.Inserts ?? Enumerable.Empty<Insert>();
    }

    internal static class CadDocumentLoader
    {
        public static DxfDocument Load(string path)
        {
            var ext = System.IO.Path.GetExtension(path);
            return string.Equals(ext, ".dwg", StringComparison.OrdinalIgnoreCase)
                ? LoadDwg(path)
                : DxfDocument.Load(path);
        }

        private static DxfDocument LoadDwg(string path)
        {
            using var editor = new CADImport.CADImportControls.CADEditorControl();
            editor.LoadFile(path);

            var image = editor.Image;
            if (image is null)
            {
                throw new InvalidOperationException("DWG 读取失败：CADImport 未返回图像对象。请确认图纸版本和运行库支持。");
            }

            var entities = ExtractAllImportEntities(image).ToList();
            if (entities.Count == 0)
            {
                throw new InvalidOperationException("DWG 读取成功，但未解析出任何实体。请检查图纸是否包含有效模型空间内容或 CADImport 是否支持该图纸结构。");
            }

            var doc = new DxfDocument();
            foreach (var entity in entities)
            {
                AddImportedEntity(doc, entity);
            }

            if (doc.Entities == null)
            {
                throw new InvalidOperationException("DWG 已打开，但未转换出可绘制的 DXF 几何实体。请检查图纸内容或 CADImport 支持范围。");
            }

            return doc;
        }

        private sealed class ReferenceComparer : IEqualityComparer<object>
        {
            public new bool Equals(object? x, object? y) => ReferenceEquals(x, y);
            public int GetHashCode(object obj) => RuntimeHelpers.GetHashCode(obj);
        }

        private static IEnumerable<object> ExtractAllImportEntities(object image)
        {
            var seen = new HashSet<object>(new ReferenceComparer());
            var result = new List<object>();

            void AddRange(IEnumerable<object> source)
            {
                foreach (var item in source)
                {
                    if (item is null || !seen.Add(item))
                    {
                        continue;
                    }

                    result.Add(item);
                }
            }

            var currentLayout = image.GetType().GetProperty("CurrentLayout")?.GetValue(image);
            if (currentLayout is not null)
            {
                AddRange(ExtractLayoutEntities(currentLayout));
                AddRange(ExtractBlockEntities(currentLayout.GetType().GetProperty("PaperSpaceBlock")?.GetValue(currentLayout)));
            }

            var layoutsValue = image.GetType().GetProperty("Layouts")?.GetValue(image);
            if (layoutsValue is System.Collections.IEnumerable enumerable && layoutsValue is not string)
            {
                foreach (var layout in enumerable.Cast<object>())
                {
                    AddRange(ExtractLayoutEntities(layout));
                    AddRange(ExtractBlockEntities(layout.GetType().GetProperty("PaperSpaceBlock")?.GetValue(layout)));
                }
            }

            var imageLayout = image.GetType().GetProperty("Layout")?.GetValue(image);
            if (imageLayout is not null)
            {
                AddRange(ExtractLayoutEntities(imageLayout));
                AddRange(ExtractBlockEntities(imageLayout.GetType().GetProperty("PaperSpaceBlock")?.GetValue(imageLayout)));
            }

            return result;
        }

        private static IEnumerable<object> ExtractLayoutEntities(object? layout)
        {
            if (layout is null)
            {
                return Enumerable.Empty<object>();
            }

            var layoutType = layout.GetType();
            var entitiesProp = layoutType.GetProperty("Entities") ?? layoutType.GetProperty("EntityList") ?? layoutType.GetProperty("Items");
            var value = entitiesProp?.GetValue(layout);
            if (value is null)
            {
                var countProp = layoutType.GetProperty("Count");
                if (countProp?.GetValue(layout) is int count && count > 0)
                {
                    var itemProp = layoutType.GetProperty("Item");
                    if (itemProp is not null)
                    {
                        var items = new List<object>();
                        for (var i = 0; i < count; i++)
                        {
                            try
                            {
                                var item = itemProp.GetValue(layout, new object[] { i });
                                if (item is not null)
                                {
                                    items.Add(item);
                                }
                            }
                            catch
                            {
                            }
                        }

                        return items;
                    }
                }

                return Enumerable.Empty<object>();
            }

            if (value is System.Collections.IEnumerable enumerable && value is not string)
            {
                return enumerable.Cast<object>().Where(x => x is not null);
            }

            return Enumerable.Empty<object>();
        }

        private static IEnumerable<object> ExtractBlockEntities(object? block)
        {
            if (block is null)
            {
                return Enumerable.Empty<object>();
            }

            var blockType = block.GetType();
            var entitiesProp = blockType.GetProperty("Entities") ?? blockType.GetProperty("Items");
            var value = entitiesProp?.GetValue(block);
            if (value is System.Collections.IEnumerable enumerable && value is not string)
            {
                return enumerable.Cast<object>().Where(x => x is not null);
            }

            return Enumerable.Empty<object>();
        }

        private static void AddImportedEntity(DxfDocument doc, object entity)
        {
            switch (entity)
            {
                case CADEllipse ellipse:
                    AddEllipseAsPolyline(doc, ellipse);
                    break;

                case CADLine line:
                    AddLine(doc, ToPoint(line.Point), ToPoint(line.Point1));
                    break;

                case CAD2DLine line2D:
                    AddLine(doc, ToPoint(line2D.StartPoint), ToPoint(line2D.EndPoint));
                    break;

                case CADArc arc:
                    AddArc(doc, arc);
                    break;

                case CADCircle circle:
                    AddCircle(doc, circle.Point.X, circle.Point.Y, circle.Radius);
                    break;

                case CADLWPolyLine lwPoly:
                    AddPolyline(doc, ExpandCadLwPolyline(lwPoly), IsClosedLike(lwPoly));
                    break;

                case CADPolyLine polyLine when polyLine.GetType() == typeof(CADPolyLine):
                    AddPolyline(doc, ExtractPoints(polyLine), IsClosedLike(polyLine));
                    break;

                case CAD2DPolyline poly2D:
                    AddPolyline(doc, ExtractPoints(poly2D), IsClosedLike(poly2D));
                    break;

                case CADInsert insert:
                    foreach (var nested in ExtractInsertEntities(insert))
                    {
                        AddImportedEntity(doc, nested);
                    }
                    break;
            }
        }

        private static void AddLine(DxfDocument doc, (double X, double Y) start, (double X, double Y) end)
        {
            if (DistanceSquared(start, end) <= 1e-12)
            {
                return;
            }

            doc.Entities.Add(new Line(
                new Vector3(start.X, start.Y, 0),
                new Vector3(end.X, end.Y, 0)));
        }

        private static void AddCircle(DxfDocument doc, double x, double y, double radius)
        {
            if (radius <= 0)
            {
                return;
            }

            doc.Entities.Add(new Circle(new Vector3(x, y, 0), radius));
        }

        private static void AddArc(DxfDocument doc, CADArc arc)
        {
            var center = new Vector3(arc.Point.X, arc.Point.Y, 0);
            var radius = Convert.ToDouble(arc.Radius, CultureInfo.InvariantCulture);
            var startAngle = TryGetArcAngle(arc, "StartParam", "StartAngle", "Start")
                ?? TryGetAngleFromPoint(center.X, center.Y, arc, "StartPoint", "FromPoint", "P1");
            var endAngle = TryGetArcAngle(arc, "EndParam", "EndAngle", "End")
                ?? TryGetAngleFromPoint(center.X, center.Y, arc, "EndPoint", "ToPoint", "P2");

            if (!startAngle.HasValue || !endAngle.HasValue || radius <= 0)
            {
                AddCircle(doc, center.X, center.Y, Math.Max(radius, 1));
                return;
            }

            var start = NormalizeDeg(startAngle.Value);
            var end = NormalizeDeg(endAngle.Value);
            var sweep = NormalizeSweep(start, end);
            if (sweep <= 0.0)
            {
                sweep += 360.0;
            }

            if (sweep >= 359.999)
            {
                AddCircle(doc, center.X, center.Y, radius);
                return;
            }

            doc.Entities.Add(new Arc(center, radius, start, start + sweep));
        }

        private static void AddEllipseAsPolyline(DxfDocument doc, CADEllipse ellipse)
        {
            var pts = SampleEllipse(ellipse, 64);
            AddPolyline(doc, pts, IsClosedLike(ellipse));
        }

        private static void AddPolyline(DxfDocument doc, IReadOnlyList<(double X, double Y)> points, bool closed)
        {
            if (points.Count < 2)
            {
                return;
            }

            var cleaned = points
                .Where((p, i) => i == 0 || DistanceSquared(points[i - 1], p) > 1e-12)
                .ToList();
            if (cleaned.Count < 2)
            {
                return;
            }

            if (closed && DistanceSquared(cleaned[0], cleaned[^1]) > 1e-12)
            {
                cleaned.Add(cleaned[0]);
            }

            var pl = new Polyline2D();
            pl.IsClosed = closed;
            foreach (var p in cleaned)
            {
                pl.Vertexes.Add(new Polyline2DVertex(new Vector2(p.X, p.Y)));
            }
            doc.Entities.Add(pl);
        }

        private static (double X, double Y) ToPoint(CAD2DPoint p)
        {
            return (p.X, p.Y);
        }

        private static (double X, double Y) ToPoint(DPoint p)
        {
            return (p.X, p.Y);
        }

        private static (double X, double Y) ToPoint(object p)
        {
            if (p is DPoint dp)
            {
                return (dp.X, dp.Y);
            }

            var xProp = p.GetType().GetProperty("X");
            var yProp = p.GetType().GetProperty("Y");
            if (xProp?.GetValue(p) is not null && yProp?.GetValue(p) is not null)
            {
                return (Convert.ToDouble(xProp.GetValue(p), CultureInfo.InvariantCulture), Convert.ToDouble(yProp.GetValue(p), CultureInfo.InvariantCulture));
            }

            return (0, 0);
        }

        private static IReadOnlyList<(double X, double Y)> ExtractPoints(CADLWPolyLine poly)
        {
            return ExtractVertices(poly);
        }

        private static IReadOnlyList<(double X, double Y)> ExtractPoints(CADPolyLine poly)
        {
            return ExtractVertices(poly);
        }

        private static IReadOnlyList<(double X, double Y)> ExtractPoints(CAD2DPolyline poly)
        {
            return ExtractVertices(poly);
        }

        private static IReadOnlyList<(double X, double Y)> ExpandCadLwPolyline(CADLWPolyLine poly)
        {
            var vertices = GetCadVertices(poly);
            if (vertices.Count < 2)
            {
                return [];
            }

            var closed = IsClosedLike(poly);
            var result = new List<(double X, double Y)>();
            var segCount = closed ? vertices.Count : vertices.Count - 1;

            for (var i = 0; i < segCount; i++)
            {
                var curr = vertices[i];
                var next = vertices[(i + 1) % vertices.Count];
                var p0 = ToPoint(curr.Point);
                var p1 = ToPoint(next.Point);
                var bulge = curr.Bulge;

                if (result.Count == 0)
                {
                    result.Add(p0);
                }
                else if (!AreSamePoint(result[^1], p0))
                {
                    result.Add(p0);
                }

                if (Math.Abs(bulge) < 1e-9)
                {
                    if (!AreSamePoint(result[^1], p1))
                    {
                        result.Add(p1);
                    }
                    continue;
                }

                var arcPts = SampleBulgeArc(p0, p1, bulge, 32);
                for (var k = 1; k < arcPts.Count; k++)
                {
                    if (!AreSamePoint(result[^1], arcPts[k]))
                    {
                        result.Add(arcPts[k]);
                    }
                }
            }

            if (closed && result.Count > 1 && !AreSamePoint(result[0], result[^1]))
            {
                result.Add(result[0]);
            }

            return result;
        }

        private static IReadOnlyList<(double X, double Y)> ExtractVertices(object poly)
        {
            var pts = new List<(double X, double Y)>();
            foreach (var v in GetCadVertices(poly))
            {
                pts.Add(ToPoint(v.Point));
            }
            return pts;
        }

        private static List<CADVertex> GetCadVertices(object poly)
        {
            var vertsProp = poly.GetType().GetProperty("Vertexes") ?? poly.GetType().GetProperty("Vertices") ?? poly.GetType().GetProperty("Points");
            if (vertsProp?.GetValue(poly) is not System.Collections.IEnumerable verts)
            {
                return [];
            }

            return verts.Cast<object>()
                .OfType<CADVertex>()
                .ToList();
        }

        private static bool AreSamePoint((double X, double Y) a, (double X, double Y) b)
        {
            return Math.Abs(a.X - b.X) <= 1e-9 && Math.Abs(a.Y - b.Y) <= 1e-9;
        }

        private static bool IsClosedLike(object poly)
        {
            var prop = poly.GetType().GetProperty("IsClosed") ?? poly.GetType().GetProperty("Closed");
            if (prop?.GetValue(poly) is bool b)
            {
                return b;
            }

            var points = poly.GetType().GetProperty("Vertexes")?.GetValue(poly) as System.Collections.IEnumerable;
            if (points is null)
            {
                return false;
            }

            var list = points.Cast<object>().ToList();
            if (list.Count < 3)
            {
                return false;
            }

            var first = GetPointFromVertex(list[0]);
            var last = GetPointFromVertex(list[^1]);
            return DistanceSquared(first, last) <= 1e-6;
        }

        private static (double X, double Y) GetPointFromVertex(object vertex)
        {
            if (vertex is CADVertex cv)
            {
                return ToPoint(cv.Point);
            }

            var xProp = vertex.GetType().GetProperty("X");
            var yProp = vertex.GetType().GetProperty("Y");
            if (xProp?.GetValue(vertex) is not null && yProp?.GetValue(vertex) is not null)
            {
                return (Convert.ToDouble(xProp.GetValue(vertex), CultureInfo.InvariantCulture), Convert.ToDouble(yProp.GetValue(vertex), CultureInfo.InvariantCulture));
            }

            var posProp = vertex.GetType().GetProperty("Position");
            var pos = posProp?.GetValue(vertex);
            xProp = pos?.GetType().GetProperty("X");
            yProp = pos?.GetType().GetProperty("Y");
            return (Convert.ToDouble(xProp?.GetValue(pos), CultureInfo.InvariantCulture), Convert.ToDouble(yProp?.GetValue(pos), CultureInfo.InvariantCulture));
        }

        private static IEnumerable<object> ExtractInsertEntities(CADInsert insert)
        {
            var items = new List<object>();

            var entitiesProp = insert.GetType().GetProperty("Entities") ?? insert.GetType().GetProperty("Items") ?? insert.GetType().GetProperty("Block")?.PropertyType.GetProperty("Entities");
            var value = entitiesProp?.GetValue(insert) ?? insert.GetType().GetProperty("Block")?.GetValue(insert)?.GetType().GetProperty("Entities")?.GetValue(insert.GetType().GetProperty("Block")?.GetValue(insert));
            if (value is System.Collections.IEnumerable enumerable && value is not string)
            {
                items.AddRange(enumerable.Cast<object>().Where(x => x is not null));
            }

            return items;
        }

        private static IReadOnlyList<(double X, double Y)> SampleEllipse(CADEllipse ellipse, int segments)
        {
            var pts = new List<(double X, double Y)>(segments + 1);
            var center = ToPoint(ellipse.Point);
            var major = ToPoint(ellipse.RadPt);
            var rx = Math.Sqrt(major.X * major.X + major.Y * major.Y);
            var ry = Math.Max(rx * Math.Max(Math.Abs(ellipse.Ratio), 0.1), 1e-6);
            if (rx <= 1e-9)
            {
                rx = Math.Max(Math.Abs(ellipse.Radius), 1e-6);
            }

            var rot = Math.Atan2(major.Y, major.X);
            var cosR = Math.Cos(rot);
            var sinR = Math.Sin(rot);

            for (var i = 0; i <= segments; i++)
            {
                var t = (double)i / segments * 2 * Math.PI;
                var x = rx * Math.Cos(t);
                var y = ry * Math.Sin(t);
                pts.Add((center.X + x * cosR - y * sinR, center.Y + x * sinR + y * cosR));
            }
            return pts;
        }

        private static IReadOnlyList<(double X, double Y)> SampleBulgeArc((double X, double Y) p0, (double X, double Y) p1, double bulge, int segments)
        {
            var dx = p1.X - p0.X;
            var dy = p1.Y - p0.Y;
            var chord = Math.Sqrt(dx * dx + dy * dy);
            if (chord <= 1e-12)
            {
                return [p0, p1];
            }

            var theta = 4.0 * Math.Atan(bulge);
            var radius = chord * (1.0 + bulge * bulge) / (4.0 * Math.Abs(bulge));
            var mx = (p0.X + p1.X) * 0.5;
            var my = (p0.Y + p1.Y) * 0.5;
            var nx = -dy / chord;
            var ny = dx / chord;
            var halfChord = chord * 0.5;
            var h = Math.Sqrt(Math.Max(0, radius * radius - halfChord * halfChord));
            var sign = bulge >= 0 ? 1.0 : -1.0;
            var cx = mx + sign * nx * h;
            var cy = my + sign * ny * h;

            var start = Math.Atan2(p0.Y - cy, p0.X - cx);
            var end = start + theta;
            var pts = new List<(double X, double Y)>(segments + 1);
            for (var i = 0; i <= segments; i++)
            {
                var t = (double)i / segments;
                var a = start + (end - start) * t;
                pts.Add((cx + radius * Math.Cos(a), cy + radius * Math.Sin(a)));
            }
            return pts;
        }

        private static double DistanceSquared((double X, double Y) a, (double X, double Y) b)
        {
            var dx = a.X - b.X;
            var dy = a.Y - b.Y;
            return dx * dx + dy * dy;
        }

        private static double? TryGetArcAngle(object arc, params string[] names)
        {
            foreach (var name in names)
            {
                var prop = arc.GetType().GetProperty(name);
                if (prop?.GetValue(arc) is null)
                {
                    continue;
                }

                var value = prop.GetValue(arc);
                if (value is double d)
                {
                    return NormalizeDeg(d * (Math.Abs(d) > 2.0 * Math.PI ? 1.0 : 180.0 / Math.PI));
                }

                if (value is float f)
                {
                    var dv = (double)f;
                    return NormalizeDeg(dv * (Math.Abs(dv) > 2.0 * Math.PI ? 1.0 : 180.0 / Math.PI));
                }

                try
                {
                    var dv = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                    return NormalizeDeg(dv * (Math.Abs(dv) > 2.0 * Math.PI ? 1.0 : 180.0 / Math.PI));
                }
                catch
                {
                }
            }

            return null;
        }

        private static double? TryGetAngleFromPoint(double cx, double cy, object arc, params string[] names)
        {
            foreach (var name in names)
            {
                var prop = arc.GetType().GetProperty(name);
                if (prop?.GetValue(arc) is null)
                {
                    continue;
                }

                var value = prop.GetValue(arc);
                var xProp = value?.GetType().GetProperty("X");
                var yProp = value?.GetType().GetProperty("Y");
                if (xProp?.GetValue(value) is null || yProp?.GetValue(value) is null)
                {
                    continue;
                }

                try
                {
                    var x = Convert.ToDouble(xProp.GetValue(value), CultureInfo.InvariantCulture);
                    var y = Convert.ToDouble(yProp.GetValue(value), CultureInfo.InvariantCulture);
                    return NormalizeDeg(Math.Atan2(y - cy, x - cx) * 180.0 / Math.PI);
                }
                catch
                {
                }
            }

            return null;
        }

        private static double NormalizeDeg(double degree)
        {
            var d = degree % 360.0;
            if (d < 0) d += 360.0;
            return d;
        }

        private static double NormalizeSweep(double startDeg, double endDeg)
        {
            var sweep = endDeg - startDeg;
            while (sweep < 0)
            {
                sweep += 360.0;
            }
            while (sweep > 360.0)
            {
                sweep -= 360.0;
            }
            return sweep;
        }

    }

    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private readonly IDxfPreviewPlugin _previewPlugin = new BasicCanvasPreviewPlugin();
        private const int CacheMaxEntries = 200;
        private readonly Dictionary<string, DxfDocument> _documentCache = [];
        private readonly LinkedList<string> _documentCacheLru = [];
        private readonly Dictionary<string, LinkedListNode<string>> _documentCacheLruNodes = [];
        private readonly Dictionary<string, ImageSource> _moldPreviewCache = [];
        private readonly LinkedList<string> _moldPreviewCacheLru = [];
        private readonly Dictionary<string, LinkedListNode<string>> _moldPreviewCacheLruNodes = [];
        private readonly InteractiveDxfPreview _viewer = new();
        private readonly ObservableCollection<MoldRow> _stage1MoldRows = [];
        private readonly ObservableCollection<MoldRow> _stage2MoldRows = [];
        private readonly ObservableCollection<PositionRow> _stage1PositionRows = [];
        private readonly ObservableCollection<PositionRow> _stage2PositionRows = [];
        private readonly ObservableCollection<PlcRegisterRow> _plcRegisters = [];
        private readonly Dictionary<string, PlcRegisterRow> _plcRegisterMap = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<string> _stage1MoldFiles = [];
        private readonly List<string> _stage2MoldFiles = [];
        private string? _selectedStage1File;
        private string? _selectedStage2File;
        private MatchResult? _lastMatchResult;
        private ProjectProfile? _lastProjectProfile;
        private List<MoldProfile> _lastMolds = [];
        private IReadOnlyList<(double X, double Y)> _lastOuterContourPoints = [];
        private readonly ModbusTcpCommService _modbusTcpCommService = ModbusTcpCommService.Shared;
        private readonly SemaphoreSlim _plcIoLock = new(1, 1);
        private ModbusTcpNet? _plcClient;
        private string _plcClientKey = string.Empty;
        private DateTime _lastPlcReconnectAttempt = DateTime.MinValue;
        private bool _plcIsConnected = false;
        private CancellationTokenSource? _automationCts;
        private CancellationTokenSource? _heartbeatCts;
        private Task? _heartbeatTask;
        private Task? _automationLoopTask;
        private bool _automationStartupRequested;
        private readonly string _projectFolderSettingsPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CADRecognition", "project-folder.txt");

        private PlcMonitorWindow? _plcMonitorWindow;
        private string? _projectFolder;
        private string? _projectFile;
        private DxfDocument? _projectDoc;
        private bool _compactAnnotation = false;
        private double _boardWidth = 0;
        private string _d600Value = string.Empty;
        private double _d620Value = 0;
        private int _d622Value = 0;
        private int _d623Value = 0;
        private int _d624Value = 0;
        private int _d625Value = 0;
        private int _d626Value = 0;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            MoldCountText.Text = "0";
            ProjectFileText.Text = "未加载";
            ProjectFolderText.Text = "未选择";
            PreviewHost.Content = _viewer;
            InitializePlcRegisters();
            _viewer.SetCompactMode(_compactAnnotation);
            FileTreeView.Items.Clear();
            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;

            // 初始化日志系统
            AppLogger.Instance.CleanOldLogs();
            AppLogger.Instance.LogOperation("软件启动", $"CADRecognition 应用程序已启动 [版本 {AppVersion.FullVersion}]");
            AppLogger.Instance.Info($"旧日志文件清理完成 [当前版本 {AppVersion.FullVersion}]");
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Loaded -= MainWindow_Loaded;
            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);
            await Task.Delay(100);
            RestoreProjectFolder();
            await TryRestoreLastMoldsAsync();
            RefreshPlcRegisters();
        }

        private void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            AppLogger.Instance.LogOperation("软件关闭", "CADRecognition 应用程序正在关闭");
            _automationCts?.Cancel();
            _heartbeatCts?.Cancel();
            SaveProjectFolder();
            LastMoldSessionSettings.Save(
                _stage1MoldFiles.Where(File.Exists).ToList(),
                _stage2MoldFiles.Where(File.Exists).ToList());
            AppLogger.Instance.Info("软件正常关闭");
            AppLogger.Instance.Dispose();
        }

        public void SetProjectFolder(string folder)
        {
            _projectFolder = folder;
            ProjectFolderText.Text = folder;
            SaveProjectFolder();
            //ProjectFolderChanged?.Invoke(this, EventArgs.Empty);
        }

        public void HandlePlcRegisterCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            PlcRegisterGrid_CellEditEnding(sender, e);
        }

        private async Task TryRestoreLastMoldsAsync()
        {
            var stored = LastMoldSessionSettings.Load();
            if (stored.Stage1.Count == 0 && stored.Stage2.Count == 0)
            {
                return;
            }

            StatusText.Text = "正在恢复上次模具...";
            var skipped = new List<string>();
            if (stored.Stage1.Count > 0)
            {
                skipped.AddRange(await ApplyMoldPathsForStageAsync(1, stored.Stage1, skipUnreadableFiles: true));
            }

            if (stored.Stage2.Count > 0)
            {
                skipped.AddRange(await ApplyMoldPathsForStageAsync(2, stored.Stage2, skipUnreadableFiles: true));
            }

            if (_stage1MoldFiles.Count == 0 && _stage2MoldFiles.Count == 0)
            {
                return;
            }

            LastMoldSessionSettings.Save(_stage1MoldFiles, _stage2MoldFiles);
            var msg = $"已自动加载上次模具：台1 {_stage1MoldFiles.Count} 张，台2 {_stage2MoldFiles.Count} 张。";
            if (skipped.Count > 0)
            {
                msg += " 无法读取已跳过：" + string.Join("、", skipped.Distinct(StringComparer.OrdinalIgnoreCase));
            }

            StatusText.Text = msg;
        }

        /// <summary>按路径加载某一工位模具；<paramref name="skipUnreadableFiles"/> 为 true 时跳过损坏或无法解析的文件。</summary>
        private IReadOnlyList<string> ApplyMoldPathsForStage(int stageId, IEnumerable<string> inputPaths, bool skipUnreadableFiles)
        {
            var distinctExisting = inputPaths
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Where(File.Exists)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var targetFiles = stageId == 1 ? _stage1MoldFiles : _stage2MoldFiles;
            targetFiles.Clear();
            var skipped = new List<string>();
            foreach (var file in distinctExisting)
            {
                try
                {
                    var moldDoc = LoadCadDocument(file);
                    RemoveDuplicateLines(moldDoc);
                    SetDocumentCache(file, moldDoc);
                    targetFiles.Add(file);
                }
                catch
                {
                    if (!skipUnreadableFiles)
                    {
                        throw;
                    }

                    skipped.Add(System.IO.Path.GetFileName(file));
                }
            }

            if (stageId == 1)
            {
                var items1 = new List<string> { "(无)" };
                items1.AddRange(_stage1MoldFiles.Select(System.IO.Path.GetFileName));
                Stage1MoldComboBox.ItemsSource = items1;
                _selectedStage1File = _stage1MoldFiles.FirstOrDefault();
                Stage1MoldComboBox.SelectedIndex = _selectedStage1File is null ? 0 : 1;
                RefreshMoldPreviewList(1);
            }
            else
            {
                var items2 = new List<string> { "(无)" };
                items2.AddRange(_stage2MoldFiles.Select(System.IO.Path.GetFileName));
                Stage2MoldComboBox.ItemsSource = items2;
                _selectedStage2File = _stage2MoldFiles.FirstOrDefault();
                Stage2MoldComboBox.SelectedIndex = _selectedStage2File is null ? 0 : 1;
                RefreshMoldPreviewList(2);
            }

            MoldCountText.Text = $"{_stage1MoldFiles.Count}/{_stage2MoldFiles.Count}";
            RefreshFileList();
            return skipped;
        }

        private async Task<IReadOnlyList<string>> ApplyMoldPathsForStageAsync(int stageId, IEnumerable<string> inputPaths, bool skipUnreadableFiles)
        {
            var distinctExisting = inputPaths
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Where(File.Exists)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var skipped = new List<string>();
            var loaded = new List<string>();
            var total = distinctExisting.Count;

            ImportProgressPanel.Visibility = Visibility.Visible;
            ImportProgressBar.Minimum = 0;
            ImportProgressBar.Maximum = Math.Max(total, 1);
            ImportProgressBar.Value = 0;
            ImportProgressText.Text = stageId == 1 ? "台1 模具导入中..." : "台2 模具导入中...";

            try
            {
                await Task.Run(() =>
                {
                    for (var i = 0; i < distinctExisting.Count; i++)
                    {
                        var file = distinctExisting[i];
                        try
                        {
                            var moldDoc = LoadCadDocument(file);
                            RemoveDuplicateLines(moldDoc);
                            Dispatcher.Invoke(() =>
                            {
                                SetDocumentCache(file, moldDoc);
                                ImportProgressBar.Value = i + 1;
                                ImportProgressText.Text = $"正在导入({i + 1}/{total})：{System.IO.Path.GetFileName(file)}";
                                StatusText.Text = $"正在导入({i + 1}/{total})：{System.IO.Path.GetFileName(file)}";
                            });
                            loaded.Add(file);
                        }
                        catch
                        {
                            if (!skipUnreadableFiles)
                            {
                                throw;
                            }

                            skipped.Add(System.IO.Path.GetFileName(file));
                        }
                    }
                });
            }
            finally
            {
                ImportProgressPanel.Visibility = Visibility.Collapsed;
            }

            var targetFiles = stageId == 1 ? _stage1MoldFiles : _stage2MoldFiles;
            targetFiles.Clear();
            targetFiles.AddRange(loaded);

            if (stageId == 1)
            {
                var items1 = new List<string> { "(无)" };
                items1.AddRange(_stage1MoldFiles.Select(System.IO.Path.GetFileName));
                Stage1MoldComboBox.ItemsSource = items1;
                _selectedStage1File = _stage1MoldFiles.FirstOrDefault();
                Stage1MoldComboBox.SelectedIndex = _selectedStage1File is null ? 0 : 1;
                RefreshMoldPreviewList(1);
            }
            else
            {
                var items2 = new List<string> { "(无)" };
                items2.AddRange(_stage2MoldFiles.Select(System.IO.Path.GetFileName));
                Stage2MoldComboBox.ItemsSource = items2;
                _selectedStage2File = _stage2MoldFiles.FirstOrDefault();
                Stage2MoldComboBox.SelectedIndex = _selectedStage2File is null ? 0 : 1;
                RefreshMoldPreviewList(2);
            }

            MoldCountText.Text = $"{_stage1MoldFiles.Count}/{_stage2MoldFiles.Count}";
            RefreshFileList();
            return skipped;
        }

        public ObservableCollection<MoldRow> Stage1MoldRows => _stage1MoldRows;
        public ObservableCollection<MoldRow> Stage2MoldRows => _stage2MoldRows;
        public ObservableCollection<PositionRow> Stage1PositionRows => _stage1PositionRows;
        public ObservableCollection<PositionRow> Stage2PositionRows => _stage2PositionRows;
        public ObservableCollection<PlcRegisterRow> PlcRegisters => _plcRegisters;
        public string? CurrentProjectFolder => _projectFolder;

        private void InitializePlcRegisters()
        {
            _plcRegisters.Clear();
            _plcRegisterMap.Clear();

            AddPlcRegisterRow("D600", string.Empty, "文件名");
            AddPlcRegisterRow("D620", "0", "识图界限");
            AddPlcRegisterRow("D622", "0", "软件状态");
            AddPlcRegisterRow("D623", "0", "识图结果（1=成功，2=失败，3=未找到）");
            AddPlcRegisterRow("D624", "0", "识图中（1=进行中，0=结束）");
            AddPlcRegisterRow("D625", "0", "心跳");
            AddPlcRegisterRow("D626", "0", "任务号");
        }

        private string GetPlcAddress(string logicalName) => _plcRegisterMap.TryGetValue(logicalName, out var row) ? row.Address : logicalName;
        private string GetPlcAddressAt(int index) => index >= 0 && index < _plcRegisters.Count ? _plcRegisters[index].Address : string.Empty;

        private void AddPlcRegisterRow(string logicalName, string value, string info)
        {
            var row = new PlcRegisterRow(logicalName, value, info);
            _plcRegisters.Add(row);
            _plcRegisterMap[logicalName] = row;
        }

        private void RefreshPlcRegisters()
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(RefreshPlcRegistersCore);
                return;
            }

            RefreshPlcRegistersCore();
        }

        private void RefreshPlcRegistersCore()
        {
            if (_plcRegisters.Count < 7) InitializePlcRegisters();
            _plcRegisters[0].Value = _d600Value;
            _plcRegisters[1].Value = _d620Value.ToString("0.###", CultureInfo.InvariantCulture);
            _plcRegisters[2].Value = _d622Value.ToString(CultureInfo.InvariantCulture);
            _plcRegisters[3].Value = _d623Value.ToString(CultureInfo.InvariantCulture);
            _plcRegisters[4].Value = _d624Value.ToString(CultureInfo.InvariantCulture);
            _plcRegisters[5].Value = _d625Value.ToString(CultureInfo.InvariantCulture);
            _plcRegisters[6].Value = _d626Value.ToString(CultureInfo.InvariantCulture);
            OnPropertyChanged(nameof(PlcRegisters));
        }

        public void RefreshPlcRegistersForMonitor() => RefreshPlcRegisters();

        public async Task LoadPlcRegistersForMonitorAsync(CancellationToken token = default)
        {
            var d600Address = GetPlcAddressAt(0);
            var d620Address = GetPlcAddressAt(1);
            var d626Address = GetPlcAddressAt(6);

            if (string.IsNullOrWhiteSpace(d600Address) || string.IsNullOrWhiteSpace(d620Address) || string.IsNullOrWhiteSpace(d626Address))
            {
                RefreshPlcRegisters();
                return;
            }

            try
            {
                var previousD626 = _d626Value;
                var plcHost = TcpExportDialog.SharedTcpHost;
                var plcPort = int.TryParse(TcpExportDialog.SharedTcpPort, out var parsedPort) ? parsedPort : 502;
                byte station = byte.TryParse(TcpExportDialog.SharedModbusStation, out var parsedStation) ? parsedStation : (byte)1;
                await ReadD600FileNameAsync(plcHost, plcPort, station, d600Address, token).ConfigureAwait(true);
                await ReadPlcBoundaryAsync(plcHost, plcPort, station, d620Address, token).ConfigureAwait(true);
                await ReadIntAsync(plcHost, plcPort, station, d626Address, 626, token).ConfigureAwait(true);
                RefreshPlcRegisters();
            }
            catch (Exception ex) when (ex is OperationCanceledException or InvalidOperationException or FormatException or TimeoutException)
            {
                StatusText.Text = $"PLC监视读取失败：{ex.Message}";
                RefreshPlcRegisters();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        private void SetStatus(string message)
        {
            if (Dispatcher.CheckAccess())
            {
                StatusText.Text = message;
            }
            else
            {
                Dispatcher.Invoke(() => StatusText.Text = message);
            }
            AppLogger.Instance.LogStatus(message);
        }

        private void RunOnUiThread(Action action)
        {
            if (Dispatcher.CheckAccess())
            {
                action();
            }
            else
            {
                Dispatcher.Invoke(action);
            }
        }

        /// <summary>
        /// 在后台线程执行任务，完成后自动回到 UI 线程。
        /// </summary>
        private async Task<T> RunInBackgroundAsync<T>(Func<Task<T>> backgroundWork)
        {
            var result = await Task.Run(async () => await backgroundWork().ConfigureAwait(false)).ConfigureAwait(false);
            return result;
        }

        /// <summary>
        /// 在后台线程执行任务，完成后自动回到 UI 线程（无返回值）。
        /// </summary>
        private async Task RunInBackgroundAsync(Func<Task> backgroundWork)
        {
            await Task.Run(async () => await backgroundWork().ConfigureAwait(false)).ConfigureAwait(false);
        }

        /// <summary>
        /// 在 UI 线程执行操作（用于更新 UI）。
        /// </summary>
        private async Task RunOnUiThreadAsync(Func<Task> action)
        {
            if (Dispatcher.CheckAccess())
            {
                await action().ConfigureAwait(true);
                return;
            }

            var operation = Dispatcher.InvokeAsync(action);
            await operation.Task.ConfigureAwait(false);
        }

        /// <summary>
        /// 在 UI 线程执行操作并获取返回值。
        /// </summary>
        private async Task<T> RunOnUiThreadAsync<T>(Func<Task<T>> action)
        {
            if (Dispatcher.CheckAccess())
            {
                return await action().ConfigureAwait(true);
            }

            var operation = Dispatcher.InvokeAsync(action);
            var task = await operation.Task.ConfigureAwait(false);
            return await task.ConfigureAwait(false);
        }

        private void CompactAnnoCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            _compactAnnotation = CompactAnnoCheckBox.IsChecked != false;
            _viewer.SetCompactMode(_compactAnnotation);

            if (_projectDoc is not null)
            {
                RenderPreview(_projectDoc, _projectFile, withAnnotation: _lastMatchResult is not null);
            }
        }

        private void SetDefaultBoardWidthFromProject()
        {
            if (_projectDoc is null || BoardWidthTextBox is null)
            {
                return;
            }

            var bounds = DxfAnalyzer.ExtractProject(_projectDoc).OuterRectangle;
            var defaultWidth = Math.Max(0, Math.Round(bounds.Height / 2.0, MidpointRounding.AwayFromZero));
            _boardWidth = defaultWidth;
            BoardWidthTextBox.Text = defaultWidth.ToString("0", CultureInfo.InvariantCulture);
        }

        private double ReadBoardWidth()
        {
            if (BoardWidthTextBox is null)
            {
                return _boardWidth;
            }

            var text = BoardWidthTextBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return _boardWidth;
            }

            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) ||
                double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value))
            {
                _boardWidth = Math.Max(0, value);
                return _boardWidth;
            }

            StatusText.Text = "板宽输入无效，请输入数字。";
            return _boardWidth;
        }

        private void BoardWidthTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var width = ReadBoardWidth();
            if (_projectDoc is not null && !string.IsNullOrWhiteSpace(_projectFile))
            {
                RenderPreview(_projectDoc, _projectFile, withAnnotation: _lastMatchResult is not null);
            }
            else if (_lastProjectProfile is not null)
            {
                _viewer.RenderCornerContours(
                    _lastProjectProfile.OuterRectangle,
                    _lastOuterContourPoints,
                    _lastMatchResult?.GuidePaths,
                    _lastProjectProfile.CornerCandidates,
                    width,
                    _lastProjectProfile.OuterRectangle.MinY + width);
            }
        }

        private void ImportProjectDxf_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new WinOpenFileDialog
            {
                Filter = "CAD 文件 (*.dxf;*.dwg)|*.dxf;*.dwg|DXF 文件 (*.dxf)|*.dxf|DWG 文件 (*.dwg)|*.dwg",
                Multiselect = false
            };
            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            _projectFile = dialog.FileName;
            _projectDoc = LoadCadDocument(_projectFile);
            var removedProjectLines = RemoveDuplicateLines(_projectDoc);
            SetDocumentCache(_projectFile, _projectDoc);

            // 导入新工程时清空上一张图纸的识别/标注展示状态；已导入的模具右侧预览保留并按文件重建（避免空白）。
            _lastMatchResult = null;
            _lastProjectProfile = DxfAnalyzer.ExtractProject(_projectDoc);
            _lastOuterContourPoints = DxfAnalyzer.ExtractOuterContourForDebug(_projectDoc);
            if (_stage1MoldFiles.Count == 0)
            {
                _stage1MoldRows.Clear();
            }
            else
            {
                RefreshMoldPreviewList(1);
            }

            if (_stage2MoldFiles.Count == 0)
            {
                _stage2MoldRows.Clear();
            }
            else
            {
                RefreshMoldPreviewList(2);
            }

            _stage1PositionRows.Clear();
            _stage2PositionRows.Clear();
            Stage1LegendPanel.Children.Clear();
            Stage2LegendPanel.Children.Clear();

            ProjectFileText.Text = System.IO.Path.GetFileName(_projectFile);
            RefreshFileList();
            SetDefaultBoardWidthFromProject();
            RenderPreview(_projectDoc, _projectFile, withAnnotation: false);
            StatusText.Text = removedProjectLines > 0
                ? $"工程 DXF 已加载，已去重重叠线段 {removedProjectLines} 条。"
                : "工程 DXF 已加载。";
        }

        private void ImportStage1MoldDxf_Click(object sender, RoutedEventArgs e)
        {
            ImportMoldsForStage(1);
        }

        private void ImportStage2MoldDxf_Click(object sender, RoutedEventArgs e)
        {
            ImportMoldsForStage(2);
        }

        private async void ImportMoldsForStage(int stageId)
        {
            var dialog = new WinOpenFileDialog
            {
                Filter = "CAD 文件 (*.dxf;*.dwg)|*.dxf;*.dwg|DXF 文件 (*.dxf)|*.dxf|DWG 文件 (*.dwg)|*.dwg",
                Multiselect = true
            };
            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            var files = dialog.FileNames;
            StatusText.Text = stageId == 1 ? "台1 模具导入中..." : "台2 模具导入中...";

            try
            {
                await ApplyMoldPathsForStageAsync(stageId, files, skipUnreadableFiles: false);
                LastMoldSessionSettings.Save(_stage1MoldFiles, _stage2MoldFiles);
                StatusText.Text = stageId == 1
                    ? $"已导入台1模具 {_stage1MoldFiles.Count} 张。"
                    : $"已导入台2模具 {_stage2MoldFiles.Count} 张。";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"导入失败：{ex.Message}";
            }
        }

        private DxfDocument LoadCadDocument(string path)
        {
            try
            {
                var ext = System.IO.Path.GetExtension(path);
                if (string.Equals(ext, ".dwg", StringComparison.OrdinalIgnoreCase))
                {
                    // DWG 文件加载需要 CADEditorControl，该控件必须在 UI 线程上创建和使用
                    // 从自动化后台线程调用时必须通过 Dispatcher 切换到 UI 线程
                    if (Dispatcher.CheckAccess())
                    {
                        return CadDocumentLoader.Load(path);
                    }

                    DxfDocument? doc = null;
                    Exception? loadEx = null;
                    Dispatcher.Invoke(() =>
                    {
                        try
                        {
                            doc = CadDocumentLoader.Load(path);
                        }
                        catch (Exception ex)
                        {
                            loadEx = ex;
                        }
                    });
                    if (loadEx is not null)
                    {
                        throw loadEx;
                    }
                    if (doc is null)
                    {
                        throw new InvalidOperationException("DWG 文件加载返回了空文档。");
                    }
                    return doc;
                }

                return CadDocumentLoader.Load(path);
            }
            catch (Exception ex)
            {
                StatusText.Text = $"读取失败：{System.IO.Path.GetFileName(path)}，{ex.Message}";
                throw;
            }
        }

        private static int RemoveDuplicateLines(DxfDocument doc)
        {
            if (doc?.Entities == null || DxfSafe.Lines(doc) == null)
            {
                return 0;
            }

            var lines = doc.Entities.Lines.Where(l => l != null).ToList();
            if (lines.Count <= 1)
            {
                return 0;
            }

            const double tol = 1e-4;
            (double X, double Y) Snap((double X, double Y) p)
                => (Math.Round(p.X / tol) * tol, Math.Round(p.Y / tol) * tol);

            string Key((double X, double Y) a, (double X, double Y) b)
            {
                var sa = Snap(a);
                var sb = Snap(b);
                var k1 = $"{sa.X:F4},{sa.Y:F4}";
                var k2 = $"{sb.X:F4},{sb.Y:F4}";
                return string.CompareOrdinal(k1, k2) <= 0 ? $"{k1}|{k2}" : $"{k2}|{k1}";
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var dup = new List<Line>();

            foreach (var l in lines)
            {
                var key = Key((l.StartPoint.X, l.StartPoint.Y), (l.EndPoint.X, l.EndPoint.Y));
                if (!seen.Add(key))
                {
                    dup.Add(l);
                }
            }

            foreach (var d in dup)
            {
                doc.Entities.Remove(d);
            }

            return dup.Count;
        }

        private string? ResolveSelectedMoldFile(ComboBox comboBox, IReadOnlyList<string> files)
        {
            if (files.Count == 0)
            {
                return null;
            }

            if (comboBox.SelectedItem is string selectedName)
            {
                if (selectedName == "(无)")
                {
                    return null;
                }

                var matched = files.FirstOrDefault(f =>
                    string.Equals(System.IO.Path.GetFileName(f), selectedName, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrWhiteSpace(matched))
                {
                    return matched;
                }
            }

            return files.First();
        }

        private async void Recognize_Click(object sender, RoutedEventArgs e)
        {
            await RecognizeAndSendAsync(sendToPlc: false, writeHeartbeatOnly: false);
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            if (_lastMatchResult is null || _projectFile is null)
            {
                StatusText.Text = "请先完成识图后再导出。";
                return;
            }

            var dialog = new TcpExportDialog(BuildTcpExportModel(ReadBoardWidth()))
            {
                Owner = this
            };

            if (dialog.ShowDialog() == true)
            {
                StatusText.Text = "已导出并通过 Modbus TCP 发送。";
            }
        }

        private void OpenLogFolder_Click(object sender, RoutedEventArgs e)
        {
            var logPath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "CADRecognition",
                "Logs");

            try
            {
                Directory.CreateDirectory(logPath);
                System.Diagnostics.Process.Start("explorer.exe", logPath);
                AppLogger.Instance.Info($"打开日志文件夹: {logPath}");
            }
            catch (Exception ex)
            {
                AppLogger.Instance.Error($"打开日志文件夹失败: {ex.Message}", ex);
                StatusText.Text = $"打开日志文件夹失败：{ex.Message}";
            }
        }

        private async void AutoImportRecognizeSend_Click(object sender, RoutedEventArgs e)
        {
            if (_plcMonitorWindow is null || !_plcMonitorWindow.IsLoaded)
            {
                _plcMonitorWindow = new PlcMonitorWindow(this)
                {
                    Owner = this
                };
                _plcMonitorWindow.Closed += async (_, _) =>
                {
                    await StopAutomationAsync().ConfigureAwait(true);
                    _plcMonitorWindow = null;
                };
                _plcMonitorWindow.Show();
            }
            else
            {
                _plcMonitorWindow.Activate();
            }

            await EnsureAutomationLoopRunningAsync().ConfigureAwait(true);
        }

        private async Task StopAutomationAsync()
        {
            AppLogger.Instance.Info("停止自动化任务...");
            _automationCts?.Cancel();
            _heartbeatCts?.Cancel();

            if (_automationLoopTask is not null)
            {
                try { await _automationLoopTask.ConfigureAwait(true); } catch { }
            }

            if (_heartbeatTask is not null)
            {
                try { await _heartbeatTask.ConfigureAwait(true); } catch { }
            }

            _automationCts?.Dispose();
            _automationCts = null;
            _heartbeatCts?.Dispose();
            _heartbeatCts = null;
            _automationLoopTask = null;
            _heartbeatTask = null;

            AppLogger.Instance.Info("自动化任务已停止");
        }

        private async Task EnsureAutomationLoopRunningAsync()
        {
            if (_automationLoopTask is not null && !_automationLoopTask.IsCompleted)
            {
                StatusText.Text = "自动识别后台已在运行。";
                return;
            }

            if (_automationCts is not null)
            {
                _automationCts.Cancel();
            }

            _automationCts = new CancellationTokenSource();
            var token = _automationCts.Token;
            _automationLoopTask = Task.Run(async () =>
            {
                try
                {
                    if (!await EnsureProjectAndMoldsReadyForAutomationAsync(token).ConfigureAwait(false))
                    {
                        return;
                    }

                    await RunFullAutomationAsync(token).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    SetStatus("自动流程已取消。");
                }
                catch (Exception ex)
                {
                    SetStatus($"自动流程失败：{ex.Message}");
                }
                finally
                {
                    _automationCts?.Dispose();
                    _automationCts = null;
                    _automationLoopTask = null;
                }
            }, token);

            await Task.CompletedTask;
        }

        private async Task TriggerAutomationOnD626ChangeAsync(int previousD626, int currentD626, CancellationToken token)
        {
            if (previousD626 == currentD626)
            {
                return;
            }

            if (_automationLoopTask is not null && !_automationLoopTask.IsCompleted)
            {
                return;
            }

            try
            {
                AppLogger.Instance.LogOperation("D626 变化触发", $"检测到任务号变化: {previousD626} -> {currentD626}");
                await RunD626TriggeredAutomationAsync(token).ConfigureAwait(true);
            }
            catch (OperationCanceledException)
            {
                AppLogger.Instance.Warn("D626 变化后启动自动化已取消");
                SetStatus("自动化任务已取消");
            }
            catch (TimeoutException ex)
            {
                AppLogger.Instance.Error($"Modbus 操作超时: {ex.Message}", ex);
                SetStatus($"Modbus 操作超时: {ex.Message}");
            }
            catch (InvalidOperationException ex)
            {
                AppLogger.Instance.Error($"D626 变化后启动自动化失败: {ex.Message}", ex);
                SetStatus($"D626 变化后启动自动化失败：{ex.Message}");
            }
            catch (Exception ex)
            {
                AppLogger.Instance.Error($"D626 变化后启动自动化发生未知错误: {ex.Message}", ex);
                SetStatus($"发生未知错误：{ex.Message}");
            }
        }

        private async Task RunD626TriggeredAutomationAsync(CancellationToken token)
        {
            AppLogger.Instance.Info("开始执行 D626 触发的自动化任务");

            if (!await EnsureProjectAndMoldsReadyForAutomationAsync(token).ConfigureAwait(true))
            {
                AppLogger.Instance.Warn("自动化任务前置条件不满足，已取消");
                return;
            }

            var plcHost = TcpExportDialog.SharedTcpHost;
            var plcPort = int.TryParse(TcpExportDialog.SharedTcpPort, out var parsedPort) ? parsedPort : 502;
            byte station = byte.TryParse(TcpExportDialog.SharedModbusStation, out var parsedStation) ? parsedStation : (byte)1;
            var d600Address = GetPlcAddressAt(0);
            var d620Address = GetPlcAddressAt(1);
            var d622Address = GetPlcAddressAt(2);
            var d623Address = GetPlcAddressAt(3);
            var d624Address = GetPlcAddressAt(4);
            var d625Address = GetPlcAddressAt(5);

            AppLogger.Instance.Info($"PLC配置: Host={plcHost}, Port={plcPort}, Station={station}");

            if (!string.IsNullOrWhiteSpace(d622Address))
            {
                _d622Value = 1;
                RefreshPlcRegisters();
                await WriteSingleIntAsync(plcHost, plcPort, station, d622Address, 1, token).ConfigureAwait(true);
                AppLogger.Instance.Info($"写入 D622={1} (地址: {d622Address})");
            }

            if (!string.IsNullOrWhiteSpace(d624Address))
            {
                _d624Value = 1;
                RefreshPlcRegisters();
                await WriteSingleIntAsync(plcHost, plcPort, station, d624Address, 1, token).ConfigureAwait(true);
                AppLogger.Instance.Info($"写入 D624=1 标记识图开始 (地址: {d624Address})");
            }

            if (!string.IsNullOrWhiteSpace(d625Address))
            {
                AppLogger.Instance.Info($"启动心跳: D625 地址={d625Address}");
                await StartHeartbeatAsync(plcHost, plcPort, station, d625Address, token).ConfigureAwait(true);
            }

            try
            {
                AppLogger.Instance.Info($"读取文件名: D600 地址={d600Address}");
                var fileName = await ReadD600FileNameAsync(plcHost, plcPort, station, d600Address, token).ConfigureAwait(true);
                int zeroIndex = fileName.IndexOf('\0');
                fileName = zeroIndex > 0 ? fileName.Substring(0, zeroIndex) : fileName;
                AppLogger.Instance.Info($"从 PLC 读取文件名: {fileName}");

                var matchedFile = FindFileInFolderByName(_projectFolder!, fileName);
                if (string.IsNullOrWhiteSpace(fileName))
                {
                    AppLogger.Instance.Debug($"D600 文件名称为空，等待 PLC 写入文件名，本次跳过。");
                    return;
                }

                if (matchedFile is null)
                {
                    _d623Value = 2;
                    RefreshPlcRegisters();
                    await WriteSingleIntAsync(plcHost, plcPort, station, d623Address, 2, token).ConfigureAwait(true);
                    _d624Value = 0;
                    RefreshPlcRegisters();
                    await WriteSingleIntAsync(plcHost, plcPort, station, d624Address, 0, token).ConfigureAwait(true);
                    StatusText.Text = $"在图纸文件夹中未找到：{fileName}";
                    AppLogger.Instance.Error($"图纸文件未找到: {fileName}");
                    return;
                }

                AppLogger.Instance.Info($"加载图纸文件: {System.IO.Path.GetFileName(matchedFile)}");
                _projectFile = matchedFile;
                _projectDoc = LoadCadDocument(matchedFile);
                SetDocumentCache(matchedFile, _projectDoc);
                _lastProjectProfile = DxfAnalyzer.ExtractProject(_projectDoc);
                _lastOuterContourPoints = DxfAnalyzer.ExtractOuterContourForDebug(_projectDoc);
                ProjectFileText.Text = System.IO.Path.GetFileName(matchedFile);
                RenderPreview(_projectDoc, _projectFile, withAnnotation: false);

                AppLogger.Instance.Info($"读取识图界限: D620 地址={d620Address}");
                var splitBoundary = await ReadPlcBoundaryAsync(plcHost, plcPort, station, d620Address, token).ConfigureAwait(true);
                _boardWidth = splitBoundary;
                AppLogger.Instance.Info($"识图界限读取成功: {_boardWidth}");
                if (BoardWidthTextBox is not null)
                {
                    BoardWidthTextBox.Text = splitBoundary.ToString("0.###", CultureInfo.InvariantCulture);
                }

                AppLogger.Instance.Info("开始执行识别并发送结果到 PLC...");
                // 识图开始，置 D623=0 表示正在进行中
                _d623Value = 0;
                RefreshPlcRegisters();
                await WriteSingleIntAsync(plcHost, plcPort, station, d623Address, 0, token).ConfigureAwait(true);
                var ok = await RecognizeAndSendAsync(sendToPlc: true, writeHeartbeatOnly: true).ConfigureAwait(true);
                _d624Value = 0;
                _d623Value = ok ? 1 : 2;
                RefreshPlcRegisters();
                await WriteSingleIntAsync(plcHost, plcPort, station, d624Address, 0, token).ConfigureAwait(true);
                await WriteSingleIntAsync(plcHost, plcPort, station, d623Address, ok ? 1 : 2, token).ConfigureAwait(true);
                StatusText.Text = ok ? "自动识别并发送成功，等待下一次 D626 变化..." : "自动识别失败，等待下一次 D626 变化...";
                AppLogger.Instance.Info($"识别完成: {(ok ? "成功" : "失败")}, D623={_d623Value}, D624=0");
            }
            catch (TimeoutException ex)
            {
                AppLogger.Instance.Error($"Modbus 操作超时导致自动化失败: {ex.Message}", ex);
                SetStatus($"操作超时：{ex.Message}");
                _d624Value = 0;
                _d623Value = 2;
                RefreshPlcRegisters();
                try { await WriteSingleIntAsync(plcHost, plcPort, station, d624Address, 0, token).ConfigureAwait(true); } catch { }
                try { await WriteSingleIntAsync(plcHost, plcPort, station, d623Address, 2, token).ConfigureAwait(true); } catch { }
            }
        }

        private async Task<bool> EnsureProjectAndMoldsReadyForAutomationAsync(CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(_projectFolder) || !Directory.Exists(_projectFolder))
            {
                using var dialog = new FormsFolderBrowserDialog
                {
                    Description = "请选择图纸文件所在文件夹",
                    ShowNewFolderButton = false,
                    SelectedPath = _projectFolder ?? string.Empty
                };

                if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK || string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    return false;
                }

                SetProjectFolder(dialog.SelectedPath);
            }

            if (_stage1MoldFiles.Count == 0 || _stage2MoldFiles.Count == 0)
            {
                SetStatus("请先按原逻辑导入台1和台2模具。");
                return false;
            }

            token.ThrowIfCancellationRequested();
            return true;
        }

        private void RestoreProjectFolder()
        {
            try
            {
                if (File.Exists(_projectFolderSettingsPath))
                {
                    var folder = File.ReadAllText(_projectFolderSettingsPath).Trim();
                    if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
                    {
                        _projectFolder = folder;
                        ProjectFolderText.Text = folder;
                    }
                }
            }
            catch
            {
            }
        }

        private void SaveProjectFolder()
        {
            try
            {
                var folder = _projectFolder?.Trim();
                if (string.IsNullOrWhiteSpace(folder))
                {
                    return;
                }

                var dir = System.IO.Path.GetDirectoryName(_projectFolderSettingsPath);
                if (!string.IsNullOrWhiteSpace(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                File.WriteAllText(_projectFolderSettingsPath, folder);
            }
            catch
            {
            }
        }

        private static string? FindFileInFolderByName(string folder, string fileName)
        {
            if (string.IsNullOrWhiteSpace(folder) || string.IsNullOrWhiteSpace(fileName) || !Directory.Exists(folder))
            {
                return null;
            }

            return Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories)
                .FirstOrDefault(f => string.Equals(System.IO.Path.GetFileName(f), fileName, StringComparison.OrdinalIgnoreCase));
        }

        private async Task RunFullAutomationAsync(CancellationToken token)
        {
            AppLogger.Instance.LogOperation("启动完整自动化", "RunFullAutomationAsync 开始执行");

            if (string.IsNullOrWhiteSpace(_projectFolder) || !Directory.Exists(_projectFolder))
            {
                var ex = new InvalidOperationException("未选择图纸文件夹。");
                AppLogger.Instance.Error(ex.Message, ex);
                throw ex;
            }

            var plcHost = TcpExportDialog.SharedTcpHost;
            var plcPort = int.TryParse(TcpExportDialog.SharedTcpPort, out var parsedPort) ? parsedPort : 502;
            byte station = byte.TryParse(TcpExportDialog.SharedModbusStation, out var parsedStation) ? parsedStation : (byte)1;
            var d600Address = GetPlcAddressAt(0);
            var d620Address = GetPlcAddressAt(1);
            var d622Address = GetPlcAddressAt(2);
            var d623Address = GetPlcAddressAt(3);
            var d624Address = GetPlcAddressAt(4);
            var d625Address = GetPlcAddressAt(5);
            var d626Address = GetPlcAddressAt(6);

            AppLogger.Instance.Info($"PLC配置: Host={plcHost}, Port={plcPort}, Station={station}");

            await WriteSingleIntAsync(plcHost, plcPort, station, d622Address, 1, token).ConfigureAwait(true);
            _d622Value = 1;
            RefreshPlcRegisters();
            AppLogger.Instance.Info("写入 D622=1，启动自动化");

            await StartHeartbeatAsync(plcHost, plcPort, station, d625Address, token).ConfigureAwait(true);
            AppLogger.Instance.Info("启动心跳线程");

            try
            {
                while (!token.IsCancellationRequested)
                {
                    AppLogger.Instance.Debug("读取 PLC 寄存器快照...");
                    await UpdatePlcSnapshotAsync(plcHost, plcPort, station, d600Address, d620Address, d622Address, d623Address, d624Address, d625Address, d626Address, token).ConfigureAwait(true);

                    var initialD626 = _d626Value;
                    SetStatus($"已记录 {d626Address} 初始值：{initialD626}，等待变化...");
                    AppLogger.Instance.Info($"等待 D626 变化，当前值={initialD626}");

                    try
                    {
                        var currentD626 = await WaitForRegisterChangeAsync(plcHost, plcPort, station, d626Address, initialD626, 626, token).ConfigureAwait(true);
                        SetStatus($"收到新任务号 {d626Address}={currentD626}，准备开始自动识别流程...");
                        AppLogger.Instance.Info($"D626 变化检测: {initialD626} -> {currentD626}");
                    }
                    catch (TimeoutException ex)
                    {
                        AppLogger.Instance.Error($"等待 D626 变化超时: {ex.Message}", ex);
                        continue;
                    }

                    await WriteSingleIntAsync(plcHost, plcPort, station, d622Address, 1, token).ConfigureAwait(true);
                    _d622Value = 1;
                    RefreshPlcRegisters();
                    AppLogger.Instance.Info($"写入 D622=1");

                    SetStatus($"检测到 {d626Address} 变化，正在读取 {d600Address} 文件名...");
                    AppLogger.Instance.Info($"读取文件名: D600 地址={d600Address}");

                    try
                    {
                        var fileName = await ReadD600FileNameAsync(plcHost, plcPort, station, d600Address, token).ConfigureAwait(true);
                        int zeroIndex = fileName.IndexOf('\0');
                        fileName = zeroIndex > 0 ? fileName.Substring(0, zeroIndex) : fileName;
                        if (string.IsNullOrWhiteSpace(fileName))
                        {
                            AppLogger.Instance.Debug($"PLC 尚未写入文件名，本次跳过，等待下一次任务。");
                            return;
                        }

                        AppLogger.Instance.Info($"从 PLC 读取文件名: {fileName}");
                        SetStatus($"正在按 {d600Address} 文件名查找图纸文件：{fileName}");

                        var matchedFile = FindFileInFolderByName(_projectFolder, fileName);
                        if (matchedFile is null)
                        {
                            _d623Value = 3;
                            RefreshPlcRegisters();
                            await WriteSingleIntAsync(plcHost, plcPort, station, d623Address, 3, token).ConfigureAwait(true);
                            SetStatus($"在图纸文件夹中未找到：{fileName}");
                            AppLogger.Instance.Error($"图纸文件未找到: {fileName}");
                            continue;
                        }

                        AppLogger.Instance.Info($"读取识图界限: D620 地址={d620Address}");
                        SetStatus($"正在读取 {d620Address} 识图界限...");

                        var splitBoundary = await ReadPlcBoundaryAsync(plcHost, plcPort, station, d620Address, token).ConfigureAwait(true);
                        _boardWidth = splitBoundary;
                        AppLogger.Instance.Info($"识图界限读取成功: {_boardWidth}");
                        if (BoardWidthTextBox is not null)
                        {
                            RunOnUiThread(() => BoardWidthTextBox.Text = splitBoundary.ToString("0.###", CultureInfo.InvariantCulture));
                        }

                        AppLogger.Instance.Info($"加载图纸文件: {System.IO.Path.GetFileName(matchedFile)}");
                        SetStatus($"正在读取图纸文件：{System.IO.Path.GetFileName(matchedFile)}");
                        _projectFile = matchedFile;
                        _projectDoc = LoadCadDocument(matchedFile);
                        SetDocumentCache(matchedFile, _projectDoc);
                        _lastProjectProfile = DxfAnalyzer.ExtractProject(_projectDoc);
                        _lastOuterContourPoints = DxfAnalyzer.ExtractOuterContourForDebug(_projectDoc);
                        RunOnUiThread(() =>
                        {
                            ProjectFileText.Text = System.IO.Path.GetFileName(matchedFile);
                            RenderPreview(_projectDoc, _projectFile, withAnnotation: false);
                        });

                        _d624Value = 1;
                        RefreshPlcRegisters();
                        SetStatus($"已读取识图界限 {d620Address}={splitBoundary:0.###}，正在写入 {d624Address}=1 并开始识别...");
                        AppLogger.Instance.Info($"写入 D624=1，识图开始");
                        await WriteSingleIntAsync(plcHost, plcPort, station, d624Address, 1, token).ConfigureAwait(true);

                        AppLogger.Instance.Info("开始执行识别并发送结果...");
                        SetStatus("正在自动识别并发送结果到 PLC...");
                        // 识图开始，置 D623=0 表示正在进行中
                        _d623Value = 0;
                        RefreshPlcRegisters();
                        await WriteSingleIntAsync(plcHost, plcPort, station, d623Address, 0, token).ConfigureAwait(true);
                        var ok = await RunOnUiThreadAsync(() => RecognizeAndSendAsync(sendToPlc: true, writeHeartbeatOnly: true)).ConfigureAwait(false);

                        _d624Value = 0;
                        _d623Value = ok ? 1 : 2;
                        RefreshPlcRegisters();
                        await WriteSingleIntAsync(plcHost, plcPort, station, d624Address, 0, token).ConfigureAwait(true);
                        await WriteSingleIntAsync(plcHost, plcPort, station, d623Address, ok ? 1 : 2, token).ConfigureAwait(true);
                        SetStatus(ok ? "自动识别并发送成功，等待下一次 D626 变化..." : "自动识别失败，等待下一次 D626 变化...");
                        AppLogger.Instance.Info($"识别完成: {(ok ? "成功" : "失败")}, D623={_d623Value}, D624=0");
                    }
                    catch (TimeoutException ex)
                    {
                        AppLogger.Instance.Error($"Modbus 操作超时: {ex.Message}", ex);
                        SetStatus($"操作超时：{ex.Message}");
                        _d623Value = 2;
                        RefreshPlcRegisters();
                        try { await WriteSingleIntAsync(plcHost, plcPort, station, d624Address, 0, token).ConfigureAwait(true); } catch { }
                        try { await WriteSingleIntAsync(plcHost, plcPort, station, d623Address, 2, token).ConfigureAwait(true); } catch { }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                AppLogger.Instance.Warn("自动化任务被取消");
                _d623Value = 0;
                RefreshPlcRegisters();
                try { await WriteSingleIntAsync(plcHost, plcPort, station, d624Address, 0, token).ConfigureAwait(true); } catch { }
                try { await WriteSingleIntAsync(plcHost, plcPort, station, d623Address, 0, token).ConfigureAwait(true); } catch { }
            }
            catch (Exception ex)
            {
                AppLogger.Instance.Error($"自动化任务发生未知错误: {ex.Message}", ex);
                _d623Value = 0;
                RefreshPlcRegisters();
                try { await WriteSingleIntAsync(plcHost, plcPort, station, d624Address, 0, token).ConfigureAwait(true); } catch { }
                try { await WriteSingleIntAsync(plcHost, plcPort, station, d623Address, 0, token).ConfigureAwait(true); } catch { }
            }
            finally
            {
                AppLogger.Instance.Info("停止心跳线程");
                await StopHeartbeatAsync().ConfigureAwait(true);
                AppLogger.Instance.Info("写入 D624=0，确保清理状态");
                try { await WriteSingleIntAsync(plcHost, plcPort, station, d624Address, 0, token).ConfigureAwait(true); } catch { }
                AppLogger.Instance.LogOperation("自动化结束", "RunFullAutomationAsync 执行完成");
            }
        }

        private async Task<bool> RecognizeAndSendAsync(bool sendToPlc, bool writeHeartbeatOnly)
        {
            // 前置检查必须在 UI 线程
            if (_projectDoc is null || string.IsNullOrWhiteSpace(_projectFile))
            {
                await RunOnUiThreadAsync(() =>
                {
                    StatusText.Text = "请先导入工程 DXF。";
                    return Task.CompletedTask;
                }).ConfigureAwait(false);
                return false;
            }
            if (_stage1MoldFiles.Count == 0 || _stage2MoldFiles.Count == 0)
            {
                await RunOnUiThreadAsync(() =>
                {
                    StatusText.Text = "请先分别导入台1模具和台2模具。";
                    return Task.CompletedTask;
                }).ConfigureAwait(false);
                return false;
            }

            // 准备数据（在 UI 线程获取 ComboBox 选择和 BoardWidth）
            string? selectedStage1File = null;
            string? selectedStage2File = null;
            double boardWidthValue = 150;
            LoadingDialog? loadingDialog = null;

            // 显示加载对话框（确保在 finally 中能关闭）
            await RunOnUiThreadAsync(() =>
            {
                selectedStage1File = ResolveSelectedMoldFile(Stage1MoldComboBox, _stage1MoldFiles);
                selectedStage2File = ResolveSelectedMoldFile(Stage2MoldComboBox, _stage2MoldFiles);
                boardWidthValue = ReadBoardWidth(); // 在 UI 线程读取
                loadingDialog = new LoadingDialog { Owner = this };
                loadingDialog.SetMessage("正在识别图纸...");
                loadingDialog.Show();
                return Task.CompletedTask;
            }).ConfigureAwait(false);

            // 确保 loadingDialog 始终被关闭
            try
            {
                // 在后台线程执行重型识别计算
                bool ok;
                try
                {
                    // 使用 Task.Run 在后台执行，用 Dispatcher 更新 UI
                    ok = await Task.Run(async () =>
                    {
                        Dispatcher.Invoke(() => loadingDialog?.SetMessage("正在解析 CAD 图纸..."));

                        var project = DxfAnalyzer.ExtractProject(_projectDoc!);
                        _lastProjectProfile = project;
                        _lastOuterContourPoints = DxfAnalyzer.ExtractOuterContourForDebug(_projectDoc!);

                        Dispatcher.Invoke(() => loadingDialog?.SetMessage("正在匹配模具..."));

                        var splitY = project.OuterRectangle.MinY + boardWidthValue;
                        var stage1Project = new ProjectProfile(project.OuterRectangle, project.Holes.Where(h => h.Centroid.Y < splitY).ToList(), project.CornerCandidates, project.EdgeCandidates.Where(e => e.Centroid.Y < splitY).ToList(), project.CornerStepPaths, project.ContourPaths, project.Stage1ContourPaths, []);
                        var stage2Project = new ProjectProfile(project.OuterRectangle, project.Holes.Where(h => h.Centroid.Y >= splitY).ToList(), project.CornerCandidates, project.EdgeCandidates.Where(e => e.Centroid.Y >= splitY).ToList(), project.CornerStepPaths, project.ContourPaths, [], project.Stage2ContourPaths);
                        var stage1Files = selectedStage1File is null ? _stage1MoldFiles.ToList() : _stage1MoldFiles.Where(f => string.Equals(f, selectedStage1File, StringComparison.OrdinalIgnoreCase)).Concat(_stage1MoldFiles.Where(f => !string.Equals(f, selectedStage1File, StringComparison.OrdinalIgnoreCase))).ToList();
                        var stage2Files = selectedStage2File is null ? _stage2MoldFiles.ToList() : _stage2MoldFiles.Where(f => string.Equals(f, selectedStage2File, StringComparison.OrdinalIgnoreCase)).Concat(_stage2MoldFiles.Where(f => !string.Equals(f, selectedStage2File, StringComparison.OrdinalIgnoreCase))).ToList();
                        var stage1Molds = stage1Files.Select((f, idx) => { TryGetDocumentFromCache(f, out var doc); return DxfAnalyzer.ExtractMold(1 + idx, f, doc); }).ToList();
                        var stage2Molds = stage2Files.Select((f, idx) => { TryGetDocumentFromCache(f, out var doc); return DxfAnalyzer.ExtractMold(1 + idx, f, doc); }).ToList();
                        _lastMolds = stage1Molds.Concat(stage2Molds).ToList();

                        Dispatcher.Invoke(() => loadingDialog?.SetMessage("正在匹配台1模具..."));

                        var matcher = new MoldMatcher();
                        var stage1Result = matcher.Match(stage1Project, stage1Molds, isStage1: true);

                        Dispatcher.Invoke(() => loadingDialog?.SetMessage("正在匹配台2模具..."));

                        var stage2Result = matcher.Match(stage2Project, stage2Molds, isStage1: false);
                        var matchResult = new MatchResult(stage1Result.HoleAssignments.Concat(stage2Result.HoleAssignments).ToList(), stage1Result.GuidePaths ?? stage2Result.GuidePaths);
                        _lastMatchResult = matchResult;

                        // UI 渲染需要在主线程
                        Dispatcher.Invoke(() =>
                        {
                            loadingDialog?.SetMessage("正在渲染结果...");
                            RenderStageResult(stage1Result, stage1Molds, isStage1: true);
                            RenderStageResult(stage2Result, stage2Molds, isStage1: false);
                            RenderPreview(_projectDoc!, _projectFile!, withAnnotation: true);
                            StatusText.Text = $"识别完成：台1 {stage1Result.HoleAssignments.Count} 个，台2 {stage2Result.HoleAssignments.Count} 个。";
                        });

                        if (sendToPlc)
                        {
                            Dispatcher.Invoke(() => loadingDialog?.SetMessage("正在发送结果到 PLC..."));
                            await SendRecognitionResultAsync(stage1Result, stage2Result, boardWidthValue).ConfigureAwait(false);
                        }

                        return true;
                    }).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    AppLogger.Instance.Error($"识别过程发生异常: {ex.Message}", ex);
                    Dispatcher.Invoke(() => StatusText.Text = $"识别失败: {ex.Message}");
                    ok = false;
                }

                return ok;
            }
            finally
            {
                // 确保关闭加载对话框
                Dispatcher.Invoke(() =>
                {
                    if (loadingDialog != null)
                    {
                        loadingDialog.Close();
                        loadingDialog = null;
                    }
                });
            }
        }

        private async Task SendRecognitionResultAsync(MatchResult stage1Result, MatchResult stage2Result, double boardWidth)
        {
            var model = BuildTcpExportModel(boardWidth);
            var host = TcpExportDialog.SharedTcpHost;
            var port = int.TryParse(TcpExportDialog.SharedTcpPort, out var parsedPort) ? parsedPort : 502;
            byte station = byte.TryParse(TcpExportDialog.SharedModbusStation, out var parsedStation) ? parsedStation : (byte)1;
            var registerAddress = TcpExportDialog.SharedModbusRegisterAddress;
            var encoding = TcpExportDialog.SharedEncoding;
            await _modbusTcpCommService.SendExportModelAsync(host, port, station, registerAddress, model, encoding).ConfigureAwait(true);
        }

        private Task<ModbusTcpNet> GetPlcClientAsync(string host, int port, byte station)
        {
            var key = $"{host}:{port}:{station}";
            if (_plcClient is not null && string.Equals(_plcClientKey, key, StringComparison.Ordinal) && _plcIsConnected)
            {
                return Task.FromResult(_plcClient);
            }

            _plcClient?.Dispose();
            _plcClient = new ModbusTcpNet(host, port, station) { AddressStartWithZero = true };
            _plcClientKey = key;
            _plcIsConnected = true;
            return Task.FromResult(_plcClient);
        }

        private void MarkPlcDisconnected()
        {
            _plcIsConnected = false;
            _lastPlcReconnectAttempt = DateTime.Now;
        }

        private bool ShouldReconnectPlc()
        {
            return !_plcIsConnected && (DateTime.Now - _lastPlcReconnectAttempt).TotalSeconds >= 5;
        }

        private static string NormalizePlcAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                return string.Empty;
            }

            var trimmed = address.Trim();
            return trimmed.StartsWith("D", StringComparison.OrdinalIgnoreCase) ? trimmed.Substring(1) : trimmed;
        }

        private async Task WriteSingleIntAsync(string host, int port, byte station, string address, int value, CancellationToken token)
        {
            await WithModbusTimeoutAsync($"Write[{address}]={value}", address, async ct =>
            {
                token.ThrowIfCancellationRequested();
                var client = await GetPlcClientAsync(host, port, station).ConfigureAwait(true);
                var plcAddress = NormalizePlcAddress(address);
                var wordValue = value < short.MinValue ? short.MinValue : value > short.MaxValue ? short.MaxValue : (short)value;
                await _plcIoLock.WaitAsync(ct).ConfigureAwait(true);
                try
                {
                    var result = await client.WriteAsync(plcAddress, new[] { wordValue }).ConfigureAwait(true);
                    if (result is null || !result.IsSuccess) throw new InvalidOperationException(result?.Message ?? $"写入 {plcAddress} 失败");
                }
                finally
                {
                    _plcIoLock.Release();
                }
                return true;
            }, timeoutMs: 3000, token: token).ConfigureAwait(true);
        }

        /// <summary>
        /// 心跳写入专用，不记录详细日志。
        /// </summary>
        private async Task WriteSingleIntSilentAsync(string host, int port, byte station, string address, int value, CancellationToken token)
        {
            try
            {
                var client = await GetPlcClientAsync(host, port, station).ConfigureAwait(false);
                var plcAddress = NormalizePlcAddress(address);
                var wordValue = value < short.MinValue ? short.MinValue : value > short.MaxValue ? short.MaxValue : (short)value;

                using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(token);
                timeoutCts.CancelAfter(500);

                bool lockAcquired = false;
                try
                {
                    await _plcIoLock.WaitAsync(timeoutCts.Token).ConfigureAwait(false);
                    lockAcquired = true;
                    await client.WriteAsync(plcAddress, new[] { wordValue }).ConfigureAwait(false);
                }
                finally
                {
                    if (lockAcquired)
                    {
                        _plcIoLock.Release();
                    }
                }
            }
            catch
            {
            }
        }

        private async Task<string> ReadD600FileNameAsync(string host, int port, byte station, string address, CancellationToken token)
        {
            return await WithModbusTimeoutAsync("ReadString[D600]", address, async ct =>
            {
                token.ThrowIfCancellationRequested();
                var client = await GetPlcClientAsync(host, port, station).ConfigureAwait(true);
                var plcAddress = NormalizePlcAddress(address);

                try
                {
                    await _plcIoLock.WaitAsync(ct).ConfigureAwait(true);

                    var raw = await client.ReadAsync(plcAddress, 20).ConfigureAwait(true);
                    if (!raw.IsSuccess)
                    {
                        raw = await client.ReadAsync(address, 20).ConfigureAwait(true);
                    }

                    if (!raw.IsSuccess)
                    {
                        throw new InvalidOperationException(raw.Message);
                    }

                    var text = DecodePlcString(raw.Content);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        UpdatePlcStringValue(600, text);
                        return text;
                    }

                    return _d600Value;
                }
                catch (FormatException ex)
                {
                    AppLogger.Instance.Warn($"D600 字符串解析失败，视为空文件名: {ex.Message}");
                    UpdatePlcStringValue(600, string.Empty);
                    return string.Empty;
                }
                finally
                {
                    _plcIoLock.Release();
                }
            }, timeoutMs: 5000, token: token).ConfigureAwait(true);
        }

        private static string DecodePlcString(byte[] registerBytes)
        {
            if (registerBytes is not { Length: > 0 })
            {
                return string.Empty;
            }

            var raw = new byte[registerBytes.Length];
            Buffer.BlockCopy(registerBytes, 0, raw, 0, raw.Length);

            var swapped = new byte[registerBytes.Length];
            Buffer.BlockCopy(registerBytes, 0, swapped, 0, swapped.Length);
            SwapRegisterBytes(swapped);

            if (TryDecode(swapped, Encoding.UTF8, out var swappedUtf8Text))
            {
                return swappedUtf8Text;
            }

            if (TryDecode(raw, Encoding.UTF8, out var utf8Text))
            {
                return utf8Text;
            }

            if (TryDecode(swapped, Encoding.GetEncoding("gb2312"), out var swappedGb2312Text))
            {
                return swappedGb2312Text;
            }

            if (TryDecode(raw, Encoding.GetEncoding("gb2312"), out var gb2312Text))
            {
                return gb2312Text;
            }

            if (TryDecode(swapped, Encoding.GetEncoding("gbk"), out var swappedGbkText))
            {
                return swappedGbkText;
            }

            if (TryDecode(raw, Encoding.GetEncoding("gbk"), out var gbkText))
            {
                return gbkText;
            }

            if (TryDecode(swapped, Encoding.ASCII, out var swappedAsciiText))
            {
                return swappedAsciiText;
            }

            if (TryDecode(raw, Encoding.ASCII, out var asciiText))
            {
                return asciiText;
            }

            return string.Empty;
        }

        private static bool TryDecode(byte[] bytes, Encoding encoding, out string text)
        {
            try
            {
                text = encoding.GetString(bytes).Trim('\0', ' ', '\r', '\n');
                return !string.IsNullOrWhiteSpace(text) && text.IndexOf('\uFFFD') < 0;
            }
            catch
            {
                text = null;
                return false;
            }
        }

        private static void SwapRegisterBytes(byte[] data)
        {
            if (data.Length < 2)
            {
                return;
            }

            for (var i = 0; i < data.Length; i += 2)
            {
                var high = data[i];
                data[i] = data[i + 1];
                data[i + 1] = high;
            }
        }

        private async Task<T> WithModbusTimeoutAsync<T>(string operation, string address, Func<CancellationToken, Task<T>> action, int timeoutMs = 5000, CancellationToken token = default, bool silent = false)
        {
            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(token);
            timeoutCts.CancelAfter(timeoutMs);

            bool lockAcquired = false;
            try
            {
                var result = await action(timeoutCts.Token).ConfigureAwait(true);
                lockAcquired = true;
                return result;
            }
            catch (OperationCanceledException) when (!token.IsCancellationRequested)
            {
                AppLogger.Instance.LogModbus("超时", $"{operation}@{address}", error: $"{timeoutMs}ms超时");
                MarkPlcDisconnected();
                throw new TimeoutException($"Modbus操作 {operation} @ {address} 超时（{timeoutMs}ms）");
            }
            catch (System.Net.Sockets.SocketException ex)
            {
                AppLogger.Instance.LogModbus("断链", $"{operation}@{address}", error: ex.Message);
                MarkPlcDisconnected();
                throw new TimeoutException($"PLC连接失败：{ex.Message}", ex);
            }
            catch (Exception ex) when (ex.Message.Contains("输入字符串的格式不正确", StringComparison.OrdinalIgnoreCase)
                                      || ex.Message.Contains("已关闭 Safe handle", StringComparison.OrdinalIgnoreCase)
                                      || ex.Message.Contains("将指定的计数添加到该信号量中会导致其超过最大计数", StringComparison.OrdinalIgnoreCase)
                                      || ex.Message.Contains("Safe handle", StringComparison.OrdinalIgnoreCase))
            {
                AppLogger.Instance.LogModbus("断链", $"{operation}@{address}", error: ex.Message);
                MarkPlcDisconnected();
                throw new TimeoutException($"PLC连接已断开：{ex.Message}", ex);
            }
            catch (Exception ex) when (ex.Message.Contains("timeout", StringComparison.OrdinalIgnoreCase)
                                      || ex.Message.Contains("连接", StringComparison.OrdinalIgnoreCase)
                                      || ex.Message.Contains("connect", StringComparison.OrdinalIgnoreCase)
                                      || ex.Message.Contains("无法连接", StringComparison.OrdinalIgnoreCase)
                                      || ex.Message.Contains("积极拒绝", StringComparison.OrdinalIgnoreCase)
                                      || ex.Message.Contains("拒绝连接", StringComparison.OrdinalIgnoreCase)
                                      || ex.Message.Contains("No connection could be made", StringComparison.OrdinalIgnoreCase))
            {
                AppLogger.Instance.LogModbus("断链", $"{operation}@{address}", error: ex.Message);
                MarkPlcDisconnected();
                throw new TimeoutException($"PLC连接失败：{ex.Message}", ex);
            }
            catch (Exception ex)
            {
                AppLogger.Instance.LogModbus("失败", $"{operation}@{address}", error: ex.Message);
                throw;
            }
        }

        private async Task<int> ReadIntAsync(string host, int port, byte station, string address, int logicalRegister, CancellationToken token)
        {
            return await WithModbusTimeoutAsync($"ReadInt[{logicalRegister}]", address, async ct =>
            {
                token.ThrowIfCancellationRequested();
                var client = await GetPlcClientAsync(host, port, station).ConfigureAwait(true);
                var plcAddress = NormalizePlcAddress(address);
                await _plcIoLock.WaitAsync(ct).ConfigureAwait(true);
                try
                {
                    var result = await client.ReadInt16Async(plcAddress, 1).ConfigureAwait(true);
                    if (!result.IsSuccess) throw new InvalidOperationException(result.Message);
                    var value = result.Content is short[] values && values.Length > 0 ? values[0] : 0;
                    UpdatePlcIntValue(logicalRegister, value);
                    return value;
                }
                finally
                {
                    _plcIoLock.Release();
                }
            }, timeoutMs: 5000, token: token).ConfigureAwait(true);
        }

        private async Task<double> ReadPlcBoundaryAsync(string host, int port, byte station, string address, CancellationToken token)
        {
            return await ReadPlcBoundaryWithRetryAsync(host, port, station, address, retryCount: 2, timeoutMs: 10000, token).ConfigureAwait(true);
        }

        private async Task<double> ReadPlcBoundaryWithRetryAsync(string host, int port, byte station, string address, int retryCount, int timeoutMs, CancellationToken token)
        {
            if (retryCount < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(retryCount));
            }

            for (var attempt = 1; ; attempt++)
            {
                try
                {
                    AppLogger.Instance.Debug($"读取识图界限: {address}，第 {attempt}/{retryCount} 次");
                    return await ReadPlcBoundaryCoreAsync(host, port, station, address, timeoutMs, token).ConfigureAwait(true);
                }
                catch (TimeoutException ex)
                {
                    AppLogger.Instance.Warn($"读取识图界限超时: {address}，第 {attempt}/{retryCount} 次，{ex.Message}");
                    if (attempt >= retryCount)
                    {
                        AppLogger.Instance.Error($"读取识图界限失败: {address}，已达到最大重试次数");
                        throw;
                    }

                    try
                    {
                        await Task.Delay(TimeSpan.FromMilliseconds(200), token).ConfigureAwait(true);
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch
                    {
                        // 忽略暂停期间的异常，继续重试
                    }
                }
            }
        }

        private async Task<double> ReadPlcBoundaryCoreAsync(string host, int port, byte station, string address, int timeoutMs, CancellationToken token)
        {
            return await WithModbusTimeoutAsync("ReadFloat[D620]", address, async ct =>
            {
                token.ThrowIfCancellationRequested();
                var client = await GetPlcClientAsync(host, port, station).ConfigureAwait(true);
                var plcAddress = NormalizePlcAddress(address);
                await _plcIoLock.WaitAsync(ct).ConfigureAwait(true);
                try
                {
                    var result = await client.ReadFloatAsync(plcAddress, 1).ConfigureAwait(true);
                    if (!result.IsSuccess) throw new InvalidOperationException(result.Message);

                    var values = result.Content;
                    var value = values is { Length: > 0 } ? values[0] : _d620Value;
                    UpdatePlcFloatValue(620, value);
                    return value;
                }
                finally
                {
                    _plcIoLock.Release();
                }
            }, timeoutMs: timeoutMs, token: token).ConfigureAwait(true);
        }

        private async Task<int> WaitForRegisterChangeAsync(string host, int port, byte station, string address, int initialValue, int logicalRegister, CancellationToken token)
        {
            var pollCount = 0;
            while (true)
            {
                token.ThrowIfCancellationRequested();
                var current = await ReadIntAsync(host, port, station, address, logicalRegister, token).ConfigureAwait(true);
                pollCount++;
                if (current != initialValue)
                {
                    AppLogger.Instance.Info($"【寄存器变化】{address}: {initialValue} -> {current}");
                    SetStatus($"检测到 {address} 变化：{initialValue} -> {current}");
                    return current;
                }

                if (pollCount % 25 == 0)
                {
                    AppLogger.Instance.Debug($"等待 {address} 变化，当前值={current}");
                    SetStatus($"正在监听 {address}，当前值={current}，等待变化...");
                }

                await Task.Delay(200, token).ConfigureAwait(true);
            }
        }

        private void UpdatePlcStringValue(int register, string value)
        {
            if (register == 600)
            {
                _d600Value = value;
                RefreshPlcRegisters();
            }
        }

        private void UpdatePlcFloatValue(int register, double value)
        {
            if (register == 620)
            {
                _d620Value = value;
                _boardWidth = Math.Max(0, value);

                if (BoardWidthTextBox is not null)
                {
                    if (Dispatcher.CheckAccess())
                    {
                        BoardWidthTextBox.Text = value.ToString("0.###", CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        Dispatcher.Invoke(() => BoardWidthTextBox.Text = value.ToString("0.###", CultureInfo.InvariantCulture));
                    }
                }

                RefreshPlcRegisters();
            }
        }

        private void UpdatePlcIntValue(int register, int value)
        {
            var previousValue = 0;
            var changed = false;
            switch (register)
            {
                case 622:
                    if (_d622Value != value) changed = true;
                    _d622Value = value; break;
                case 623:
                    if (_d623Value != value) changed = true;
                    _d623Value = value; break;
                case 624:
                    if (_d624Value != value) changed = true;
                    _d624Value = value; break;
                case 625:
                    if (_d625Value != value) changed = true;
                    _d625Value = value; break;
                case 626:
                    if (_d626Value != value) changed = true;
                    _d626Value = value; break;
                default: return;
            }

            if (changed)
            {
                AppLogger.Instance.Info($"【PLC {register} 变化】{previousValue} -> {value}");
            }

            RefreshPlcRegisters();
        }

        private async Task UpdatePlcSnapshotAsync(string host, int port, byte station, string d600Address, string d620Address, string d622Address, string d623Address, string d624Address, string d625Address, string d626Address, CancellationToken token)
        {
            _d600Value = await ReadD600FileNameAsync(host, port, station, d600Address, token).ConfigureAwait(true);
            _d620Value = await ReadPlcBoundaryAsync(host, port, station, d620Address, token).ConfigureAwait(true);
            _d622Value = await ReadIntAsync(host, port, station, d622Address, 622, token).ConfigureAwait(true);
            _d623Value = await ReadIntAsync(host, port, station, d623Address, 623, token).ConfigureAwait(true);
            _d624Value = await ReadIntAsync(host, port, station, d624Address, 624, token).ConfigureAwait(true);
            _d626Value = await ReadIntAsync(host, port, station, d626Address, 626, token).ConfigureAwait(true);
            _d625Value = 1;
            RefreshPlcRegisters();
        }

        private void PlcRegisterGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.Row.Item is not PlcRegisterRow row)
            {
                return;
            }

            try
            {
                if (e.EditingElement is System.Windows.Controls.TextBox editor)
                {
                    var newText = editor.Text?.Trim();
                    if (!string.IsNullOrWhiteSpace(newText))
                    {
                        row.Address = newText;
                        RefreshPlcRegisters();
                    }
                }
            }
            catch
            {
            }
        }

        private async Task StartHeartbeatIfNeededAsync()
        {
            if (_heartbeatTask is not null && !_heartbeatTask.IsCompleted)
            {
                return;
            }

            await StartHeartbeatAsync(TcpExportDialog.SharedTcpHost, int.TryParse(TcpExportDialog.SharedTcpPort, out var parsedPort) ? parsedPort : 502, byte.TryParse(TcpExportDialog.SharedModbusStation, out var parsedStation) ? parsedStation : (byte)1, "625", CancellationToken.None).ConfigureAwait(true);
        }

        private async Task StartHeartbeatAsync(string host, int port, byte station, string register, CancellationToken token)
        {
            _heartbeatCts = CancellationTokenSource.CreateLinkedTokenSource(token);
            var heartbeatToken = _heartbeatCts.Token;
            _heartbeatTask = Task.Run(async () =>
            {
                while (!heartbeatToken.IsCancellationRequested)
                {
                    try
                    {
                        await WriteSingleIntSilentAsync(host, port, station, register, 1, heartbeatToken).ConfigureAwait(false);
                        _d625Value = 1;
                        RefreshPlcRegisters();
                    }
                    catch
                    {
                    }

                    try
                    {
                        await Task.Delay(1000, heartbeatToken).ConfigureAwait(false);
                    }
                    catch (OperationCanceledException)
                    {
                        break;
                    }
                }
            }, heartbeatToken);
            await Task.CompletedTask;
        }

        private async Task StopHeartbeatAsync()
        {
            _heartbeatCts?.Cancel();
            if (_heartbeatTask is not null)
            {
                try { await _heartbeatTask.ConfigureAwait(true); } catch { }
            }

            _heartbeatCts?.Dispose();
            _heartbeatCts = null;
            _heartbeatTask = null;
        }

        private TcpExportModel BuildTcpExportModel(double boardWidth)
        {
            var model = new TcpExportModel();
            model.ProgramName = System.IO.Path.GetFileNameWithoutExtension(_projectFile);
            model.ProgramNo = System.IO.Path.GetFileNameWithoutExtension(_projectFile);
            model.LeftRightDoor = 0;
            model.Material = 0;
            model.Type = 0;
            model.FormingLength = 0;
            model.FormingWidth = 0;
            model.FormingThickness = 0;
            model.PlateLength = _lastProjectProfile?.OuterRectangle.Width ?? 0;
            model.PlateWidth = _lastProjectProfile?.OuterRectangle.Height ?? 0;
            model.PlateWidth2 = boardWidth > 0 ? boardWidth
                : (_lastProjectProfile?.OuterRectangle.Height ?? 0);
            model.PlateThickness = 0;
            model.Spare2 = 0;
            model.Spare3 = 0;
            model.Spare4 = 0;
            model.CustomContent = string.Empty;

            var boundaryWidth = model.PlateWidth2 > 0 ? model.PlateWidth2 : (_lastProjectProfile?.OuterRectangle.Height ?? 0);
            var stage1Rows = _stage1PositionRows.OrderBy(r => r.Index).ToList();
            var stage2Rows = _stage2PositionRows.OrderBy(r => r.Index).ToList();
            if (stage1Rows.Count == 0 && stage2Rows.Count == 0 && boundaryWidth > 0)
            {
                stage1Rows = [];
                stage2Rows = [];
            }

            model.Stage1PunchCount = stage1Rows.Count;
            model.Stage2PunchCount = stage2Rows.Count;
            model.Stage1DiagramCoordinates = stage1Rows.Select(r => new TcpCoordinateRow { X = r.PosX, Y = r.PosY }).ToList();
            model.Stage2DiagramCoordinates = stage2Rows.Select(r => new TcpCoordinateRow { X = r.PosX, Y = r.PosY }).ToList();
            model.Stage1PositionMoldIds = stage1Rows.Select(r => ToStationMoldCode(r.MoldId, "M")).ToList();
            model.Stage2PositionMoldIds = stage2Rows.Select(r => ToStationMoldCode(r.MoldId, "N")).ToList();
            model.Stage1PunchMoldIds = stage1Rows.Select(r => ToStationMoldCode(r.MoldId, "M")).ToList();
            model.Stage2PunchMoldIds = stage2Rows.Select(r => ToStationMoldCode(r.MoldId, "N")).ToList();
            return model;
        }

        private static string ToStationMoldCode(int moldId, string prefix)
        {
            if (moldId <= 0) return string.Empty;
            return $"{prefix}{moldId:D2}";
        }

        private RectBounds GetRecognitionBoundary()
        {
            if (_lastProjectProfile is not null)
            {
                return _lastProjectProfile.OuterRectangle;
            }

            if (_stage1PositionRows.Count > 0 || _stage2PositionRows.Count > 0)
            {
                var all = _stage1PositionRows.Concat(_stage2PositionRows).ToList();
                var minX = all.Min(x => x.PosX);
                var minY = all.Min(x => x.PosY);
                var maxX = all.Max(x => x.PosX);
                var maxY = all.Max(x => x.PosY);
                return new RectBounds(minX, minY, maxX, maxY);
            }

            return new RectBounds(0, 0, 1, 1);
        }

        private IEnumerable<PositionRow> SplitRowsByBoundary(IEnumerable<PositionRow> rows, RectBounds boundary, bool upperHalf)
        {
            var midpoint = boundary.MinY + boundary.Height / 2.0;
            foreach (var row in rows.OrderBy(x => x.Index))
            {
                var isUpper = row.PosY >= midpoint;
                if (upperHalf == isUpper)
                {
                    yield return row;
                }
            }
        }

        private static string BuildMoldSizeText(HoleFeature feature)
        {
            static string FormatDim(double value)
            {
                var rounded = Math.Round(value, 1, MidpointRounding.AwayFromZero);
                return Math.Abs(rounded % 1.0) < 1e-9 ? $"{rounded:0}" : $"{rounded:0.0}";
            }

            if (MoldMatcher.IsCircleLike(feature))
            {
                var dia = (feature.Width + feature.Height) / 2.0;
                return $"φ{FormatDim(dia)}";
            }

            return $"{FormatDim(feature.Width)}*{FormatDim(feature.Height)}";
        }

        private void RenderStageResult(MatchResult result, IReadOnlyList<MoldProfile> molds, bool isStage1)
        {
            var moldRows = isStage1 ? _stage1MoldRows : _stage2MoldRows;
            var positionRows = isStage1 ? _stage1PositionRows : _stage2PositionRows;
            moldRows.Clear();
            positionRows.Clear();

            var useCounter = result.HoleAssignments
                .Where(x => !x.Hole.HoleType.StartsWith("EdgeNotch:", StringComparison.Ordinal))
                .GroupBy(x => x.MoldId)
                .ToDictionary(g => g.Key, g => g.Count());

            var moldPrefix = isStage1 ? "M" : "N";
            foreach (var mold in molds.OrderBy(x => x.MoldId))
            {
                useCounter.TryGetValue(mold.MoldId, out var count);
                moldRows.Add(new MoldRow
                {
                    MoldPreview = BuildMoldPreview(mold.FilePath),
                    MoldCode = $"{moldPrefix}{mold.MoldId:D2}",
                    MoldName = BuildMoldSizeText(mold.Feature),
                    UsedCount = count,
                    MatchType = mold.MoldId == 1 ? "角落连续冲压" : "单次冲压",
                    Remark = mold.MoldId == 1 ? "仅四角孔洞" : "普通孔洞"
                });
            }

            var index = 1;
            foreach (var row in result.HoleAssignments)
            {
                if (row.Hole.HoleType.StartsWith("EdgeNotch:", StringComparison.Ordinal))
                {
                    continue;
                }

                positionRows.Add(new PositionRow
                {
                    Index = index++,
                    HoleType = row.Hole.HoleType,
                    MoldId = row.MoldId,
                    MoldCode = row.MoldId > 0 ? $"{(isStage1 ? "M" : "N")}{row.MoldId:D2}" : "未匹配",
                    PosX = Math.Round(row.Hole.Centroid.X - _lastProjectProfile!.OuterRectangle.MinX, 2),
                    PosY = Math.Round(row.Hole.Centroid.Y - _lastProjectProfile!.OuterRectangle.MinY, 2),
                    AbsX = row.Hole.Centroid.X,
                    AbsY = row.Hole.Centroid.Y,
                    PositionRelation = isStage1 ? "台1区域" : "台2区域",
                    IsCornerCandidate = row.IsCornerCandidate ? "是" : "否",
                    IsEdgeHole = row.IsEdgeHole ? "是" : "否",
                    TopCandidates = row.TopCandidates,
                    AreaRatio = row.AreaRatioInfo,
                    FailureReason = row.FailureReason,
                    RotationDeg = row.RotationDeg,
                    IsMirrored = row.IsMirrored
                });
            }
        }

        private void RenderLegend(IReadOnlyList<MoldProfile> molds, bool isStage1)
        {
            Stage1LegendPanel.Children.Clear();
            Stage2LegendPanel.Children.Clear();
            foreach (var mold in molds.OrderBy(x => x.MoldId))
            {
                var color = InteractiveDxfPreview.GetMoldColor(mold.MoldId);
                var chip = new Border
                {
                    Background = new SolidColorBrush(WpfColor.FromRgb(34, 34, 34)),
                    BorderBrush = new SolidColorBrush(WpfColor.FromRgb(64, 64, 64)),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(4),
                    Margin = new Thickness(0, 0, 6, 6),
                    Padding = new Thickness(6, 3, 6, 3)
                };

                var panel = new StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal };
                panel.Children.Add(new Border
                {
                    Width = 10,
                    Height = 10,
                    Background = new SolidColorBrush(color),
                    Margin = new Thickness(0, 0, 6, 0),
                    CornerRadius = new CornerRadius(2)
                });
                var legendPrefix = isStage1 ? "M" : "N";
                panel.Children.Add(new TextBlock
                {
                    Text = $"{legendPrefix}{mold.MoldId:D2}",
                    Foreground = WpfBrushes.White,
                    FontSize = 9,
                    VerticalAlignment = VerticalAlignment.Center
                });
                chip.Child = panel;
                if (mold.MoldId == 1)
                {
                    Stage1LegendPanel.Children.Add(chip);
                }
                else
                {
                    Stage2LegendPanel.Children.Add(chip);
                }
            }
        }

        private void RefreshFileList()
        {
            FileTreeView.Items.Clear();

            var root = CreateFileTreeItem("图纸列表", isExpanded: true);

            var projectNode = CreateFileTreeItem("工程图", isExpanded: true);
            if (!string.IsNullOrWhiteSpace(_projectFile))
            {
                projectNode.Items.Add(CreateFileTreeItem(System.IO.Path.GetFileName(_projectFile), _projectFile));
            }

            var stage1Node = CreateFileTreeItem("台1模具", isExpanded: true);
            foreach (var file in _stage1MoldFiles)
            {
                stage1Node.Items.Add(CreateFileTreeItem(System.IO.Path.GetFileName(file), file));
            }

            var stage2Node = CreateFileTreeItem("台2模具", isExpanded: true);
            foreach (var file in _stage2MoldFiles)
            {
                stage2Node.Items.Add(CreateFileTreeItem(System.IO.Path.GetFileName(file), file));
            }

            root.Items.Add(projectNode);
            root.Items.Add(stage1Node);
            root.Items.Add(stage2Node);
            FileTreeView.Items.Add(root);
        }

        private static TreeViewItem CreateFileTreeItem(string text, string? tag = null, bool isExpanded = false)
        {
            return new TreeViewItem
            {
                Header = new TextBlock
                {
                    Text = text,
                    Foreground = Brushes.White
                },
                Tag = tag,
                IsExpanded = isExpanded,
                Foreground = Brushes.White
            };
        }

        private void RenderPreview(DxfDocument doc, string? path, bool withAnnotation)
        {
            _previewPlugin.CreatePreview(doc, _viewer);
            if (!string.IsNullOrWhiteSpace(path) && path == _projectFile && _lastProjectProfile is not null)
            {
                var splitY = _lastProjectProfile.OuterRectangle.MinY + _boardWidth;
                _viewer.RenderCornerContours(
                    _lastProjectProfile.OuterRectangle,
                    _lastOuterContourPoints,
                    withAnnotation ? _lastMatchResult?.GuidePaths : null,
                    _lastProjectProfile.CornerCandidates,
                    _boardWidth,
                    splitY);
                _viewer.RenderPendingEdgeRecognitionHoles(_lastProjectProfile.EdgeCandidates);

                if (withAnnotation && _lastMatchResult is not null)
                {
                    _viewer.RenderAnnotations(_lastMatchResult.HoleAssignments, _lastMolds, _lastProjectProfile.OuterRectangle.MinY + _boardWidth);
                }
                else
                {
                    _viewer.RenderAnnotations([], [], null);
                }
            }
            else
            {
                _viewer.RenderCornerContours(null, null, null, null, 0, null);
                _viewer.RenderAnnotations([], [], null);
            }
            PreviewHintText.Visibility = Visibility.Collapsed;
        }

        private void FileTreeView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (FileTreeView.SelectedItem is not TreeViewItem item || item.Tag is not string path)
            {
                return;
            }
            if (!TryGetDocumentFromCache(path, out var doc))
            {
                return;
            }
            var showAnnotation = _lastMatchResult is not null && path == _projectFile;
            RenderPreview(doc, path, showAnnotation);
            StatusText.Text = $"预览图纸：{System.IO.Path.GetFileName(path)}";
        }

        private void OpenFileLocation_Click(object sender, RoutedEventArgs e)
        {
            if (FileTreeView.SelectedItem is not TreeViewItem item || item.Tag is not string path)
            {
                return;
            }
            if (!System.IO.File.Exists(path))
            {
                System.Windows.MessageBox.Show($"文件不存在：{path}", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            var directory = System.IO.Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory))
            {
                System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{path}\"");
            }
        }

        private void PositionGrid_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is not System.Windows.Controls.DataGrid grid || grid.SelectedItem is not PositionRow row)
            {
                return;
            }
            _viewer.FocusHole(row.AbsX, row.AbsY, row.MoldId);
            StatusText.Text = $"已定位孔位 #{row.Index}（{row.MoldCode}），角候选={row.IsCornerCandidate}，边缘孔={row.IsEdgeHole}，Top3={row.TopCandidates}";
        }

        private async void PositionGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is not System.Windows.Controls.DataGrid grid || grid.SelectedItem is not PositionRow row)
            {
                return;
            }
            var isStage1 = ReferenceEquals(grid, PositionGrid);
            var stage1Count = _stage1MoldFiles.Count;
            var stageMolds = isStage1
                ? _lastMolds.Take(stage1Count)
                : _lastMolds.Skip(stage1Count);
            var mold = stageMolds.FirstOrDefault(m => m.MoldId == row.MoldId);
            var hasOutline = mold?.OutlinePoints is { Count: >= 2 };

            _viewer.FocusHole(row.AbsX, row.AbsY, row.MoldId, targetZoom: 4.0);
            if (hasOutline && mold is not null)
            {
                await _viewer.BlinkMoldOutlineAsync(row.AbsX, row.AbsY, row.MoldId, mold.OutlinePoints, row.RotationDeg, row.IsMirrored);
            }
            else
            {
                await _viewer.BlinkFocusAsync(row.AbsX, row.AbsY, row.MoldId);
            }
            StatusText.Text = $"已放大定位孔位 #{row.Index}（{row.MoldCode}），角候选={row.IsCornerCandidate}，边缘孔={row.IsEdgeHole}，Top3={row.TopCandidates}";
        }

        private void RefreshMoldPreviewList(int stageId)
        {
            var moldRows = stageId == 1 ? _stage1MoldRows : _stage2MoldRows;
            moldRows.Clear();
            var files = stageId == 1 ? _stage1MoldFiles : _stage2MoldFiles;
            var prefix = stageId == 1 ? "M" : "N";
            var index = 1;
            foreach (var file in files)
            {
                _ = BuildMoldPreview(file);
                var moldName = System.IO.Path.GetFileNameWithoutExtension(file);
                if (TryGetDocumentFromCache(file, out var doc))
                {
                    var feature = DxfAnalyzer.ExtractMold(0, file, doc).Feature;
                    moldName = BuildMoldSizeText(feature);
                }

                moldRows.Add(new MoldRow
                {
                    MoldPreview = TryGetMoldPreviewFromCache(file, out var preview) ? preview : null,
                    MoldCode = $"{prefix}{index:D2}",
                    MoldName = moldName,
                    UsedCount = 0,
                    MatchType = stageId == 1 ? "台1" : "台2",
                    Remark = System.IO.Path.GetFileName(file)
                });
                index++;
            }
        }

        private void TouchLruKey(string key, LinkedList<string> lru, Dictionary<string, LinkedListNode<string>> lruNodes)
        {
            if (lruNodes.TryGetValue(key, out var existingNode))
            {
                lru.Remove(existingNode);
            }

            var newNode = lru.AddLast(key);
            lruNodes[key] = newNode;
        }

        private void TrimCacheToLimit<T>(
            Dictionary<string, T> cache,
            LinkedList<string> lru,
            Dictionary<string, LinkedListNode<string>> lruNodes,
            int maxEntries)
        {
            while (cache.Count > maxEntries && lru.First is not null)
            {
                var oldestKey = lru.First.Value;
                lru.RemoveFirst();
                lruNodes.Remove(oldestKey);
                cache.Remove(oldestKey);
            }
        }

        private bool TryGetDocumentFromCache(string path, out DxfDocument doc)
        {
            if (_documentCache.TryGetValue(path, out doc!))
            {
                TouchLruKey(path, _documentCacheLru, _documentCacheLruNodes);
                return true;
            }

            doc = null!;
            return false;
        }

        private void SetDocumentCache(string path, DxfDocument doc)
        {
            _documentCache[path] = doc;
            TouchLruKey(path, _documentCacheLru, _documentCacheLruNodes);
            TrimCacheToLimit(_documentCache, _documentCacheLru, _documentCacheLruNodes, CacheMaxEntries);
        }

        private bool TryGetMoldPreviewFromCache(string path, out ImageSource preview)
        {
            if (_moldPreviewCache.TryGetValue(path, out preview!))
            {
                TouchLruKey(path, _moldPreviewCacheLru, _moldPreviewCacheLruNodes);
                return true;
            }

            preview = null!;
            return false;
        }

        private void SetMoldPreviewCache(string path, ImageSource preview)
        {
            _moldPreviewCache[path] = preview;
            TouchLruKey(path, _moldPreviewCacheLru, _moldPreviewCacheLruNodes);
            TrimCacheToLimit(_moldPreviewCache, _moldPreviewCacheLru, _moldPreviewCacheLruNodes, CacheMaxEntries);
        }

        private ImageSource? BuildMoldPreview(string path)
        {
            if (TryGetMoldPreviewFromCache(path, out var cached))
            {
                return cached;
            }

            if (!TryGetDocumentFromCache(path, out var doc))
            {
                if (!System.IO.File.Exists(path))
                {
                    return null;
                }
                doc = LoadCadDocument(path);
                SetDocumentCache(path, doc);
            }

            const int width = 160;
            const int height = 100;
            var bounds = DxfAnalyzer.GetRawBounds(doc);
            var sceneW = Math.Max(bounds.Width, 1);
            var sceneH = Math.Max(bounds.Height, 1);
            var margin = 10.0;
            var scale = Math.Min((width - margin * 2) / sceneW, (height - margin * 2) / sceneH);
            scale = Math.Max(scale, 0.0001);

            var dv = new DrawingVisual();
            using (var dc = dv.RenderOpen())
            {
                dc.DrawRectangle(new SolidColorBrush(WpfColor.FromRgb(17, 17, 17)), null, new Rect(0, 0, width, height));

                var linePen = new System.Windows.Media.Pen(new SolidColorBrush(WpfColor.FromRgb(80, 210, 120)), 1);
                var circlePen = new System.Windows.Media.Pen(new SolidColorBrush(WpfColor.FromRgb(240, 200, 80)), 1);
                var polyPen = new System.Windows.Media.Pen(new SolidColorBrush(WpfColor.FromRgb(80, 180, 255)), 1);
                var arcPen = new System.Windows.Media.Pen(new SolidColorBrush(WpfColor.FromRgb(255, 140, 0)), 1);

                System.Windows.Point Map(double x, double y) => new(
                    (x - bounds.MinX) * scale + margin,
                    height - ((y - bounds.MinY) * scale + margin));

                foreach (var l in doc.Entities.Lines)
                {
                    dc.DrawLine(linePen, Map(l.StartPoint.X, l.StartPoint.Y), Map(l.EndPoint.X, l.EndPoint.Y));
                }

                foreach (var shape in CadPreviewGeometry.BuildPreviewShapes(doc, DxfAnalyzer.ExpandPolyline2D, DxfAnalyzer.SampleArc))
                {
                    var pen = shape.Kind switch
                    {
                        CadPreviewShapeKind.Line => linePen,
                        CadPreviewShapeKind.Circle => circlePen,
                        CadPreviewShapeKind.Arc => arcPen,
                        CadPreviewShapeKind.Polyline => polyPen,
                        _ => linePen
                    };

                    var geo = shape.ToGeometry(Map);
                    if (geo is not null)
                    {
                        dc.DrawGeometry(null, pen, geo);
                    }
                }
            }

            var bmp = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(dv);
            bmp.Freeze();
            SetMoldPreviewCache(path, bmp);
            return bmp;
        }
    }

    public interface IDxfPreviewPlugin
    {
        void CreatePreview(DxfDocument document, InteractiveDxfPreview viewer);
    }

    public sealed class BasicCanvasPreviewPlugin : IDxfPreviewPlugin
    {
        public void CreatePreview(DxfDocument document, InteractiveDxfPreview viewer)
        {
            var bounds = DxfAnalyzer.GetRawBounds(document);
            var width = Math.Max(bounds.Width, 1.0);
            var height = Math.Max(bounds.Height, 1.0);
            var viewWidth = 900.0;
            var viewHeight = 650.0;
            var scale = Math.Min(viewWidth / width, viewHeight / height) * 0.92;
            var margin = 20.0;

            var canvas = new Canvas
            {
                Width = viewWidth,
                Height = viewHeight,
                Background = new SolidColorBrush(WpfColor.FromRgb(17, 17, 17)),
                ClipToBounds = true
            };

            var unifiedStroke = new SolidColorBrush(WpfColor.FromRgb(144, 238, 144));

            foreach (var shape in CadPreviewGeometry.BuildPreviewShapes(document, DxfAnalyzer.ExpandPolyline2D, DxfAnalyzer.SampleArc))
            {
                var geometry = shape.ToGeometry((x, y) => new System.Windows.Point(
                    (x - bounds.MinX) * scale + margin,
                    viewHeight - ((y - bounds.MinY) * scale + margin)));
                if (geometry is null)
                {
                    continue;
                }

                var path = new Path
                {
                    Data = geometry,
                    Stroke = unifiedStroke,
                    StrokeThickness = 1,
                    Fill = WpfBrushes.Transparent
                };
                canvas.Children.Add(path);
            }

            viewer.LoadScene(canvas, bounds, viewWidth, viewHeight, scale, margin);
        }
    }

    public sealed class InteractiveDxfPreview : Border
    {
        private bool _compactMode = true;

        private static readonly WpfColor[] Palette =
        [
            WpfColor.FromRgb(255, 87, 34), WpfColor.FromRgb(76, 175, 80), WpfColor.FromRgb(33, 150, 243),
            WpfColor.FromRgb(255, 193, 7), WpfColor.FromRgb(156, 39, 176), WpfColor.FromRgb(0, 188, 212)
        ];

        private readonly Grid _root = new();
        private readonly Canvas _sceneCanvas = new();
        private readonly Canvas _markCanvas = new();
        private readonly Canvas _zoneCanvas = new();
        private readonly TransformGroup _group = new();
        private readonly ScaleTransform _scale = new(1, 1);
        private readonly TranslateTransform _translate = new(0, 0);

        private System.Windows.Point _dragStart;
        private bool _dragging;
        private RawBounds _bounds = new(0, 0, 1, 1);
        private double _viewHeight;
        private double _drawScale;
        private double _margin;
        private WpfEllipse? _focusRing;

        public InteractiveDxfPreview()
        {
            Background = new SolidColorBrush(WpfColor.FromRgb(17, 17, 17));
            BorderBrush = new SolidColorBrush(WpfColor.FromRgb(61, 61, 61));
            BorderThickness = new Thickness(1);
            ClipToBounds = true;

            _group.Children.Add(_scale);
            _group.Children.Add(_translate);
            _sceneCanvas.RenderTransform = _group;
            _zoneCanvas.RenderTransform = _group;
            _markCanvas.RenderTransform = _group;

            _root.Children.Add(_sceneCanvas);
            _root.Children.Add(_zoneCanvas);
            _root.Children.Add(_markCanvas);
            Child = _root;

            MouseWheel += OnMouseWheel;
            MouseLeftButtonDown += OnMouseLeftButtonDown;
            MouseLeftButtonUp += OnMouseLeftButtonUp;
            MouseDown += OnMouseDown;
            MouseUp += OnMouseUp;
            MouseMove += OnMouseMove;
        }

        public void LoadScene(Canvas scene, RawBounds bounds, double viewWidth, double viewHeight, double drawScale, double margin)
        {
            _bounds = bounds;
            _viewHeight = viewHeight;
            _drawScale = drawScale;
            _margin = margin;
            _sceneCanvas.Width = viewWidth;
            _sceneCanvas.Height = viewHeight;
            _zoneCanvas.Width = viewWidth;
            _zoneCanvas.Height = viewHeight;
            _markCanvas.Width = viewWidth;
            _markCanvas.Height = viewHeight;
            _sceneCanvas.Children.Clear();
            _zoneCanvas.Children.Clear();
            _markCanvas.Children.Clear();
            // Move children from the temporary canvas to avoid
            // "specified Visual is already a child of another Visual".
            while (scene.Children.Count > 0)
            {
                var child = scene.Children[0];
                scene.Children.RemoveAt(0);
                _sceneCanvas.Children.Add(child);
            }
            ResetView();
        }

        public void RenderPendingEdgeRecognitionHoles(IReadOnlyList<EdgeCandidate> edgeCandidates)
        {
            // 边缘孔识别结果已通过 RenderAnnotations 中的 IsEdgeHole 标注体现，此处不再绘制额外标记。
            _ = edgeCandidates;
        }

        public void RenderAnnotations(IReadOnlyList<HoleAssignment> assignments, IReadOnlyList<MoldProfile> molds, double? splitY)
        {
            _markCanvas.Children.Clear();
            _focusRing = null;
            var moldMap = molds
                .GroupBy(m => m.MoldId)
                .ToDictionary(g => g.Key, g => g.First());

            foreach (var ass in assignments)
            {
                var c = ass.Hole.Centroid;
                var p = ModelToCanvas(c.X, c.Y);
                var color = GetMoldColor(ass.MoldId);
                var brush = new SolidColorBrush(color);

                if (!_compactMode)
                {
                    // Draw a tiny cross at the reported centroid for coordinate verification.
                    var cross = new Polyline
                    {
                        Stroke = WpfBrushes.White,
                        StrokeThickness = 0.1,
                        Points = new PointCollection
                        {
                            new System.Windows.Point(p.X - 4, p.Y),
                            new System.Windows.Point(p.X + 4, p.Y),
                        }
                    };
                    _markCanvas.Children.Add(cross);
                    var cross2 = new Polyline
                    {
                        Stroke = WpfBrushes.White,
                        StrokeThickness = 0.1,
                        Points = new PointCollection
                        {
                            new System.Windows.Point(p.X, p.Y - 4),
                            new System.Windows.Point(p.X, p.Y + 4),
                        }
                    };
                    _markCanvas.Children.Add(cross2);
                }

                if (moldMap.TryGetValue(ass.MoldId, out var mold) && mold.OutlinePoints.Count >= 2)
                {
                    var rad = ass.RotationDeg * Math.PI / 180.0;
                    var outline = new PointCollection();
                    foreach (var pt in mold.OutlinePoints)
                    {
                        var x = ass.IsMirrored ? -pt.X : pt.X;
                        var y = pt.Y;
                        var xr = x * Math.Cos(rad) - y * Math.Sin(rad);
                        var yr = x * Math.Sin(rad) + y * Math.Cos(rad);
                        outline.Add(new System.Windows.Point(p.X + xr * _drawScale, p.Y - yr * _drawScale));
                    }
                    var poly = new Polyline
                    {
                        Points = outline,
                        Stroke = brush,
                        StrokeThickness = 0.7
                    };
                    _markCanvas.Children.Add(poly);
                }
                else
                {
                    var mark = new WpfEllipse
                    {
                        Width = 12,
                        Height = 12,
                        Fill = brush,
                        Stroke = WpfBrushes.White,
                        StrokeThickness = 0.7
                    };
                    Canvas.SetLeft(mark, p.X - 6);
                    Canvas.SetTop(mark, p.Y - 6);
                    _markCanvas.Children.Add(mark);
                }

                var shouldShowLabel = !_compactMode && ass.MoldId > 0 && !ass.Hole.HoleType.StartsWith("Contour", StringComparison.Ordinal) && !ass.Hole.HoleType.StartsWith("EdgeNotch:", StringComparison.Ordinal);

                if (shouldShowLabel)
                {
                    var prefix = splitY.HasValue && ass.Hole.Centroid.Y >= splitY.Value ? "N" : "M";
                    var text = new TextBlock
                    {
                        Text = $"{prefix}{ass.MoldId:D2}",
                        Foreground = brush,
                        FontWeight = FontWeights.SemiBold,
                        FontSize = 7
                    };
                    Canvas.SetLeft(text, p.X + 8);
                    Canvas.SetTop(text, p.Y - 8);
                    _markCanvas.Children.Add(text);
                }
            }
        }

        public void RenderCornerContours(
            RectBounds? rect,
            IReadOnlyList<(double X, double Y)>? outerContourPoints,
            IReadOnlyList<CornerStepPath>? cornerPaths,
            IReadOnlyList<HoleFeature>? cornerHints,
            double boardWidth,
            double? splitY = null)
        {
            _zoneCanvas.Children.Clear();
            if (rect is null)
            {
                return;
            }

            // 1) 画最小外包矩形（黄色）
            var r1 = ModelToCanvas(rect.MinX, rect.MinY);
            var r2 = ModelToCanvas(rect.MaxX, rect.MaxY);
            var rectBox = new WpfRectangle
            {
                Width = Math.Abs(r2.X - r1.X),
                Height = Math.Abs(r2.Y - r1.Y),
                Stroke = new SolidColorBrush(WpfColor.FromArgb(240, 255, 235, 59)),
                StrokeThickness = 0.5,
                StrokeDashArray = new DoubleCollection([4, 3]),
                Fill = WpfBrushes.Transparent
            };
            Canvas.SetLeft(rectBox, Math.Min(r1.X, r2.X));
            Canvas.SetTop(rectBox, Math.Min(r1.Y, r2.Y));
            _zoneCanvas.Children.Add(rectBox);

            if (boardWidth > 0 && splitY.HasValue)
            {
                var splitLeft = ModelToCanvas(rect.MinX, splitY.Value);
                var splitRight = ModelToCanvas(rect.MaxX, splitY.Value);
                var threshold = new WpfLine
                {
                    X1 = splitLeft.X,
                    Y1 = splitLeft.Y,
                    X2 = splitRight.X,
                    Y2 = splitRight.Y,
                    Stroke = new SolidColorBrush(WpfColor.FromArgb(255, 255, 0, 0)),
                    StrokeThickness = 1.2,
                    StrokeDashArray = new DoubleCollection([8, 4, 2, 4])
                };
                _zoneCanvas.Children.Add(threshold);
            }

            // 2) 画当前识别到的真实外轮廓（红色）
            if (outerContourPoints is not null && outerContourPoints.Count >= 2)
            {
                var polyAll = new Polyline
                {
                    Stroke = new SolidColorBrush(WpfColor.FromArgb(245, 255, 82, 82)),
                    StrokeThickness = 0.1,
                    StrokeLineJoin = PenLineJoin.Round,
                    StrokeStartLineCap = PenLineCap.Round,
                    StrokeEndLineCap = PenLineCap.Round
                };
                foreach (var p in outerContourPoints)
                {
                    polyAll.Points.Add(ModelToCanvas(p.X, p.Y));
                }
                _zoneCanvas.Children.Add(polyAll);
            }

            // 3) 青色线 = 红色外轮廓 - 矩形外轮廓（按“线段差集”绘制，避免跨段误连）
            if (outerContourPoints is null || outerContourPoints.Count < 2)
            {
                return;
            }

            var edgeTol = Compat.Clamp(Math.Min(rect.Width, rect.Height) * 0.0004, 0.05, 0.1);
            bool IsSegmentOnRectEdge((double X, double Y) a, (double X, double Y) b)
            {
                var horizontal = Math.Abs(a.Y - b.Y) <= edgeTol;
                var vertical = Math.Abs(a.X - b.X) <= edgeTol;

                if (vertical)
                {
                    var onLeft = Math.Abs(a.X - rect.MinX) <= edgeTol && Math.Abs(b.X - rect.MinX) <= edgeTol;
                    var onRight = Math.Abs(a.X - rect.MaxX) <= edgeTol && Math.Abs(b.X - rect.MaxX) <= edgeTol;
                    return onLeft || onRight;
                }

                if (horizontal)
                {
                    var onBottom = Math.Abs(a.Y - rect.MinY) <= edgeTol && Math.Abs(b.Y - rect.MinY) <= edgeTol;
                    var onTop = Math.Abs(a.Y - rect.MaxY) <= edgeTol && Math.Abs(b.Y - rect.MaxY) <= edgeTol;
                    return onBottom || onTop;
                }

                return false;
            }

            var contour = outerContourPoints.ToList();
            if (contour.Count > 2)
            {
                var first = contour[0];
                var last = contour[^1];
                if (Math.Sqrt((first.X - last.X) * (first.X - last.X) + (first.Y - last.Y) * (first.Y - last.Y)) < edgeTol)
                {
                    contour.RemoveAt(contour.Count - 1);
                }
            }

            // 去除相邻重复点/极短边，避免差集切段时丢失短台阶。
            var normalized = new List<(double X, double Y)>();
            foreach (var p in contour)
            {
                if (normalized.Count == 0)
                {
                    normalized.Add(p);
                    continue;
                }

                var prev = normalized[^1];
                var d = Math.Sqrt((p.X - prev.X) * (p.X - prev.X) + (p.Y - prev.Y) * (p.Y - prev.Y));
                if (d > Math.Max(edgeTol * 0.2, 1e-6))
                {
                    normalized.Add(p);
                }
            }
            contour = normalized;

            var runs = new List<List<(double X, double Y)>>();
            var removedSegments = new List<((double X, double Y) A, (double X, double Y) B)>();
            var current = new List<(double X, double Y)>();

            // 按“线段”做差集：只要线段中点不在矩形边上，就保留该段。
            // 注意要包含首尾闭合段，避免漏掉轮廓起点附近的一段。
            var closeGap = Math.Sqrt(
                (contour[0].X - contour[^1].X) * (contour[0].X - contour[^1].X) +
                (contour[0].Y - contour[^1].Y) * (contour[0].Y - contour[^1].Y));
            var isClosed = closeGap <= edgeTol * 1.5;
            var segCount = isClosed ? contour.Count : contour.Count - 1;

            for (var i = 0; i < segCount; i++)
            {
                var a = contour[i];
                var b = contour[(i + 1) % contour.Count];
                var keepSeg = !IsSegmentOnRectEdge(a, b);

                // 边保护：若线段明显位于矩形内部（离任何外边有安全间距），强制保留。
                var minEdgeDistA = Math.Min(
                    Math.Min(Math.Abs(a.X - rect.MinX), Math.Abs(rect.MaxX - a.X)),
                    Math.Min(Math.Abs(a.Y - rect.MinY), Math.Abs(rect.MaxY - a.Y)));
                var minEdgeDistB = Math.Min(
                    Math.Min(Math.Abs(b.X - rect.MinX), Math.Abs(rect.MaxX - b.X)),
                    Math.Min(Math.Abs(b.Y - rect.MinY), Math.Abs(rect.MaxY - b.Y)));
                var innerSafe = Math.Max(edgeTol * 1.6, 0.1);
                if (minEdgeDistA > innerSafe && minEdgeDistB > innerSafe)
                {
                    keepSeg = true;
                }

                if (!keepSeg)
                {
                    removedSegments.Add((a, b));
                    if (current.Count >= 2)
                    {
                        runs.Add(current);
                    }
                    current = new List<(double X, double Y)>();
                    continue;
                }

                if (current.Count == 0)
                {
                    current.Add(a);
                    current.Add(b);
                }
                else
                {
                    var last = current[^1];
                    if (Math.Abs(last.X - a.X) <= 1e-6 && Math.Abs(last.Y - a.Y) <= 1e-6)
                    {
                        current.Add(b);
                    }
                    else
                    {
                        if (current.Count >= 2)
                        {
                            runs.Add(current);
                        }
                        current = new List<(double X, double Y)> { a, b };
                    }
                }
            }

            if (current.Count >= 2)
            {
                runs.Add(current);
            }

            if (runs.Count == 0)
            {
                return;
            }

            foreach (var run in runs)
            {
                var cyanPoly = new Polyline
                {
                    Stroke = new SolidColorBrush(WpfColor.FromArgb(245, 0, 255, 255)),
                    StrokeThickness = 0.1,
                    StrokeLineJoin = PenLineJoin.Round,
                    StrokeStartLineCap = PenLineCap.Round,
                    StrokeEndLineCap = PenLineCap.Round
                };

                foreach (var p in run)
                {
                    cyanPoly.Points.Add(ModelToCanvas(p.X, p.Y));
                }
                _zoneCanvas.Children.Add(cyanPoly);
            }

            if (!_compactMode)
            {
                // 辅助线（紫色）：M01 连续冲压使用的外偏移路径。
                if (cornerPaths is not null)
                {
                    foreach (var gp in cornerPaths)
                    {
                        if (gp.Points is null || gp.Points.Count < 2)
                        {
                            continue;
                        }

                        // 紫色线：显示完整 offset 路径（不做端点/拐点压缩），便于核对几何本身。
                        var guide = new Polyline
                        {
                            Stroke = new SolidColorBrush(WpfColor.FromArgb(235, 186, 104, 200)),
                            StrokeThickness = 0.7,
                            StrokeLineJoin = PenLineJoin.Round,
                            StrokeStartLineCap = PenLineCap.Round,
                            StrokeEndLineCap = PenLineCap.Round,
                            StrokeDashArray = new DoubleCollection([5, 3])
                        };
                        foreach (var p in gp.Points)
                        {
                            guide.Points.Add(ModelToCanvas(p.X, p.Y));
                        }
                        _zoneCanvas.Children.Add(guide);

                    }
                }

                if (!_compactMode)
                {
                    // 调试层：被判定为“矩形边而删除”的线段（橙色）
                    foreach (var seg in removedSegments)
                    {
                        var p1 = ModelToCanvas(seg.A.X, seg.A.Y);
                        var p2 = ModelToCanvas(seg.B.X, seg.B.Y);
                        var dbg = new WpfLine
                        {
                            X1 = p1.X,
                            Y1 = p1.Y,
                            X2 = p2.X,
                            Y2 = p2.Y,
                            Stroke = new SolidColorBrush(WpfColor.FromArgb(235, 255, 152, 0)),
                            StrokeThickness = 0.4,
                            StrokeDashArray = new DoubleCollection([2, 2])
                        };
                        _zoneCanvas.Children.Add(dbg);
                    }

                    var firstRun = runs[0];
                    var labelAnchor = ModelToCanvas(firstRun[0].X, firstRun[0].Y);
                    var label = new TextBlock
                    {
                        Text = "待冲轮廓",
                        Foreground = new SolidColorBrush(WpfColor.FromArgb(245, 0, 255, 255)),
                        FontSize = 8,
                        FontWeight = FontWeights.Bold,
                        Background = new SolidColorBrush(WpfColor.FromArgb(120, 0, 0, 0))
                    };
                    Canvas.SetLeft(label, labelAnchor.X + 6);
                    Canvas.SetTop(label, labelAnchor.Y - 18);
                    _zoneCanvas.Children.Add(label);
                }
            }
        }

        public static WpfColor GetMoldColor(int moldId)
        {
            if (moldId <= 0)
            {
                return Palette[0];
            }
            return Palette[(moldId - 1) % Palette.Length];
        }

        public void SetCompactMode(bool compact)
        {
            _compactMode = compact;
        }

        public void FocusHole(double modelX, double modelY, int moldId, double? targetZoom = null, bool showFocusRing = false)
        {
            var p = ModelToCanvas(modelX, modelY);
            var color = new SolidColorBrush(GetMoldColor(moldId));
            if (!showFocusRing)
            {
                if (_focusRing is not null)
                {
                    _markCanvas.Children.Remove(_focusRing);
                    _focusRing = null;
                }
            }
            else
            {
                if (_focusRing is null)
                {
                    _focusRing = new WpfEllipse
                    {
                        Width = 26,
                        Height = 26,
                        Stroke = color,
                        StrokeThickness = 2.5,
                        Fill = WpfBrushes.Transparent
                    };
                    _markCanvas.Children.Add(_focusRing);
                }
                _focusRing.Stroke = color;
                Canvas.SetLeft(_focusRing, p.X - 13);
                Canvas.SetTop(_focusRing, p.Y - 13);
            }

            if (targetZoom.HasValue)
            {
                var z = Compat.Clamp(targetZoom.Value, 0.2, 30);
                _scale.ScaleX = z;
                _scale.ScaleY = z;
            }
            var scale = _scale.ScaleX;
            var centerX = ActualWidth > 0 ? ActualWidth / 2 : _sceneCanvas.Width / 2;
            var centerY = ActualHeight > 0 ? ActualHeight / 2 : _sceneCanvas.Height / 2;
            _translate.X = centerX - p.X * scale;
            _translate.Y = centerY - p.Y * scale;
        }

        public async Task BlinkMoldOutlineAsync(double modelX, double modelY, int moldId, IReadOnlyList<(double X, double Y)> outlinePoints, double rotationDeg, bool isMirrored)
        {
            if (outlinePoints.Count < 2)
            {
                await BlinkFocusAsync(modelX, modelY, moldId);
                return;
            }

            var p = ModelToCanvas(modelX, modelY);
            var rad = rotationDeg * Math.PI / 180.0;
            var outline = new PointCollection();
            foreach (var pt in outlinePoints)
            {
                var x = isMirrored ? -pt.X : pt.X;
                var y = pt.Y;
                var xr = x * Math.Cos(rad) - y * Math.Sin(rad);
                var yr = x * Math.Sin(rad) + y * Math.Cos(rad);
                outline.Add(new System.Windows.Point(p.X + xr * _drawScale, p.Y - yr * _drawScale));
            }

            var strokeColor = GetMoldColor(moldId);
            // 使用局部引用：连续双击或并发 await 时，不应依赖会被其它调用清掉的共享字段。
            var blinkPoly = new Polyline
            {
                Points = outline,
                Stroke = new SolidColorBrush(strokeColor),
                StrokeThickness = 2.8,
                Fill = WpfBrushes.Transparent
            };
            _markCanvas.Children.Add(blinkPoly);

            for (var i = 0; i < 3; i++)
            {
                blinkPoly.Visibility = Visibility.Hidden;
                await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Background);
                await Task.Delay(120);
                blinkPoly.Visibility = Visibility.Visible;
                await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Background);
                await Task.Delay(120);
            }

            if (_markCanvas.Children.Contains(blinkPoly))
            {
                _markCanvas.Children.Remove(blinkPoly);
            }
        }

        public async Task BlinkFocusAsync(double modelX, double modelY, int moldId)
        {
            var p = ModelToCanvas(modelX, modelY);
            var brush = new SolidColorBrush(GetMoldColor(moldId));
            const double arm = 12.0;
            var crossH = new Polyline
            {
                Stroke = brush,
                StrokeThickness = 2.2,
                Points = new PointCollection
                {
                    new System.Windows.Point(p.X - arm, p.Y),
                    new System.Windows.Point(p.X + arm, p.Y)
                }
            };
            var crossV = new Polyline
            {
                Stroke = brush,
                StrokeThickness = 2.2,
                Points = new PointCollection
                {
                    new System.Windows.Point(p.X, p.Y - arm),
                    new System.Windows.Point(p.X, p.Y + arm)
                }
            };
            _markCanvas.Children.Add(crossH);
            _markCanvas.Children.Add(crossV);

            for (var i = 0; i < 3; i++)
            {
                crossH.Visibility = Visibility.Hidden;
                crossV.Visibility = Visibility.Hidden;
                await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Background);
                await Task.Delay(120);
                crossH.Visibility = Visibility.Visible;
                crossV.Visibility = Visibility.Visible;
                await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Background);
                await Task.Delay(120);
            }

            if (_markCanvas.Children.Contains(crossH))
            {
                _markCanvas.Children.Remove(crossH);
            }
            if (_markCanvas.Children.Contains(crossV))
            {
                _markCanvas.Children.Remove(crossV);
            }
        }

        private System.Windows.Point ModelToCanvas(double x, double y)
        {
            return new System.Windows.Point(
                (x - _bounds.MinX) * _drawScale + _margin,
                _viewHeight - ((y - _bounds.MinY) * _drawScale + _margin));
        }

        private void ResetView()
        {
            _scale.ScaleX = 1;
            _scale.ScaleY = 1;
            _translate.X = 0;
            _translate.Y = 0;
        }

        private void OnMouseWheel(object sender, MouseWheelEventArgs e)
        {
            var factor = e.Delta > 0 ? 1.12 : 0.89;
            var old = _scale.ScaleX;
            var target = Compat.Clamp(old * factor, 0.2, 30);
            factor = target / old;
            var center = e.GetPosition(this);
            _translate.X = center.X - factor * (center.X - _translate.X);
            _translate.Y = center.Y - factor * (center.Y - _translate.Y);
            _scale.ScaleX = target;
            _scale.ScaleY = target;
        }

        private void OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                ResetView();
                return;
            }
        }

        private void OnMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (!_dragging)
            {
                return;
            }
            var now = e.GetPosition(this);
            var delta = now - _dragStart;
            _translate.X += delta.X;
            _translate.Y += delta.Y;
            _dragStart = now;
        }

        private void OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_dragging && e.ChangedButton == MouseButton.Left)
            {
                _dragging = false;
                Cursor = System.Windows.Input.Cursors.Arrow;
                ReleaseMouseCapture();
            }
        }

        private void OnMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Middle)
            {
                return;
            }
            _dragging = true;
            _dragStart = e.GetPosition(this);
            Cursor = System.Windows.Input.Cursors.ScrollAll;
            CaptureMouse();
            e.Handled = true;
        }

        private void OnMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Middle)
            {
                return;
            }
            _dragging = false;
            Cursor = System.Windows.Input.Cursors.Arrow;
            ReleaseMouseCapture();
            e.Handled = true;
        }
    }

    public enum CadPreviewShapeKind
    {
        Line,
        Circle,
        Arc,
        Polyline
    }

    public sealed class CadPreviewShape
    {
        public CadPreviewShapeKind Kind { get; init; }
        public IReadOnlyList<(double X, double Y)> Points { get; init; } = [];
        public bool IsClosed { get; init; }

        public Geometry? ToGeometry(Func<double, double, System.Windows.Point> map)
        {
            if (Points.Count < 2)
            {
                return null;
            }

            var geo = new StreamGeometry();
            using (var g = geo.Open())
            {
                g.BeginFigure(map(Points[0].X, Points[0].Y), false, IsClosed);
                g.PolyLineTo(Points.Skip(1).Select(p => map(p.X, p.Y)).ToList(), true, false);
            }
            geo.Freeze();
            return geo;
        }
    }

    public static class CadPreviewGeometry
    {
        public static IEnumerable<CadPreviewShape> BuildPreviewShapes(
            DxfDocument doc,
            Func<Polyline2D, int, IReadOnlyList<(double X, double Y)>> expandPolyline,
            Func<Arc, int, IReadOnlyList<(double X, double Y)>> sampleArc)
        {
            foreach (var line in doc.Entities.Lines)
            {
                yield return new CadPreviewShape
                {
                    Kind = CadPreviewShapeKind.Line,
                    Points = [(line.StartPoint.X, line.StartPoint.Y), (line.EndPoint.X, line.EndPoint.Y)]
                };
            }

            foreach (var circle in doc.Entities.Circles)
            {
                yield return new CadPreviewShape
                {
                    Kind = CadPreviewShapeKind.Circle,
                    Points = sampleArc(new Arc(circle.Center, circle.Radius, 0, 360), 72),
                    IsClosed = true
                };
            }

            foreach (var poly in doc.Entities.Polylines2D)
            {
                var pts = expandPolyline(poly, 24);
                if (pts.Count < 2)
                {
                    continue;
                }

                yield return new CadPreviewShape
                {
                    Kind = CadPreviewShapeKind.Polyline,
                    Points = pts,
                    IsClosed = poly.IsClosed
                };
            }

            foreach (var arc in doc.Entities.Arcs)
            {
                var pts = sampleArc(arc, 32);
                if (pts.Count < 2)
                {
                    continue;
                }

                yield return new CadPreviewShape
                {
                    Kind = CadPreviewShapeKind.Arc,
                    Points = pts
                };
            }
        }
    }

    public static class DxfAnalyzer
    {
        private const int SignatureSamples = 72;

        public static MoldProfile ExtractMold(int moldId, string path, DxfDocument? preloadedDoc = null)
        {
            var doc = preloadedDoc ?? CadDocumentLoader.Load(path);
            // 模具特征只取“可闭合几何”，避免 OpenPolyline 的伪面积干扰面积比匹配。
            var holes = ExtractHoles(doc, includeOpenPolylines: false);
            var outer = DetectOuterRectangle(doc);
            var holesInfo = string.Join(", ", holes.Select(h => $"W={h.Width:F2},H={h.Height:F2},A={h.Area:F2}"));
            AppLogger.Instance.Info($"[模具提取] M{moldId:D2} 提取完成: 共{holes.Count}个孔, 列表:[{holesInfo}], outer.Area={outer.Area:F2}");
            var candidates = holes
                .Where(h => h.Area <= Math.Max(outer.Area * 0.75, 10.0))
                .OrderByDescending(h => h.Area)
                .ToList();

            // Fallback: 如果 outer.Area 太小（DXF 没有外部矩形）导致 candidates 被全部过滤掉，
            // 直接用所有 holes 的最大者作为模具特征，避免误用 BuildFeatureFromEntities 兜底造成面积失真。
            if (candidates.Count == 0 && holes.Count > 0)
            {
                candidates = [holes.OrderByDescending(h => h.Area).First()];
                AppLogger.Instance.Warn($"[模具提取] M{moldId:D2} candidates 为空, fallback 使用最大hole: W={candidates[0].Width:F2},H={candidates[0].Height:F2},A={candidates[0].Area:F2}");
            }

            var feature = candidates.FirstOrDefault();
            if (feature is null)
            {
                feature = BuildFeatureFromEntities(doc);
                if (feature is not null)
                {
                    candidates.Add(feature);
                }
            }

            feature ??= new HoleFeature("Unknown", (0, 0), 1, 1, 1, 1, 0, CreateCircleSignature(1, SignatureSamples));
            if (candidates.Count == 0)
            {
                candidates.Add(feature);
            }

            // M01 连续冲压路径依赖“模具本体中心”来保证边缘贴合；
            // 其余模具采用特征中心以保证单孔定位不偏移。
            var useBodyCenter = moldId == 1;
            var outline = ExtractMoldOutline(doc, feature.Centroid, useBodyCenter);
            return new MoldProfile(moldId, path, feature, outline, candidates);
        }

        public static ProjectProfile ExtractProject(DxfDocument doc)
        {
            var holes = ExtractHoles(doc, includeOpenPolylines: false);
            var outer = DetectOuterRectangle(doc);
            var edgeCandidates = ExtractEdgePartialCandidates(doc, outer);
            var cornerCandidates = ExtractCornerMissingFeatures(doc, outer);
            var cornerStepPaths = ExtractCornerStepPaths(doc, outer);
            var contourPaths = ExtractContourDifferencePaths(doc, outer);
            var splitY = outer.MinY + outer.Height * 0.5;
            var stage1ContourPaths = new List<CornerStepPath>();
            var stage2ContourPaths = new List<CornerStepPath>();
            foreach (var path in contourPaths)
            {
                if (path.Points is null || path.Points.Count == 0)
                {
                    continue;
                }

                var avgY = path.Points.Average(p => p.Y);
                var minY = path.Points.Min(p => p.Y);
                var maxY = path.Points.Max(p => p.Y);
                var crosses = minY < splitY && maxY >= splitY;
                if (!crosses)
                {
                    (avgY < splitY ? stage1ContourPaths : stage2ContourPaths).Add(path);
                    continue;
                }

                var stage1Count = path.Points.Count(p => p.Y < splitY);
                var stage2Count = path.Points.Count - stage1Count;
                if (stage1Count >= stage2Count)
                {
                    stage1ContourPaths.Add(path);
                }
                else
                {
                    stage2ContourPaths.Add(path);
                }
            }

            var maxHoleArea = Math.Max(outer.Area * 0.2, 1.0);
            var innerHoles = holes
                .Where(h =>
                {
                    var margin = Math.Max(Math.Max(h.Width, h.Height) * 0.5, 1.0);
                    var intersectsOuter = h.Centroid.X >= outer.MinX - margin &&
                                          h.Centroid.X <= outer.MaxX + margin &&
                                          h.Centroid.Y >= outer.MinY - margin &&
                                          h.Centroid.Y <= outer.MaxY + margin;
                    return intersectsOuter && h.Area <= maxHoleArea * 1.5;
                })
                .ToList();

            // Fallback: if strict filtering removes all holes, keep geometric holes
            // under area threshold so matching can still run.
            if (innerHoles.Count == 0)
            {
                innerHoles = holes
                    .Where(h => h.Area <= maxHoleArea)
                    .ToList();
            }

            // 输出每个内部孔相对外框左下角的坐标（绝对值与相对值同时输出，便于核对）。
            var relHolesInfo = string.Join(" | ", innerHoles.Select(h =>
                $"[W={h.Width:F2},H={h.Height:F2},A={h.Area:F2}] Abs=({h.Centroid.X:F2},{h.Centroid.Y:F2}) -> Rel=({h.Centroid.X - outer.MinX:F2},{h.Centroid.Y - outer.MinY:F2})"));
            AppLogger.Instance.Info($"[项目提取] 共{innerHoles.Count}个内部孔, 外框: MinX={outer.MinX:F2}, MinY={outer.MinY:F2}, W={outer.Width:F2}, H={outer.Height:F2}");
            AppLogger.Instance.Info($"[项目提取] 孔坐标(相对外框左下角): {relHolesInfo}");

            return new ProjectProfile(
                outer,
                DeduplicateHoles(innerHoles),
                cornerCandidates,
                edgeCandidates,
                cornerStepPaths,
                contourPaths,
                stage1ContourPaths,
                stage2ContourPaths);
        }

        private static IReadOnlyList<CornerStepPath> ExtractCornerStepPaths(DxfDocument doc, RectBounds outer)
        {
            // 旧接口保留：连续冲压改走 ContourPaths，不再按四角拆分。
            return [];
        }

        public static IReadOnlyList<(double X, double Y)> ExtractOuterContourForDebug(DxfDocument doc)
        {
            var outer = DetectOuterRectangle(doc);
            var contour = SelectOuterContourPoints(doc, outer).ToList();
            if (contour.Count >= 2)
            {
                return contour;
            }

            // Debug fallback: pick points near outer rectangle edges so red contour is always visible.
            var tol = Math.Max(Math.Min(outer.Width, outer.Height) * 0.01, 2.0);
            var edgePts = CollectGeometryPoints(doc)
                .Where(p =>
                    Math.Abs(p.X - outer.MinX) <= tol ||
                    Math.Abs(p.X - outer.MaxX) <= tol ||
                    Math.Abs(p.Y - outer.MinY) <= tol ||
                    Math.Abs(p.Y - outer.MaxY) <= tol)
                .DistinctBy(p => ($"{Math.Round(p.X, 2)}|{Math.Round(p.Y, 2)}"))
                .ToList();

            if (edgePts.Count < 2)
            {
                return [];
            }

            var cx = (outer.MinX + outer.MaxX) * 0.5;
            var cy = (outer.MinY + outer.MaxY) * 0.5;
            return edgePts
                .OrderBy(p => Math.Atan2(p.Y - cy, p.X - cx))
                .ThenBy(p => Math.Sqrt((p.X - cx) * (p.X - cx) + (p.Y - cy) * (p.Y - cy)))
                .ToList();
        }

        private static IReadOnlyList<CornerStepPath> ExtractContourDifferencePaths(DxfDocument doc, RectBounds rect)
        {
            var contour = SelectOuterContourPoints(doc, rect).ToList();
            if (contour.Count < 2)
            {
                return [];
            }

            var edgeTol = Compat.Clamp(Math.Min(rect.Width, rect.Height) * 0.0004, 0.05, 0.1);
            bool IsSegmentOnRectEdge((double X, double Y) a, (double X, double Y) b)
            {
                var horizontal = Math.Abs(a.Y - b.Y) <= edgeTol;
                var vertical = Math.Abs(a.X - b.X) <= edgeTol;

                if (vertical)
                {
                    var onLeft = Math.Abs(a.X - rect.MinX) <= edgeTol && Math.Abs(b.X - rect.MinX) <= edgeTol;
                    var onRight = Math.Abs(a.X - rect.MaxX) <= edgeTol && Math.Abs(b.X - rect.MaxX) <= edgeTol;
                    return onLeft || onRight;
                }

                if (horizontal)
                {
                    var onBottom = Math.Abs(a.Y - rect.MinY) <= edgeTol && Math.Abs(b.Y - rect.MinY) <= edgeTol;
                    var onTop = Math.Abs(a.Y - rect.MaxY) <= edgeTol && Math.Abs(b.Y - rect.MaxY) <= edgeTol;
                    return onBottom || onTop;
                }

                return false;
            }

            var closeGap = Math.Sqrt(
                (contour[0].X - contour[^1].X) * (contour[0].X - contour[^1].X) +
                (contour[0].Y - contour[^1].Y) * (contour[0].Y - contour[^1].Y));
            var isClosed = closeGap <= edgeTol * 1.5;
            var segCount = isClosed ? contour.Count : contour.Count - 1;

            var runs = new List<List<(double X, double Y)>>();
            var current = new List<(double X, double Y)>();

            for (var i = 0; i < segCount; i++)
            {
                var a = contour[i];
                var b = contour[(i + 1) % contour.Count];
                var keepSeg = !IsSegmentOnRectEdge(a, b);

                if (!keepSeg)
                {
                    if (current.Count >= 2)
                    {
                        runs.Add(current);
                    }
                    current = new List<(double X, double Y)>();
                    continue;
                }

                if (current.Count == 0)
                {
                    current.Add(a);
                    current.Add(b);
                }
                else
                {
                    var last = current[^1];
                    if (Math.Abs(last.X - a.X) <= 1e-6 && Math.Abs(last.Y - a.Y) <= 1e-6)
                    {
                        current.Add(b);
                    }
                    else
                    {
                        if (current.Count >= 2)
                        {
                            runs.Add(current);
                        }
                        current = new List<(double X, double Y)> { a, b };
                    }
                }
            }

            if (current.Count >= 2)
            {
                runs.Add(current);
            }

            return runs
                .Select((r, idx) => new CornerStepPath($"Contour{idx + 1}", r))
                .ToList();
        }

        private static List<(double X, double Y)> SelectOuterContourPoints(DxfDocument doc, RectBounds outer)
        {
            var loop = BuildOuterContourByStitching(doc, outer);
            return loop ?? [];
        }

        private static List<(double X, double Y)>? BuildOuterContourByStitching(DxfDocument doc, RectBounds outer)
        {
            // 仅基于 Line + Polyline2D 重建外轮廓（不使用 Arc）。
            var chains = new List<List<(double X, double Y)>>();

            foreach (var l in doc.Entities.Lines)
            {
                chains.Add([(l.StartPoint.X, l.StartPoint.Y), (l.EndPoint.X, l.EndPoint.Y)]);
            }

            foreach (var pl in doc.Entities.Polylines2D.Where(p => p.Vertexes.Count >= 2))
            {
                var pts = ExpandPolyline2D(pl, 24).ToList();
                if (pts.Count >= 2)
                {
                    chains.Add(pts);
                }
            }

            if (chains.Count == 0)
            {
                return null;
            }

            var diag = Math.Sqrt(outer.Width * outer.Width + outer.Height * outer.Height);
            var tol = Compat.Clamp(diag * 0.0005, 1e-4, 0.2);

            (double X, double Y) Snap((double X, double Y) p)
            {
                var nx = Math.Round(p.X / tol) * tol;
                var ny = Math.Round(p.Y / tol) * tol;
                return (nx, ny);
            }

            string EdgeKey((double X, double Y) a, (double X, double Y) b)
            {
                var k1 = $"{a.X:F4},{a.Y:F4}";
                var k2 = $"{b.X:F4},{b.Y:F4}";
                return string.CompareOrdinal(k1, k2) <= 0 ? $"{k1}|{k2}" : $"{k2}|{k1}";
            }

            var adj = new Dictionary<(double X, double Y), HashSet<(double X, double Y)>>();
            var edgeSet = new HashSet<string>();

            void AddEdge((double X, double Y) a, (double X, double Y) b)
            {
                var dx = Math.Abs(a.X - b.X);
                var dy = Math.Abs(a.Y - b.Y);
                if (dx <= 1e-9 && dy <= 1e-9)
                {
                    return;
                }

                // 只接受水平/竖直边，避免把错误候选中的斜边并入外轮廓。
                if (dx > tol && dy > tol)
                {
                    return;
                }

                var key = EdgeKey(a, b);
                if (!edgeSet.Add(key))
                {
                    return;
                }

                if (!adj.TryGetValue(a, out var la))
                {
                    la = [];
                    adj[a] = la;
                }
                la.Add(b);

                if (!adj.TryGetValue(b, out var lb))
                {
                    lb = [];
                    adj[b] = lb;
                }
                lb.Add(a);
            }

            var axisTol = Math.Max(tol * 0.6, 1e-5);

            foreach (var c in chains)
            {
                for (var i = 1; i < c.Count; i++)
                {
                    var a = Snap(c[i - 1]);
                    var b = Snap(c[i]);
                    if (Math.Abs(a.X - b.X) <= 1e-9 && Math.Abs(a.Y - b.Y) <= 1e-9)
                    {
                        continue;
                    }

                    // 外轮廓只使用正交边（水平/垂直）。
                    var dx = Math.Abs(a.X - b.X);
                    var dy = Math.Abs(a.Y - b.Y);
                    var orthogonal = dx <= axisTol || dy <= axisTol;
                    if (!orthogonal)
                    {
                        continue;
                    }

                    AddEdge(a, b);
                }

                var first = Snap(c[0]);
                var last = Snap(c[^1]);
                var closedLike = Math.Abs(first.X - last.X) <= tol && Math.Abs(first.Y - last.Y) <= tol;
                if (closedLike)
                {
                    var dx = Math.Abs(first.X - last.X);
                    var dy = Math.Abs(first.Y - last.Y);
                    var orthogonal = dx <= axisTol || dy <= axisTol;
                    if (orthogonal)
                    {
                        AddEdge(last, first);
                    }
                }
            }

            var nodes = adj.Keys.ToList();
            if (nodes.Count < 4)
            {
                return null;
            }

            var adjList = adj.ToDictionary(kv => kv.Key, kv => kv.Value.ToList());
            var comps = GetConnectedComponents(nodes, adjList);

            List<(double X, double Y)>? bestOrthogonalPath = null;
            var bestOrthogonalArea = -1.0;
            List<(double X, double Y)>? bestAnyPath = null;
            var bestAnyArea = -1.0;

            foreach (var comp in comps)
            {
                if (comp.Count < 4)
                {
                    continue;
                }

                // 闭环候选：分量内每个节点都必须是度2
                var degree2 = comp.All(n => adj.TryGetValue(n, out var nbs) && nbs.Count == 2);
                if (!degree2)
                {
                    continue;
                }

                var start = comp.First();
                var path = new List<(double X, double Y)>();
                (double X, double Y)? prev = null;
                var curr = start;

                for (var step = 0; step < comp.Count + 2; step++)
                {
                    path.Add(curr);
                    var nbs = adj[curr].Where(comp.Contains).ToList();
                    if (nbs.Count != 2)
                    {
                        path.Clear();
                        break;
                    }

                    var next = prev is null
                        ? nbs[0]
                        : (Math.Abs(nbs[0].X - prev.Value.X) <= 1e-9 && Math.Abs(nbs[0].Y - prev.Value.Y) <= 1e-9 ? nbs[1] : nbs[0]);

                    prev = curr;
                    curr = next;

                    if (Math.Abs(curr.X - start.X) <= 1e-9 && Math.Abs(curr.Y - start.Y) <= 1e-9)
                    {
                        path.Add(start);
                        break;
                    }
                }

                if (path.Count < 5)
                {
                    continue;
                }

                var unique = path.ToList();
                if (unique.Count < 4)
                {
                    continue;
                }

                var area = Math.Abs(PolygonArea(unique));
                if (area > bestAnyArea)
                {
                    bestAnyArea = area;
                    bestAnyPath = unique;
                }

                var allOrthogonal = true;
                for (var i = 1; i < unique.Count; i++)
                {
                    var dx = Math.Abs(unique[i].X - unique[i - 1].X);
                    var dy = Math.Abs(unique[i].Y - unique[i - 1].Y);
                    if (dx > tol && dy > tol)
                    {
                        allOrthogonal = false;
                        break;
                    }
                }

                if (allOrthogonal && area > bestOrthogonalArea)
                {
                    bestOrthogonalArea = area;
                    bestOrthogonalPath = unique;
                }
            }

            return bestOrthogonalPath ?? bestAnyPath;
        }

        private static List<HashSet<(double X, double Y)>> GetConnectedComponents(
            IReadOnlyList<(double X, double Y)> nodes,
            Dictionary<(double X, double Y), List<(double X, double Y)>> adj)
        {
            var result = new List<HashSet<(double X, double Y)>>();
            var visited = new HashSet<(double X, double Y)>();

            foreach (var n in nodes)
            {
                if (visited.Contains(n))
                {
                    continue;
                }

                var comp = new HashSet<(double X, double Y)>();
                var q = new Queue<(double X, double Y)>();
                q.Enqueue(n);
                visited.Add(n);

                while (q.Count > 0)
                {
                    var cur = q.Dequeue();
                    comp.Add(cur);
                    if (!adj.TryGetValue(cur, out var nbs))
                    {
                        continue;
                    }

                    foreach (var nb in nbs)
                    {
                        if (visited.Add(nb))
                        {
                            q.Enqueue(nb);
                        }
                    }
                }

                result.Add(comp);
            }

            return result;
        }

        private static bool IsOnOuterRectangleEdge((double X, double Y) p, RectBounds outer, double tol)
        {
            var onLeft = Math.Abs(p.X - outer.MinX) <= tol;
            var onRight = Math.Abs(p.X - outer.MaxX) <= tol;
            var onBottom = Math.Abs(p.Y - outer.MinY) <= tol;
            var onTop = Math.Abs(p.Y - outer.MaxY) <= tol;
            return onLeft || onRight || onBottom || onTop;
        }

        private static List<(double X, double Y)> SimplifyStepPoints(
            IReadOnlyList<(double X, double Y)> points,
            bool byX,
            double epsilon)
        {
            if (points.Count == 0) return [];
            var result = new List<(double X, double Y)> { points[0] };
            for (var i = 1; i < points.Count; i++)
            {
                var prev = result[^1];
                var curr = points[i];
                var dv = byX ? Math.Abs(curr.Y - prev.Y) : Math.Abs(curr.X - prev.X);
                var du = byX ? Math.Abs(curr.X - prev.X) : Math.Abs(curr.Y - prev.Y);
                if (dv >= epsilon || du >= epsilon)
                {
                    result.Add(curr);
                }
            }
            return result;
        }

        private static List<EdgeCandidate> ExtractEdgePartialCandidates(DxfDocument doc, RectBounds outer)
        {
            var result = new List<EdgeCandidate>();
            var edgeTol = Compat.Clamp(Math.Min(outer.Width, outer.Height) * 0.0006, 0.08, 0.3);
            var connectTol = Math.Max(Math.Min(outer.Width, outer.Height) * 0.002, 1.0);
            var minElementLength = Math.Max(Math.Min(outer.Width, outer.Height) * 0.01, 5.0);
            var segments = new List<List<(double X, double Y)>>();
            var boardContour = SelectOuterContourPoints(doc, outer).ToList();
            if (boardContour.Count >= 2)
            {
                var first = boardContour[0];
                var last = boardContour[^1];
                var closeDist = Math.Sqrt((first.X - last.X) * (first.X - last.X) + (first.Y - last.Y) * (first.Y - last.Y));
                if (closeDist > edgeTol)
                {
                    boardContour.Add(first);
                }
            }

            static double PointToSegmentDistance((double X, double Y) p, (double X, double Y) a, (double X, double Y) b)
            {
                var vx = b.X - a.X;
                var vy = b.Y - a.Y;
                var wx = p.X - a.X;
                var wy = p.Y - a.Y;
                var len2 = vx * vx + vy * vy;
                if (len2 <= 1e-12)
                {
                    var dx0 = p.X - a.X;
                    var dy0 = p.Y - a.Y;
                    return Math.Sqrt(dx0 * dx0 + dy0 * dy0);
                }

                var t = Compat.Clamp((wx * vx + wy * vy) / len2, 0.0, 1.0);
                var projX = a.X + vx * t;
                var projY = a.Y + vy * t;
                var dx = p.X - projX;
                var dy = p.Y - projY;
                return Math.Sqrt(dx * dx + dy * dy);
            }

            bool IsOnRealBoardContour((double X, double Y) p)
            {
                if (boardContour.Count < 2)
                {
                    return Math.Abs(p.X - outer.MinX) <= edgeTol
                        || Math.Abs(p.X - outer.MaxX) <= edgeTol
                        || Math.Abs(p.Y - outer.MinY) <= edgeTol
                        || Math.Abs(p.Y - outer.MaxY) <= edgeTol;
                }

                for (var i = 1; i < boardContour.Count; i++)
                {
                    if (PointToSegmentDistance(p, boardContour[i - 1], boardContour[i]) <= edgeTol)
                    {
                        return true;
                    }
                }

                return false;
            }

            bool IsBoardEdgeSegment(IReadOnlyList<(double X, double Y)> points)
            {
                if (points.Count < 2)
                {
                    return true;
                }

                return points.All(IsOnRealBoardContour);
            }

            void AddSegment(IReadOnlyList<(double X, double Y)> points)
            {
                if (points.Count < 2 || IsBoardEdgeSegment(points))
                {
                    return;
                }

                var cleaned = new List<(double X, double Y)>();
                foreach (var p in points)
                {
                    if (cleaned.Count == 0)
                    {
                        cleaned.Add(p);
                        continue;
                    }

                    var last = cleaned[^1];
                    var d = Math.Sqrt((p.X - last.X) * (p.X - last.X) + (p.Y - last.Y) * (p.Y - last.Y));
                    if (d > 1e-6)
                    {
                        cleaned.Add(p);
                    }
                }

                if (cleaned.Count >= 2 && !IsBoardEdgeSegment(cleaned))
                {
                    segments.Add(cleaned);
                }
            }

            // 完整孔：闭合多段线、圆、整圆弧不进入边缘孔元素；板材边缘线也排除。
            foreach (var line in doc.Entities.Lines)
            {
                AddSegment([(line.StartPoint.X, line.StartPoint.Y), (line.EndPoint.X, line.EndPoint.Y)]);
            }

            foreach (var arc in doc.Entities.Arcs)
            {
                if (NormalizeArcSweep(arc.StartAngle, arc.EndAngle) >= 350.0)
                {
                    continue;
                }

                AddSegment(SampleArc(arc, 24));
            }

            foreach (var pl in doc.Entities.Polylines2D.Where(p => !IsPolylineClosedLike(p) && p.Vertexes.Count >= 2))
            {
                AddSegment(ExpandPolyline2D(pl, 24).ToList());
            }

            static double Distance((double X, double Y) a, (double X, double Y) b)
            {
                var dx = a.X - b.X;
                var dy = a.Y - b.Y;
                return Math.Sqrt(dx * dx + dy * dy);
            }

            static bool IsClosedCandidate(IReadOnlyList<(double X, double Y)> points, double tolerance)
            {
                if (points.Count < 3)
                {
                    return false;
                }

                return Distance(points[0], points[^1]) <= tolerance;
            }

            static void ReverseInPlace(List<(double X, double Y)> points)
            {
                points.Reverse();
            }

            var groups = new List<List<(double X, double Y)>>();
            foreach (var raw in segments)
            {
                var seg = raw.ToList();
                var merged = false;
                for (var i = 0; i < groups.Count; i++)
                {
                    var group = groups[i];
                    var gStart = group[0];
                    var gEnd = group[^1];
                    var sStart = seg[0];
                    var sEnd = seg[^1];

                    if (Distance(gEnd, sStart) <= connectTol)
                    {
                        group.AddRange(seg.Skip(1));
                        merged = true;
                        break;
                    }

                    if (Distance(gEnd, sEnd) <= connectTol)
                    {
                        ReverseInPlace(seg);
                        group.AddRange(seg.Skip(1));
                        merged = true;
                        break;
                    }

                    if (Distance(gStart, sEnd) <= connectTol)
                    {
                        group.InsertRange(0, seg.Take(seg.Count - 1));
                        merged = true;
                        break;
                    }

                    if (Distance(gStart, sStart) <= connectTol)
                    {
                        ReverseInPlace(seg);
                        group.InsertRange(0, seg.Take(seg.Count - 1));
                        merged = true;
                        break;
                    }
                }

                if (!merged)
                {
                    groups.Add(seg);
                }
            }

            // 二次合并，处理 A-B、C-D 先分组后又能接上的情况。
            var changed = true;
            while (changed)
            {
                changed = false;
                for (var i = 0; i < groups.Count && !changed; i++)
                {
                    for (var j = i + 1; j < groups.Count; j++)
                    {
                        var a = groups[i];
                        var b = groups[j];
                        var aStart = a[0];
                        var aEnd = a[^1];
                        var bStart = b[0];
                        var bEnd = b[^1];

                        if (Distance(aEnd, bStart) <= connectTol)
                        {
                            a.AddRange(b.Skip(1));
                        }
                        else if (Distance(aEnd, bEnd) <= connectTol)
                        {
                            b.Reverse();
                            a.AddRange(b.Skip(1));
                        }
                        else if (Distance(aStart, bEnd) <= connectTol)
                        {
                            a.InsertRange(0, b.Take(b.Count - 1));
                        }
                        else if (Distance(aStart, bStart) <= connectTol)
                        {
                            b.Reverse();
                            a.InsertRange(0, b.Take(b.Count - 1));
                        }
                        else
                        {
                            continue;
                        }

                        groups.RemoveAt(j);
                        changed = true;
                        break;
                    }
                }
            }

            foreach (var group in groups)
            {
                var length = PolylineLength(group);
                if (group.Count < 2 || length < minElementLength || IsBoardEdgeSegment(group))
                {
                    continue;
                }

                // 边缘孔只处理非完整孔的开放线条；多条 Line/Arc 拼成闭环后应视为完整孔，不能进入边缘孔。
                if (IsClosedCandidate(group, connectTol))
                {
                    continue;
                }

                var cx = group.Average(p => p.X);
                var cy = group.Average(p => p.Y);
                var sideDistances = new[]
                {
                    (Side: "Left", Distance: Math.Abs(cx - outer.MinX)),
                    (Side: "Right", Distance: Math.Abs(outer.MaxX - cx)),
                    (Side: "Bottom", Distance: Math.Abs(cy - outer.MinY)),
                    (Side: "Top", Distance: Math.Abs(outer.MaxY - cy))
                };
                var side = sideDistances.OrderBy(x => x.Distance).First().Side;
                result.Add(BuildEdgeCandidate(side, group));
            }

            return result;
        }

        private static EdgeCandidate BuildEdgeCandidate(string side, IReadOnlyList<(double X, double Y)> points)
        {
            var minX = points.Min(p => p.X);
            var maxX = points.Max(p => p.X);
            var minY = points.Min(p => p.Y);
            var maxY = points.Max(p => p.Y);
            var w = Math.Max(maxX - minX, 1.0);
            var h = Math.Max(maxY - minY, 1.0);
            double peri = 0;
            for (var i = 1; i < points.Count; i++)
            {
                var dx = points[i].X - points[i - 1].X;
                var dy = points[i].Y - points[i - 1].Y;
                peri += Math.Sqrt(dx * dx + dy * dy);
            }
            var cx = points.Average(p => p.X);
            var cy = points.Average(p => p.Y);
            var sig = CreatePolylineSignature(points, SignatureSamples);
            return new EdgeCandidate(side, points, (cx, cy), w, h, Math.Max(peri, 1.0), sig);
        }

        public static RawBounds GetRawBounds(DxfDocument doc)
        {
            var points = new List<(double X, double Y)>();

            points.AddRange(doc.Entities.Lines.SelectMany(l => new[] { (l.StartPoint.X, l.StartPoint.Y), (l.EndPoint.X, l.EndPoint.Y) }));
            points.AddRange(doc.Entities.Circles.SelectMany(c => new[]
            {
                (c.Center.X - c.Radius, c.Center.Y - c.Radius),
                (c.Center.X + c.Radius, c.Center.Y + c.Radius)
            }));
            points.AddRange(doc.Entities.Polylines2D.SelectMany(pl => ExpandPolyline2D(pl, 16)));
            points.AddRange(doc.Entities.Arcs.SelectMany(a => SampleArc(a, 72)));

            if (points.Count == 0)
            {
                return new RawBounds(0, 0, 100, 100);
            }

            var minX = points.Min(p => p.X);
            var minY = points.Min(p => p.Y);
            var maxX = points.Max(p => p.X);
            var maxY = points.Max(p => p.Y);
            return new RawBounds(minX, minY, maxX, maxY);
        }

        private static RectBounds DetectOuterRectangle(DxfDocument doc)
        {
            var bestArea = 0.0;
            RectBounds? bestRect = null;
            var all = GetRawBounds(doc);
            var allRect = new RectBounds(all.MinX, all.MinY, all.MaxX, all.MaxY);
            foreach (var pl in doc.Entities.Polylines2D.Where(p => p.IsClosed && p.Vertexes.Count >= 3))
            {
                var pts = ExpandPolyline2D(pl, 24).ToList();
                var area = Math.Abs(PolygonArea(pts));
                if (area < 1e-6)
                {
                    continue;
                }
                var xs = pl.Vertexes.Select(v => v.Position.X).ToList();
                var ys = pl.Vertexes.Select(v => v.Position.Y).ToList();
                var candidate = new RectBounds(xs.Min(), ys.Min(), xs.Max(), ys.Max());
                if (area > bestArea)
                {
                    bestArea = area;
                    bestRect = candidate;
                }
            }

            if (bestRect is not null && bestRect.Area > 0)
            {
                // If closed-polyline candidate is far smaller than global bounds,
                // the real outer contour is likely made by discrete lines/arcs.
                if (bestRect.Area >= allRect.Area * 0.5)
                {
                    return bestRect;
                }
            }
            return allRect;
        }

        private static List<HoleFeature> ExtractHoles(DxfDocument doc, bool includeOpenPolylines)
        {
            var holes = new List<HoleFeature>();
            var allBounds = GetRawBounds(doc);
            var allArea = Math.Max((allBounds.MaxX - allBounds.MinX) * (allBounds.MaxY - allBounds.MinY), 1.0);

            foreach (var c in doc.Entities.Circles)
            {
                var d = c.Radius * 2;
                var area = Math.PI * c.Radius * c.Radius;
                var perimeter = 2 * Math.PI * c.Radius;
                var circleMinX = c.Center.X - c.Radius;
                var circleMaxX = c.Center.X + c.Radius;
                var circleMinY = c.Center.Y - c.Radius;
                var circleMaxY = c.Center.Y + c.Radius;
                holes.Add(new HoleFeature(
                    "Circle",
                    ((circleMinX + circleMaxX) * 0.5, (circleMinY + circleMaxY) * 0.5),
                    d,
                    d,
                    area,
                    perimeter,
                    0,
                    CreateCircleSignature(c.Radius, SignatureSamples),
                    CreateCirclePoints(c.Center.X, c.Center.Y, c.Radius, SignatureSamples)));
            }

            foreach (var a in doc.Entities.Arcs)
            {
                var sweep = NormalizeArcSweep(a.StartAngle, a.EndAngle);
                if (sweep < 350.0)
                {
                    continue;
                }
                // Avoid duplicating with a real circle at same center/radius.
                var hasCircle = doc.Entities.Circles.Any(c =>
                    Math.Abs(c.Center.X - a.Center.X) < 1e-3 &&
                    Math.Abs(c.Center.Y - a.Center.Y) < 1e-3 &&
                    Math.Abs(c.Radius - a.Radius) < 1e-3);
                if (hasCircle)
                {
                    continue;
                }
                var d = a.Radius * 2;
                var area = Math.PI * a.Radius * a.Radius;
                var perimeter = 2 * Math.PI * a.Radius;
                var arcMinX = a.Center.X - a.Radius;
                var arcMaxX = a.Center.X + a.Radius;
                var arcMinY = a.Center.Y - a.Radius;
                var arcMaxY = a.Center.Y + a.Radius;
                holes.Add(new HoleFeature(
                    "ArcCircle",
                    ((arcMinX + arcMaxX) * 0.5, (arcMinY + arcMaxY) * 0.5),
                    d,
                    d,
                    area,
                    perimeter,
                    0,
                    CreateCircleSignature(a.Radius, SignatureSamples),
                    CreateCirclePoints(a.Center.X, a.Center.Y, a.Radius, SignatureSamples)));
            }

            foreach (var pl in doc.Entities.Polylines2D.Where(p => IsPolylineClosedLike(p) && p.Vertexes.Count >= 3))
            {
                var pts = ExpandPolyline2D(pl, 24).ToList();
                var area = PolygonArea(pts);
                if (Math.Abs(area) < 1e-6)
                {
                    continue;
                }
                var perimeter = PolygonPerimeter(pts);
                var minX = pts.Min(p => p.Item1);
                var maxX = pts.Max(p => p.Item1);
                var minY = pts.Min(p => p.Item2);
                var maxY = pts.Max(p => p.Item2);
                holes.Add(new HoleFeature(
                    "Polyline",
                    ((minX + maxX) * 0.5, (minY + maxY) * 0.5),
                    maxX - minX,
                    maxY - minY,
                    Math.Abs(area),
                    perimeter,
                    pl.Elevation,
                    CreatePolylineSignature(pts, SignatureSamples),
                    pts));
            }

            // 混合闭环（线段+圆弧）：把可闭合的 line/arc 组装成轮廓，识别“圆弧+线段孔”。
            foreach (var loop in ExtractMixedClosedLoops(doc))
            {
                if (loop.Count < 3)
                {
                    continue;
                }

                var area = Math.Abs(PolygonArea(loop));
                if (area < 1e-6)
                {
                    continue;
                }

                var minX = loop.Min(p => p.X);
                var maxX = loop.Max(p => p.X);
                var minY = loop.Min(p => p.Y);
                var maxY = loop.Max(p => p.Y);
                var perimeter = PolylineLength(loop);

                holes.Add(new HoleFeature(
                    "MixedArcLine",
                    ((minX + maxX) * 0.5, (minY + maxY) * 0.5),
                    maxX - minX,
                    maxY - minY,
                    area,
                    perimeter,
                    0,
                    CreatePolylineSignature(loop, SignatureSamples),
                    loop));
            }

            if (includeOpenPolylines)
            {
                foreach (var pl in doc.Entities.Polylines2D.Where(p => !IsPolylineClosedLike(p) && p.Vertexes.Count >= 3))
                {
                    var pts = ExpandPolyline2D(pl, 24).ToList();
                    if (pts.Count < 3)
                    {
                        continue;
                    }
                    var minX = pts.Min(p => p.X);
                    var maxX = pts.Max(p => p.X);
                    var minY = pts.Min(p => p.Y);
                    var maxY = pts.Max(p => p.Y);
                    var width = maxX - minX;
                    var height = maxY - minY;
                    var bboxArea = Math.Max(width * height, 0);
                    if (bboxArea < 1e-6 || bboxArea > allArea * 0.03)
                    {
                        continue;
                    }

                    var perimeter = PolylineLength(pts);
                    var pseudoArea = bboxArea * 0.6;
                    holes.Add(new HoleFeature(
                        "OpenPolyline",
                        ((minX + maxX) * 0.5, (minY + maxY) * 0.5),
                        width,
                        height,
                        pseudoArea,
                        perimeter,
                        pl.Elevation,
                        CreatePolylineSignature(pts, SignatureSamples),
                        pts));
                }
            }

            return DeduplicateHoles(holes);
        }

        public static IReadOnlyList<(double X, double Y)> SampleArc(Arc arc, int segments)
        {
            var startDeg = arc.StartAngle;
            var endDeg = arc.EndAngle;
            while (endDeg <= startDeg)
            {
                endDeg += 360.0;
            }
            var pts = new List<(double X, double Y)>(segments + 1);
            for (var i = 0; i <= segments; i++)
            {
                var t = (double)i / segments;
                var deg = startDeg + (endDeg - startDeg) * t;
                var rad = deg * Math.PI / 180.0;
                pts.Add((arc.Center.X + arc.Radius * Math.Cos(rad), arc.Center.Y + arc.Radius * Math.Sin(rad)));
            }
            return pts;
        }

        private static List<List<(double X, double Y)>> ExtractMixedClosedLoops(DxfDocument doc)
        {
            var segments = new List<(double X, double Y)[]>();

            foreach (var l in doc.Entities.Lines)
            {
                segments.Add([(l.StartPoint.X, l.StartPoint.Y), (l.EndPoint.X, l.EndPoint.Y)]);
            }

            foreach (var a in doc.Entities.Arcs)
            {
                var sweep = NormalizeArcSweep(a.StartAngle, a.EndAngle);
                if (sweep < 5.0 || sweep > 355.0)
                {
                    continue;
                }
                var pts = SampleArc(a, 24).ToArray();
                if (pts.Length >= 2)
                {
                    segments.Add(pts);
                }
            }

            var tol = 1.0;
            (double X, double Y) Snap((double X, double Y) p)
                => (Math.Round(p.X / tol) * tol, Math.Round(p.Y / tol) * tol);

            var edgeMap = new Dictionary<(double X, double Y), List<int>>();
            for (var i = 0; i < segments.Count; i++)
            {
                var s0 = Snap(segments[i][0]);
                var s1 = Snap(segments[i][^1]);

                if (!edgeMap.TryGetValue(s0, out var l0))
                {
                    l0 = [];
                    edgeMap[s0] = l0;
                }
                l0.Add(i);

                if (!edgeMap.TryGetValue(s1, out var l1))
                {
                    l1 = [];
                    edgeMap[s1] = l1;
                }
                l1.Add(i);
            }

            var used = new bool[segments.Count];
            var loops = new List<List<(double X, double Y)>>();

            for (var i = 0; i < segments.Count; i++)
            {
                if (used[i])
                {
                    continue;
                }

                var chain = new List<(double X, double Y)>(segments[i]);
                used[i] = true;

                var advanced = true;
                while (advanced)
                {
                    advanced = false;
                    var end = Snap(chain[^1]);
                    if (!edgeMap.TryGetValue(end, out var cands))
                    {
                        break;
                    }

                    foreach (var idx in cands)
                    {
                        if (used[idx])
                        {
                            continue;
                        }

                        var seg = segments[idx];
                        var s0 = Snap(seg[0]);
                        var s1 = Snap(seg[^1]);
                        if (s0.Equals(end))
                        {
                            chain.AddRange(seg.Skip(1));
                            used[idx] = true;
                            advanced = true;
                            break;
                        }
                        if (s1.Equals(end))
                        {
                            var reversed = seg.ToArray();
                            Array.Reverse(reversed);
                            chain.AddRange(reversed.Skip(1));
                            used[idx] = true;
                            advanced = true;
                            break;
                        }
                    }
                }

                if (chain.Count >= 4)
                {
                    var start = Snap(chain[0]);
                    var end = Snap(chain[^1]);
                    if (start.Equals(end))
                    {
                        loops.Add(chain);
                    }
                }
            }

            return loops;
        }

        public static IReadOnlyList<(double X, double Y)> ExpandPolyline2D(Polyline2D polyline, int bulgeSamplesPerSegment)
        {
            var result = new List<(double X, double Y)>();
            if (polyline.Vertexes.Count == 0)
            {
                return result;
            }

            for (var i = 0; i < polyline.Vertexes.Count; i++)
            {
                var current = polyline.Vertexes[i];
                var nextIndex = (i + 1) % polyline.Vertexes.Count;
                if (!polyline.IsClosed && nextIndex == 0)
                {
                    break;
                }
                var next = polyline.Vertexes[nextIndex];
                var p0 = (current.Position.X, current.Position.Y);
                var p1 = (next.Position.X, next.Position.Y);
                var bulge = current.Bulge;

                if (result.Count == 0)
                {
                    result.Add(p0);
                }

                if (Math.Abs(bulge) < 1e-9)
                {
                    result.Add(p1);
                    continue;
                }

                var arcPts = SampleBulgeArc(p0, p1, bulge, bulgeSamplesPerSegment);
                for (var k = 1; k < arcPts.Count; k++)
                {
                    result.Add(arcPts[k]);
                }
            }
            return result;
        }

        private static IReadOnlyList<(double X, double Y)> SampleBulgeArc((double X, double Y) p0, (double X, double Y) p1, double bulge, int segments)
        {
            var dx = p1.X - p0.X;
            var dy = p1.Y - p0.Y;
            var chord = Math.Sqrt(dx * dx + dy * dy);
            if (chord < 1e-12)
            {
                return [p0, p1];
            }

            var theta = 4.0 * Math.Atan(bulge);
            var radius = chord * (1.0 + bulge * bulge) / (4.0 * Math.Abs(bulge));
            var mx = (p0.X + p1.X) * 0.5;
            var my = (p0.Y + p1.Y) * 0.5;
            var nx = -dy / chord;
            var ny = dx / chord;
            var halfChord = chord * 0.5;
            var h = Math.Sqrt(Math.Max(0, radius * radius - halfChord * halfChord));
            var sign = bulge >= 0 ? 1.0 : -1.0;
            var cx = mx + sign * nx * h;
            var cy = my + sign * ny * h;

            var start = Math.Atan2(p0.Y - cy, p0.X - cx);
            var end = start + theta;

            var pts = new List<(double X, double Y)>(segments + 1);
            for (var i = 0; i <= segments; i++)
            {
                var t = (double)i / segments;
                var a = start + (end - start) * t;
                pts.Add((cx + radius * Math.Cos(a), cy + radius * Math.Sin(a)));
            }
            return pts;
        }

        private static bool IsPolylineClosedLike(Polyline2D p)
        {
            if (p.IsClosed)
            {
                return true;
            }
            if (p.Vertexes.Count < 3)
            {
                return false;
            }
            var first = p.Vertexes[0].Position;
            var last = p.Vertexes[^1].Position;
            var dx = first.X - last.X;
            var dy = first.Y - last.Y;
            return Math.Sqrt(dx * dx + dy * dy) <= 1e-3;
        }

        private static double NormalizeArcSweep(double startAngle, double endAngle)
        {
            var sweep = endAngle - startAngle;
            while (sweep < 0)
            {
                sweep += 360.0;
            }
            while (sweep > 360.0)
            {
                sweep -= 360.0;
            }
            return sweep;
        }

        private static List<HoleFeature> DeduplicateHoles(IEnumerable<HoleFeature> source)
        {
            static int TypePriority(string t)
            {
                if (t.StartsWith("Circle", StringComparison.OrdinalIgnoreCase) || t.StartsWith("ArcCircle", StringComparison.OrdinalIgnoreCase)) return 5;
                if (t.StartsWith("Polyline", StringComparison.OrdinalIgnoreCase)) return 4;
                if (t.StartsWith("MixedArcLine", StringComparison.OrdinalIgnoreCase)) return 3;
                if (t.StartsWith("EntityComposite", StringComparison.OrdinalIgnoreCase)) return 2;
                if (t.StartsWith("OpenPolyline", StringComparison.OrdinalIgnoreCase)) return 1;
                return 0;
            }

            var ordered = source
                .OrderByDescending(h => TypePriority(h.HoleType))
                .ThenByDescending(h => h.Area)
                .ToList();

            var result = new List<HoleFeature>();
            foreach (var h in ordered)
            {
                var tolPos = Math.Max(Math.Min(h.Width, h.Height) * 0.12, 1.5);
                var tolSize = Math.Max(Math.Min(h.Width, h.Height) * 0.10, 1.2);
                var exists = result.Any(r =>
                {
                    var dx = r.Centroid.X - h.Centroid.X;
                    var dy = r.Centroid.Y - h.Centroid.Y;
                    var near = Math.Sqrt(dx * dx + dy * dy) <= tolPos;
                    if (!near)
                    {
                        return false;
                    }

                    var sizeNear = Math.Abs(r.Width - h.Width) <= tolSize &&
                                   Math.Abs(r.Height - h.Height) <= tolSize;
                    var areaNear = Math.Abs(r.Area - h.Area) <= Math.Max(Math.Min(r.Area, h.Area) * 0.15, 8.0);
                    return sizeNear || areaNear;
                });

                if (!exists)
                {
                    result.Add(h);
                }
            }
            return result;
        }

        private static List<HoleFeature> ExtractCornerMissingFeatures(DxfDocument doc, RectBounds outer)
        {
            var features = new List<HoleFeature>();
            var pts = CollectGeometryPoints(doc);
            if (pts.Count == 0)
            {
                return features;
            }

            var zoneW = outer.Width * 0.22;
            var zoneH = outer.Height * 0.22;
            foreach (var c in outer.Corners)
            {
                var x0 = c.X <= (outer.MinX + outer.MaxX) * 0.5 ? outer.MinX : outer.MaxX - zoneW;
                var x1 = x0 + zoneW;
                var y0 = c.Y <= (outer.MinY + outer.MaxY) * 0.5 ? outer.MinY : outer.MaxY - zoneH;
                var y1 = y0 + zoneH;

                var inZone = pts.Where(p => p.X >= x0 && p.X <= x1 && p.Y >= y0 && p.Y <= y1).ToList();
                if (inZone.Count == 0)
                {
                    continue;
                }

                var cornerPoint = (c.X, c.Y);
                var minDist = inZone.Min(p =>
                {
                    var dx = p.X - cornerPoint.Item1;
                    var dy = p.Y - cornerPoint.Item2;
                    return Math.Sqrt(dx * dx + dy * dy);
                });
                var zoneDiag = Math.Sqrt(zoneW * zoneW + zoneH * zoneH);
                if (minDist < zoneDiag * 0.2)
                {
                    continue;
                }

                var sig = CreatePolylineSignature(inZone, SignatureSamples);
                features.Add(new HoleFeature(
                    $"CornerMissing:{c.Name}",
                    (x0 + zoneW * 0.5, y0 + zoneH * 0.5),
                    zoneW * 0.6,
                    zoneH * 0.6,
                    Math.Max(zoneW * zoneH * 0.18, 1.0),
                    Math.Max((zoneW + zoneH) * 0.6, 1.0),
                    0,
                    sig,
                    inZone));
            }
            return features;
        }

        private static List<HoleFeature> ExtractEdgeNotchFeatures(DxfDocument doc, RectBounds outer)
        {
            var result = new List<HoleFeature>();
            var pts = CollectGeometryPoints(doc);
            if (pts.Count == 0)
            {
                return result;
            }

            var cornerX = outer.Width * 0.22;
            var cornerY = outer.Height * 0.22;
            var band = Math.Max(Math.Min(outer.Width, outer.Height) * 0.08, 10.0);
            var depth = Math.Max(Math.Min(outer.Width, outer.Height) * 0.02, 4.0);

            void Build(IEnumerable<(double X, double Y)> points, bool byX, string side)
            {
                var list = points.ToList();
                if (list.Count < 3)
                {
                    return;
                }
                var sorted = byX ? list.OrderBy(p => p.X).ToList() : list.OrderBy(p => p.Y).ToList();
                var groups = new List<List<(double X, double Y)>>();
                var tol = byX ? outer.Width * 0.04 : outer.Height * 0.04;
                foreach (var p in sorted)
                {
                    if (groups.Count == 0)
                    {
                        groups.Add([p]);
                        continue;
                    }
                    var keyLast = byX ? groups[^1][^1].X : groups[^1][^1].Y;
                    var keyNow = byX ? p.X : p.Y;
                    if (Math.Abs(keyNow - keyLast) <= tol)
                    {
                        groups[^1].Add(p);
                    }
                    else
                    {
                        groups.Add([p]);
                    }
                }

                foreach (var g in groups.Where(g => g.Count >= 3))
                {
                    var minX = g.Min(p => p.X);
                    var maxX = g.Max(p => p.X);
                    var minY = g.Min(p => p.Y);
                    var maxY = g.Max(p => p.Y);
                    var w = Math.Max(maxX - minX, 1.0);
                    var h = Math.Max(maxY - minY, 1.0);
                    result.Add(new HoleFeature(
                        $"EdgeNotch:{side}",
                        (g.Average(p => p.X), g.Average(p => p.Y)),
                        w,
                        h,
                        Math.Max(w * h * 0.45, 1.0),
                        Math.Max(PolylineLength(g), 1.0),
                        0,
                        CreatePolylineSignature(g, SignatureSamples),
                        g));
                }
            }

            var top = pts.Where(p => p.Y <= outer.MaxY - depth && p.Y >= outer.MaxY - band &&
                                     p.X > outer.MinX + cornerX && p.X < outer.MaxX - cornerX);
            var bottom = pts.Where(p => p.Y >= outer.MinY + depth && p.Y <= outer.MinY + band &&
                                        p.X > outer.MinX + cornerX && p.X < outer.MaxX - cornerX);
            var left = pts.Where(p => p.X >= outer.MinX + depth && p.X <= outer.MinX + band &&
                                      p.Y > outer.MinY + cornerY && p.Y < outer.MaxY - cornerY);
            var right = pts.Where(p => p.X <= outer.MaxX - depth && p.X >= outer.MaxX - band &&
                                       p.Y > outer.MinY + cornerY && p.Y < outer.MaxY - cornerY);

            Build(top, true, "Top");
            Build(bottom, true, "Bottom");
            Build(left, false, "Left");
            Build(right, false, "Right");
            return result;
        }

        private static List<(double X, double Y)> CollectGeometryPoints(DxfDocument doc)
        {
            var pts = new List<(double X, double Y)>();
            pts.AddRange(doc.Entities.Lines.SelectMany(l => new[] { (l.StartPoint.X, l.StartPoint.Y), (l.EndPoint.X, l.EndPoint.Y) }));
            pts.AddRange(doc.Entities.Circles.SelectMany(c => SampleArc(new Arc(c.Center, c.Radius, 0, 360), 24)));
            pts.AddRange(doc.Entities.Arcs.SelectMany(a => SampleArc(a, 24)));
            pts.AddRange(doc.Entities.Polylines2D.SelectMany(pl => ExpandPolyline2D(pl, 24)));
            return pts;
        }

        private static HoleFeature? BuildFeatureFromEntities(DxfDocument doc)
        {
            var points = new List<(double X, double Y)>();
            double perimeter = 0;

            foreach (var l in doc.Entities.Lines)
            {
                var p0 = (l.StartPoint.X, l.StartPoint.Y);
                var p1 = (l.EndPoint.X, l.EndPoint.Y);
                points.Add(p0);
                points.Add(p1);
                var dx = p1.Item1 - p0.Item1;
                var dy = p1.Item2 - p0.Item2;
                perimeter += Math.Sqrt(dx * dx + dy * dy);
            }

            foreach (var a in doc.Entities.Arcs)
            {
                var arcPts = SampleArc(a, 24);
                points.AddRange(arcPts);
                var sweep = NormalizeArcSweep(a.StartAngle, a.EndAngle) * Math.PI / 180.0;
                perimeter += Math.Abs(a.Radius * sweep);
            }

            foreach (var p in doc.Entities.Polylines2D)
            {
                var pts = ExpandPolyline2D(p, 24);
                points.AddRange(pts);
                perimeter += PolylineLength(pts);
            }

            foreach (var c in doc.Entities.Circles)
            {
                perimeter += 2 * Math.PI * c.Radius;
                points.Add((c.Center.X - c.Radius, c.Center.Y - c.Radius));
                points.Add((c.Center.X + c.Radius, c.Center.Y + c.Radius));
            }

            if (points.Count < 2)
            {
                return null;
            }

            var minX = points.Min(p => p.X);
            var maxX = points.Max(p => p.X);
            var minY = points.Min(p => p.Y);
            var maxY = points.Max(p => p.Y);
            var width = maxX - minX;
            var height = maxY - minY;
            if (width < 1e-6 || height < 1e-6)
            {
                return null;
            }

            var area = width * height * 0.6;
            var signature = CreatePolylineSignature(points, SignatureSamples);
            return new HoleFeature(
                "EntityComposite",
                (points.Average(p => p.X), points.Average(p => p.Y)),
                width,
                height,
                area,
                Math.Max(perimeter, 1.0),
                0,
                signature,
                points);
        }

        private static List<(double X, double Y)> ExtractMoldOutline(DxfDocument doc, (double X, double Y) anchor, bool useBodyCenter)
        {
            var pts = CollectGeometryPoints(doc);
            if (pts.Count < 2)
            {
                return [];
            }

            // M01（连续冲压）使用模具本体几何中心，保证边缘与图纸边缘贴合；
            // 其余模具使用识别特征中心，保证单孔叠加不偏。
            double centerX;
            double centerY;
            if (useBodyCenter)
            {
                var minX = pts.Min(p => p.X);
                var maxX = pts.Max(p => p.X);
                var minY = pts.Min(p => p.Y);
                var maxY = pts.Max(p => p.Y);
                centerX = (minX + maxX) * 0.5;
                centerY = (minY + maxY) * 0.5;
            }
            else
            {
                centerX = anchor.X;
                centerY = anchor.Y;
            }

            var ordered = pts
                .Select(p => (X: p.X - centerX, Y: p.Y - centerY))
                .OrderBy(p => Math.Atan2(p.Y, p.X))
                .ThenByDescending(p => Math.Sqrt(p.X * p.X + p.Y * p.Y))
                .ToList();

            var step = Math.Max(1, ordered.Count / 80);
            var outline = new List<(double X, double Y)>();
            for (var i = 0; i < ordered.Count; i += step)
            {
                outline.Add((ordered[i].X, ordered[i].Y));
            }

            if (outline.Count > 0)
            {
                outline.Add(outline[0]);
            }
            return outline;
        }

        private static double PolygonArea(IReadOnlyList<(double X, double Y)> points)
        {
            double sum = 0;
            for (var i = 0; i < points.Count; i++)
            {
                var j = (i + 1) % points.Count;
                sum += points[i].X * points[j].Y - points[j].X * points[i].Y;
            }
            return sum / 2.0;
        }

        private static double PolygonPerimeter(IReadOnlyList<(double X, double Y)> points)
        {
            double sum = 0;
            for (var i = 0; i < points.Count; i++)
            {
                var j = (i + 1) % points.Count;
                var dx = points[i].X - points[j].X;
                var dy = points[i].Y - points[j].Y;
                sum += Math.Sqrt(dx * dx + dy * dy);
            }
            return sum;
        }

        private static double PolylineLength(IReadOnlyList<(double X, double Y)> points)
        {
            double sum = 0;
            for (var i = 1; i < points.Count; i++)
            {
                var dx = points[i].X - points[i - 1].X;
                var dy = points[i].Y - points[i - 1].Y;
                sum += Math.Sqrt(dx * dx + dy * dy);
            }
            return sum;
        }

        private static double[] CreateCircleSignature(double radius, int samples)
        {
            var normalized = new double[samples];
            for (var i = 0; i < samples; i++)
            {
                normalized[i] = 1.0;
            }
            return normalized;
        }

        private static IReadOnlyList<(double X, double Y)> CreateCirclePoints(double cx, double cy, double radius, int samples)
        {
            var points = new List<(double X, double Y)>(samples);
            for (var i = 0; i < samples; i++)
            {
                var angle = 2.0 * Math.PI * i / Math.Max(samples, 1);
                points.Add((cx + radius * Math.Cos(angle), cy + radius * Math.Sin(angle)));
            }
            return points;
        }

        private static double[] CreatePolylineSignature(IReadOnlyList<(double X, double Y)> points, int samples)
        {
            var closed = points.ToList();
            if (closed.Count == 0)
            {
                return CreateCircleSignature(1, samples);
            }
            if (closed[0] != closed[^1])
            {
                closed.Add(closed[0]);
            }

            var cumulative = new double[closed.Count];
            for (var i = 1; i < closed.Count; i++)
            {
                var dx = closed[i].X - closed[i - 1].X;
                var dy = closed[i].Y - closed[i - 1].Y;
                cumulative[i] = cumulative[i - 1] + Math.Sqrt(dx * dx + dy * dy);
            }

            var total = cumulative[^1];
            if (total < 1e-9)
            {
                return CreateCircleSignature(1, samples);
            }

            var sampled = new List<(double X, double Y)>(samples);
            for (var s = 0; s < samples; s++)
            {
                var target = total * s / samples;
                var seg = 1;
                while (seg < cumulative.Length && cumulative[seg] < target)
                {
                    seg++;
                }
                seg = Math.Min(seg, cumulative.Length - 1);
                var prev = seg - 1;
                var segLen = cumulative[seg] - cumulative[prev];
                var t = segLen < 1e-12 ? 0 : (target - cumulative[prev]) / segLen;
                var x = closed[prev].X + (closed[seg].X - closed[prev].X) * t;
                var y = closed[prev].Y + (closed[seg].Y - closed[prev].Y) * t;
                sampled.Add((x, y));
            }

            var cx = sampled.Average(p => p.X);
            var cy = sampled.Average(p => p.Y);
            var radial = sampled
                .Select(p =>
                {
                    var dx = p.X - cx;
                    var dy = p.Y - cy;
                    return Math.Sqrt(dx * dx + dy * dy);
                })
                .ToArray();

            var maxR = Math.Max(radial.Max(), 1e-9);
            for (var i = 0; i < radial.Length; i++)
            {
                radial[i] /= maxR;
            }
            return radial;
        }
    }

    public sealed class MoldMatcher
    {
        private const double CornerZoneRatio = 0.22;
        private const double CornerMaxNormalizedDistance = 0.16;
        private const double Mold1AbsoluteScoreThreshold = 0.55;
        private const double Mold1MarginFactor = 0.82;
        private const double EdgePartialDistanceRatio = 0.06;
        private const int CornerPathMaxPointsPerCorner = 24;

        // 严格匹配阈值（0.01精度 = ±1%）
        private const double StrictAreaRatioMin = 0.99;
        private const double StrictAreaRatioMax = 1.01;
        private const double StrictPerimRatioMin = 0.99;
        private const double StrictPerimRatioMax = 1.01;
        private const double StrictLongRatioMin = 0.99;
        private const double StrictLongRatioMax = 1.01;
        private const double StrictShortRatioMin = 0.99;
        private const double StrictShortRatioMax = 1.01;
        private const double StrictSigMax = 0.10;
        private const double ImpossibleMatchScoreThreshold = 0.50;

        public MatchResult Match(ProjectProfile project, IReadOnlyList<MoldProfile> molds, bool isStage1)
        {
            var rows = new List<HoleAssignment>();
            var guidePaths = new List<CornerStepPath>();
            if (molds.Count == 0 || (project.Holes.Count == 0 && project.EdgeCandidates.Count == 0))
            {
                return new MatchResult(rows, guidePaths);
            }

            var corners = project.OuterRectangle.Corners;
            var holePool = project.Holes.ToList();
            var validMolds = molds.Where(m => m.MoldId > 0 && m.MoldId < 999).ToList();
            if (validMolds.Count == 0)
            {
                validMolds = molds.ToList();
            }

            var mold1 = validMolds.FirstOrDefault(m => m.MoldId == 1) ?? validMolds[0];
            var nonCornerMolds = validMolds.Where(m => m.MoldId != mold1.MoldId).ToList();

            // M01：沿"青色差集线的外偏移路径"做连续冲压。
            var contourStamps = GenerateContinuousContourStampCenters(project, mold1, guidePaths, isStage1);

            // 当图纸四角无冲压区域时（contourStamps 为空），M01 无专属角落任务，
            // 退化为普通模具：重新纳入 nonCornerMolds 以识别内部及边缘孔。
            if (contourStamps.Count == 0)
            {
                nonCornerMolds = validMolds.ToList();
            }
            else if (nonCornerMolds.Count == 0)
            {
                nonCornerMolds.Add(mold1);
            }

            foreach (var s in contourStamps)
            {
                rows.Add(new HoleAssignment(
                    s,
                    mold1.MoldId,
                    "连续冲压",
                    true,
                    true,
                    $"M{mold1.MoldId:D2}:ContourPath"));
            }

            // 角落补冲先停用：当前 CornerMissing 候选是“角区特征中心”，会稳定产生四个固定点（四角各1个），
            // 与连续冲压路径重复，导致列表前四个 M01 点异常。
            // 后续如需恢复角补冲，应改成基于 contour pass 的真实残料区求交后再加点。
            var diag = Math.Sqrt(project.OuterRectangle.Width * project.OuterRectangle.Width +
                                 project.OuterRectangle.Height * project.OuterRectangle.Height);
            var safeDiag = Math.Max(diag, 1e-6);

            foreach (var hole in holePool)
            {
                // 调试日志：输出当前孔信息和可用的模具列表
                AppLogger.Instance.Info($"[匹配] 孔: {hole.HoleType}, W={hole.Width:F2}, H={hole.Height:F2}, A={hole.Area:F1}, Sig={hole.Signature:F2}");
                AppLogger.Instance.Info($"[匹配] nonCornerMolds列表: [{string.Join(", ", nonCornerMolds.Select(m => $"M{m.MoldId:D2}({m.Feature.HoleType}, W={m.Feature.Width:F1}, H={m.Feature.Height:F1}, A={m.Feature.Area:F1})"))}]");

                // 严格匹配：遍历模具库全部候选特征，必须满足同类+几何一致。
                var ranked = nonCornerMolds
                    .SelectMany(m =>
                    {
                        var features = (m.CandidateFeatures is { Count: > 0 } ? m.CandidateFeatures : [m.Feature]);
                        AppLogger.Instance.Info($"[匹配] M{m.MoldId:D2} 使用特征: {string.Join(", ", features.Select(f => $"W={f.Width:F2},H={f.Height:F2},A={f.Area:F1},Type={f.HoleType}"))}");
                        return features
                            .Where(f => IsShapeFamilyCompatible(hole, f))
                            .Select(f =>
                            {
                                var areaRatio = hole.Area / Math.Max(f.Area, 1e-6);
                                var perimRatio = hole.Perimeter / Math.Max(f.Perimeter, 1e-6);
                                var signature = SignatureDistance(hole.Signature, f.Signature);
                                var typeMatch = IsSameShapeType(hole, f);

                                var hLong = Math.Max(hole.Width, hole.Height);
                                var hShort = Math.Max(Math.Min(hole.Width, hole.Height), 1e-6);
                                var fLong = Math.Max(f.Width, f.Height);
                                var fShort = Math.Max(Math.Min(f.Width, f.Height), 1e-6);
                                var longRatio = hLong / Math.Max(fLong, 1e-6);
                                var shortRatio = hShort / Math.Max(fShort, 1e-6);

                                var strict = typeMatch
                                    && areaRatio >= StrictAreaRatioMin && areaRatio <= StrictAreaRatioMax
                                    && perimRatio >= StrictPerimRatioMin && perimRatio <= StrictPerimRatioMax
                                    && longRatio >= StrictLongRatioMin && longRatio <= StrictLongRatioMax
                                    && shortRatio >= StrictShortRatioMin && shortRatio <= StrictShortRatioMax
                                    && signature <= StrictSigMax;

                                var score = Math.Abs(areaRatio - 1.0)
                                            + Math.Abs(perimRatio - 1.0)
                                            + Math.Abs(longRatio - 1.0)
                                            + Math.Abs(shortRatio - 1.0)
                                            + signature * 0.5;
                                var impossible = !typeMatch
                                               || areaRatio < 0.80 || areaRatio > 1.20
                                               || perimRatio < 0.85 || perimRatio > 1.15
                                               || longRatio < 0.85 || longRatio > 1.15
                                               || shortRatio < 0.85 || shortRatio > 1.15
                                               || signature > 0.40
                                               || score > ImpossibleMatchScoreThreshold;

                                return new
                                {
                                    MoldId = m.MoldId,
                                    AreaRatio = areaRatio,
                                    PerimRatio = perimRatio,
                                    LongRatio = longRatio,
                                    ShortRatio = shortRatio,
                                    Signature = signature,
                                    TypeMatch = typeMatch,
                                    Strict = strict,
                                    Impossible = impossible,
                                    Score = score
                                };
                            });
                    })
                    .OrderBy(x => x.Score)
                    .ToList();

                if (ranked.Count == 0)
                {
                    continue;
                }

                AppLogger.Instance.Info($"[匹配] 孔(W={hole.Width:F2},H={hole.Height:F2}): ranked.Count={ranked.Count}, strictPass.Count={ranked.Count(x => x.Strict)}");
                var strictPass = ranked.Where(x => x.Strict).OrderBy(x => x.Score).ToList();
                if (strictPass.Count == 0)
                {
                    AppLogger.Instance.Info($"[匹配] 孔(W={hole.Width:F2},H={hole.Height:F2}): 严格条件未通过，跳过匹配");
                    continue;
                }

                var pick = strictPass[0];
                rows.Add(new HoleAssignment(
                    hole,
                    pick.MoldId,
                    "单次冲压",
                    IsAnyCornerZone(hole, project.OuterRectangle),
                    IsNearOuterEdge(hole, project.OuterRectangle),
                    string.Join(" | ", strictPass.Take(3).Select(r => $"{(isStage1 ? "M" : "N")}{r.MoldId:D2}:{r.Score:F1}")),
                    $"A={pick.AreaRatio:F1},P={pick.PerimRatio:F1}"));
            }

            foreach (var edge in project.EdgeCandidates)
            {
                var edgeHole = new HoleFeature(
                    $"EdgePartial:{edge.Side}",
                    edge.Centroid,
                    edge.Width,
                    edge.Height,
                    Math.Max(edge.Width * edge.Height * 0.45, 1.0),
                    Math.Max(edge.Perimeter, 1.0),
                    0,
                    edge.Signature,
                    edge.Points);

                var partial = TryPartialBBoxContourMatch(edgeHole, nonCornerMolds, isStage1);
                AppLogger.Instance.Info($"[边缘匹配] {edgeHole.HoleType}: W={edgeHole.Width:F2},H={edgeHole.Height:F2},Area={edgeHole.Area:F2},Pos=({edgeHole.Centroid.X:F2},{edgeHole.Centroid.Y:F2}), 匹配结果={(partial is null ? "无" : $"M{partial.MoldId:D2}(Cover={partial.Coverage:P1},D={partial.TrimmedDistance:F2})")}");
                if (partial is null)
                {
                    continue;
                }

                var placementHole = new HoleFeature(
                    edgeHole.HoleType,
                    partial.Placement,
                    edgeHole.Width,
                    edgeHole.Height,
                    edgeHole.Area,
                    edgeHole.Perimeter,
                    edgeHole.Rotation,
                    edgeHole.Signature,
                    edgeHole.Points);

                var prefix = isStage1 ? "M" : "N";
                rows.Add(new HoleAssignment(
                    placementHole,
                    partial.MoldId,
                    "边缘孔局部冲压",
                    false,
                    true,
                    $"{prefix}{partial.MoldId:D2}:EdgePartial={partial.Score:F2},Cover={partial.Coverage:P0},Corner={partial.CornerName},Align=({partial.Placement.X:F1},{partial.Placement.Y:F1})",
                    $"EdgePartial,W={partial.WidthRatio:F2},H={partial.HeightRatio:F2},D={partial.TrimmedDistance:F2}",
                    "边缘点对齐后使用模具定位点",
                    0,
                    false));
            }

            var cleaned = DeduplicateAssignments(rows);
            var ordered = OrderAssignmentsForStamping(cleaned, guidePaths, project.OuterRectangle);
            return new MatchResult(ordered, guidePaths);
        }

        /// <summary>
        /// 冲压顺序：M01 连续冲压按板料竖中线分块——先左侧沿路径，再内孔（X 再 Y），再右侧沿路径；EdgeNotch 置尾。
        /// </summary>
        public static IReadOnlyList<HoleAssignment> OrderAssignmentsForStamping(
            IReadOnlyList<HoleAssignment> assignments,
            IReadOnlyList<CornerStepPath>? guidePaths,
            RectBounds outer)
        {
            if (assignments.Count <= 1)
            {
                return assignments;
            }

            static bool IsEdgeNotch(HoleAssignment a) =>
                a.Hole.HoleType.StartsWith("EdgeNotch:", StringComparison.Ordinal);

            static bool IsCornerSupplement(HoleAssignment a) =>
                string.Equals(a.PositionRelation, "连续冲压-角落补冲", StringComparison.Ordinal)
                || a.Hole.HoleType.StartsWith("CornerMissing:", StringComparison.Ordinal);

            static bool IsContourStamp(HoleAssignment a) =>
                string.Equals(a.PositionRelation, "连续冲压", StringComparison.Ordinal)
                || a.Hole.HoleType.StartsWith("ContourCornerHit:", StringComparison.Ordinal)
                || a.Hole.HoleType.StartsWith("ContourPath:", StringComparison.Ordinal);

            static string? ContourPathTag(HoleAssignment a)
            {
                var t = a.Hole.HoleType;
                const string p1 = "ContourCornerHit:";
                const string p2 = "ContourPath:";
                if (t.StartsWith(p1, StringComparison.Ordinal))
                {
                    return t.Substring(p1.Length);
                }

                if (t.StartsWith(p2, StringComparison.Ordinal))
                {
                    return t.Substring(p2.Length);
                }

                return null;
            }

            static int ContourTagOrder(string tag)
            {
                if (tag.Length > 7
                    && tag.StartsWith("Contour", StringComparison.Ordinal)
                    && int.TryParse(tag.Substring(7), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n))
                {
                    return n;
                }

                return int.MaxValue;
            }

            var pathOrder = new Dictionary<string, int>(StringComparer.Ordinal);
            if (guidePaths is not null)
            {
                for (var i = 0; i < guidePaths.Count; i++)
                {
                    var name = guidePaths[i].CornerName;
                    if (!pathOrder.ContainsKey(name))
                    {
                        pathOrder[name] = i;
                    }
                }
            }

            IReadOnlyList<(double X, double Y)>? GuideChain(string pathTag)
            {
                if (guidePaths is null)
                {
                    return null;
                }

                foreach (var gp in guidePaths)
                {
                    if (string.Equals(gp.CornerName, pathTag, StringComparison.Ordinal) && gp.Points is { Count: >= 2 })
                    {
                        return gp.Points;
                    }
                }

                return null;
            }

            var notches = assignments.Where(IsEdgeNotch).ToList();
            var work = assignments.Where(a => !IsEdgeNotch(a)).ToList();

            var cornerSupplement = work.Where(IsCornerSupplement).ToList();
            var contour = work.Where(a => IsContourStamp(a) && !IsCornerSupplement(a)).ToList();
            var inner = work.Where(a => !IsContourStamp(a) && !IsCornerSupplement(a)).ToList();

            // M01 连续点按“每个角落独立排序”：
            // 每个角内部：先自上而下（Y降序），再向内收敛（离中线越近越后）。
            var allContour = contour.Concat(cornerSupplement).ToList();
            var midX = outer.MinX + outer.Width * 0.5;

            RectCorner NearestCorner((double X, double Y) p)
            {
                return outer.Corners
                    .OrderBy(c =>
                    {
                        var dx = p.X - c.X;
                        var dy = p.Y - c.Y;
                        return dx * dx + dy * dy;
                    })
                    .First();
            }

            int CornerOrder(RectCorner c)
            {
                var top = c.Y >= (outer.MinY + outer.MaxY) * 0.5;
                var left = c.X <= (outer.MinX + outer.MaxX) * 0.5;

                // 右侧角落整体后置：先左侧两角，再右侧两角。
                if (top && left) return 0;    // 左上
                if (!top && left) return 1;   // 左下
                if (top && !left) return 2;   // 右上
                return 3;                     // 右下
            }

            var contourByCorner = allContour
                .GroupBy(a => NearestCorner(a.Hole.Centroid))
                .OrderBy(g => CornerOrder(g.Key))
                .Select(g =>
                {
                    var corner = g.Key;
                    var top = corner.Y >= (outer.MinY + outer.MaxY) * 0.5;
                    var left = corner.X <= (outer.MinX + outer.MaxX) * 0.5;

                    var avgWidth = g.Any() ? g.Average(x => Math.Max(x.Hole.Width, 1.0)) : 1.0;
                    var layerStep = Math.Max(avgWidth * 0.5, 1.0);

                    double InwardDistance(HoleAssignment a)
                    {
                        var x = a.Hole.Centroid.X;
                        var y = a.Hole.Centroid.Y;
                        var dx = left ? (x - outer.MinX) : (outer.MaxX - x);
                        var dy = top ? (outer.MaxY - y) : (y - outer.MinY);
                        return Math.Min(Math.Max(dx, 0), Math.Max(dy, 0));
                    }

                    var ordered = g
                        .Select(a => new { Row = a, Inward = InwardDistance(a) })
                        .OrderBy(x => (int)Math.Floor(x.Inward / layerStep))
                        .ThenBy(x => top ? -x.Row.Hole.Centroid.Y : x.Row.Hole.Centroid.Y)
                        .ThenBy(x => left ? x.Row.Hole.Centroid.X : -x.Row.Hole.Centroid.X)
                        .Select(x => x.Row)
                        .ToList();

                    return new { IsLeft = left, Rows = ordered };
                })
                .ToList();

            var contourLeftOrdered = contourByCorner
                .Where(x => x.IsLeft)
                .SelectMany(x => x.Rows)
                .ToList();

            var contourRightOrdered = contourByCorner
                .Where(x => !x.IsLeft)
                .SelectMany(x => x.Rows)
                .ToList();

            var innerOrdered = inner
                .OrderBy(a => a.Hole.Centroid.X)
                .ThenBy(a => a.Hole.Centroid.Y)
                .ToList();

            var notchesOrdered = notches
                .OrderBy(a => a.Hole.Centroid.X)
                .ThenBy(a => a.Hole.Centroid.Y)
                .ToList();

            return contourLeftOrdered
                .Concat(innerOrdered)
                .Concat(notchesOrdered)
                .Concat(contourRightOrdered)
                .ToList();
        }

        private static double ClosestPathStation(IReadOnlyList<(double X, double Y)> chain, (double X, double Y) p)
        {
            double bestDist2 = double.PositiveInfinity;
            double bestStation = 0;
            var cum = 0.0;
            for (var i = 1; i < chain.Count; i++)
            {
                var ax = chain[i - 1].X;
                var ay = chain[i - 1].Y;
                var bx = chain[i].X;
                var by = chain[i].Y;
                var abx = bx - ax;
                var aby = by - ay;
                var segLen = Math.Sqrt(abx * abx + aby * aby);
                double stationOnSeg;
                double px;
                double py;
                if (segLen <= 1e-12)
                {
                    stationOnSeg = 0;
                    px = ax;
                    py = ay;
                }
                else
                {
                    var apx = p.X - ax;
                    var apy = p.Y - ay;
                    var t = (apx * abx + apy * aby) / (segLen * segLen);
                    if (t < 0)
                    {
                        t = 0;
                    }
                    else if (t > 1)
                    {
                        t = 1;
                    }

                    stationOnSeg = t * segLen;
                    px = ax + t * abx;
                    py = ay + t * aby;
                }

                var dx = p.X - px;
                var dy = p.Y - py;
                var d2 = dx * dx + dy * dy;
                if (d2 < bestDist2 - 1e-18)
                {
                    bestDist2 = d2;
                    bestStation = cum + stationOnSeg;
                }
                else if (Math.Abs(d2 - bestDist2) < 1e-12)
                {
                    var alt = cum + stationOnSeg;
                    if (alt < bestStation)
                    {
                        bestStation = alt;
                    }
                }

                cum += segLen;
            }

            return bestStation;
        }

        private static double TrimmedChamferScore(
            IReadOnlyList<(double X, double Y)> moldOutlineCentered,
            IReadOnlyList<(double X, double Y)> candidatePointsAbs,
            out (double X, double Y) bestPlacement)
        {
            bestPlacement = (0, 0);
            if (moldOutlineCentered.Count < 8 || candidatePointsAbs.Count < 12)
            {
                return double.PositiveInfinity;
            }

            var cx = candidatePointsAbs.Average(p => p.X);
            var cy = candidatePointsAbs.Average(p => p.Y);
            var basePlacement = (X: cx, Y: cy);

            var searchRadius = 25.0;
            var step = 2.5;
            var best = double.PositiveInfinity;
            for (var dx = -searchRadius; dx <= searchRadius; dx += step)
            {
                for (var dy = -searchRadius; dy <= searchRadius; dy += step)
                {
                    var placement = (basePlacement.X + dx, basePlacement.Y + dy);
                    var score = TrimmedChamferOnce(moldOutlineCentered, candidatePointsAbs, placement, trimRatio: 0.65);
                    if (score < best)
                    {
                        best = score;
                        bestPlacement = placement;
                    }
                }
            }
            return best;
        }

        private static double TrimmedChamferOnce(
            IReadOnlyList<(double X, double Y)> moldOutlineCentered,
            IReadOnlyList<(double X, double Y)> candidatePointsAbs,
            (double X, double Y) placement,
            double trimRatio)
        {
            var n = moldOutlineCentered.Count;
            var d2 = new double[n];
            for (var i = 0; i < n; i++)
            {
                var mx = moldOutlineCentered[i].X + placement.X;
                var my = moldOutlineCentered[i].Y + placement.Y;
                var best = double.PositiveInfinity;
                for (var j = 0; j < candidatePointsAbs.Count; j++)
                {
                    var dx = mx - candidatePointsAbs[j].X;
                    var dy = my - candidatePointsAbs[j].Y;
                    var dist = dx * dx + dy * dy;
                    if (dist < best) best = dist;
                }
                d2[i] = best;
            }
            Array.Sort(d2);
            var keep = Math.Max(6, (int)Math.Round(n * trimRatio));
            double sum = 0;
            for (var i = 0; i < keep; i++) sum += d2[i];
            return Math.Sqrt(sum / keep);
        }

        private static RectCorner FindNearestCorner(RectBounds rect, (double X, double Y) point)
        {
            return rect.Corners
                .OrderBy(c =>
                {
                    var dx = point.X - c.X;
                    var dy = point.Y - c.Y;
                    return dx * dx + dy * dy;
                })
                .First();
        }

        private static (double X, double Y) PushPointTowardCorner(
            (double X, double Y) point,
            RectCorner corner,
            double pushX,
            double pushY)
        {
            var sx = corner.X >= point.X ? 1.0 : -1.0;
            var sy = corner.Y >= point.Y ? 1.0 : -1.0;
            return (point.X + sx * pushX, point.Y + sy * pushY);
        }

        private static IReadOnlyList<HoleFeature> GenerateContinuousContourStampCenters(ProjectProfile project, MoldProfile mold1, List<CornerStepPath> guidePaths, bool isStage1)
        {
            var contourPaths = isStage1 ? project.Stage1ContourPaths : project.Stage2ContourPaths;
            if (contourPaths is null || contourPaths.Count == 0)
            {
                return [];
            }

            var outline = mold1.OutlinePoints;
            if (outline is null || outline.Count < 2)
            {
                return [];
            }

            // 用 M01 轮廓外包尺寸，而不是 Feature 尺寸（Feature 是识别特征，不一定等于模具本体）。
            var minOx = outline.Min(p => p.X);
            var maxOx = outline.Max(p => p.X);
            var minOy = outline.Min(p => p.Y);
            var maxOy = outline.Max(p => p.Y);
            var moldOutlineWidth = Math.Max(maxOx - minOx, 1.0);
            var moldOutlineHeight = Math.Max(maxOy - minOy, 1.0);
            var moldEdgeLength = Math.Max(moldOutlineWidth, moldOutlineHeight);

            // 首次外偏移按半个模具尺寸，后续重复外移按完整 M01 长/宽。
            var initialOffsetX = Math.Max(moldOutlineWidth * 0.5, 0.8);
            var initialOffsetY = Math.Max(moldOutlineHeight * 0.5, 0.8);
            var repeatOffsetX = Math.Max(moldOutlineWidth, 1.6);
            var repeatOffsetY = Math.Max(moldOutlineHeight, 1.6);
            var points = new List<HoleFeature>();

            static bool IsPolylineOutsideWorkingZone(IReadOnlyList<(double X, double Y)> pts, RectBounds rect, double marginX, double marginY)
            {
                return pts.Count > 0 && pts.All(p =>
                    p.X < rect.MinX - marginX ||
                    p.X > rect.MaxX + marginX ||
                    p.Y < rect.MinY - marginY ||
                    p.Y > rect.MaxY + marginY);
            }

            static IReadOnlyList<(double X, double Y)> TrimPolylineEndpointsByDistance(
                IReadOnlyList<(double X, double Y)> polyline,
                double trimStart,
                double trimEnd)
            {
                if (polyline.Count < 2)
                {
                    return [];
                }

                var lengths = new double[polyline.Count];
                var total = 0.0;
                for (var i = 1; i < polyline.Count; i++)
                {
                    var dx = polyline[i].X - polyline[i - 1].X;
                    var dy = polyline[i].Y - polyline[i - 1].Y;
                    total += Math.Sqrt(dx * dx + dy * dy);
                    lengths[i] = total;
                }

                if (total <= 1e-9)
                {
                    return [];
                }

                trimStart = Math.Max(0, trimStart);
                trimEnd = Math.Max(0, trimEnd);
                var startStation = Math.Min(trimStart, total);
                var endStation = Math.Max(startStation, total - trimEnd);
                if (endStation - startStation <= 1e-9)
                {
                    return [];
                }

                (double X, double Y) SampleAt(double station)
                {
                    if (station <= 0)
                    {
                        return polyline[0];
                    }

                    if (station >= total)
                    {
                        return polyline[^1];
                    }

                    for (var i = 1; i < lengths.Length; i++)
                    {
                        if (lengths[i] + 1e-12 < station)
                        {
                            continue;
                        }

                        var segStart = lengths[i - 1];
                        var segLen = lengths[i] - segStart;
                        if (segLen <= 1e-12)
                        {
                            return polyline[i];
                        }

                        var t = (station - segStart) / segLen;
                        var a = polyline[i - 1];
                        var b = polyline[i];
                        return (a.X + (b.X - a.X) * t, a.Y + (b.Y - a.Y) * t);
                    }

                    return polyline[^1];
                }

                var result = new List<(double X, double Y)> { SampleAt(startStation) };
                for (var i = 1; i < polyline.Count - 1; i++)
                {
                    if (lengths[i] > startStation + 1e-9 && lengths[i] < endStation - 1e-9)
                    {
                        var p = polyline[i];
                        var last = result[^1];
                        if (Math.Sqrt((p.X - last.X) * (p.X - last.X) + (p.Y - last.Y) * (p.Y - last.Y)) > 1e-9)
                        {
                            result.Add(p);
                        }
                    }
                }
                result.Add(SampleAt(endStation));

                var cleaned = new List<(double X, double Y)>();
                foreach (var p in result)
                {
                    if (cleaned.Count == 0)
                    {
                        cleaned.Add(p);
                        continue;
                    }

                    var last = cleaned[^1];
                    if (Math.Sqrt((p.X - last.X) * (p.X - last.X) + (p.Y - last.Y) * (p.Y - last.Y)) > 1e-9)
                    {
                        cleaned.Add(p);
                    }
                }

                return cleaned;
            }

            string BuildPassName(string baseName, int passIndex)
            {
                return passIndex == 1 ? baseName : $"{baseName}_P{passIndex}";
            }

            double GetStepLengthForDirection(double ux, double uy)
            {
                var dirToolLength = Math.Max(Math.Abs(ux) * moldOutlineWidth + Math.Abs(uy) * moldOutlineHeight, 1.0);
                var overlapMm = Compat.Clamp(dirToolLength * 0.15, 2.0, 10.0);
                return Math.Max(dirToolLength - overlapMm, 0.5);
            }

            double GetHalfTrimForDirection(double ux, double uy)
            {
                return GetStepLengthForDirection(ux, uy) * 0.5;
            }

            void EmitPassPoints(string passName, IReadOnlyList<(double X, double Y)> pathPoints)
            {
                if (pathPoints.Count < 2)
                {
                    return;
                }

                var keyPoints = new List<(double X, double Y)> { pathPoints[0] };
                for (var i = 1; i < pathPoints.Count - 1; i++)
                {
                    var a = pathPoints[i - 1];
                    var b = pathPoints[i];
                    var c = pathPoints[i + 1];
                    var abx = b.X - a.X;
                    var aby = b.Y - a.Y;
                    var bcx = c.X - b.X;
                    var bcy = c.Y - b.Y;
                    var lab = Math.Sqrt(abx * abx + aby * aby);
                    var lbc = Math.Sqrt(bcx * bcx + bcy * bcy);
                    if (lab <= 1e-9 || lbc <= 1e-9)
                    {
                        continue;
                    }

                    var cross = Math.Abs(abx * bcy - aby * bcx) / (lab * lbc);
                    if (cross > 0.02)
                    {
                        keyPoints.Add(b);
                    }
                }
                keyPoints.Add(pathPoints[^1]);

                foreach (var v in keyPoints)
                {
                    points.Add(new HoleFeature(
                        $"ContourCornerHit:{passName}",
                        v,
                        mold1.Feature.Width,
                        mold1.Feature.Height,
                        Math.Max(mold1.Feature.Area, 1.0),
                        Math.Max(mold1.Feature.Perimeter, 1.0),
                        0,
                        mold1.Feature.Signature));
                }

                for (var i = 1; i < pathPoints.Count; i++)
                {
                    var a = pathPoints[i - 1];
                    var b = pathPoints[i];
                    var dx = b.X - a.X;
                    var dy = b.Y - a.Y;
                    var segLen = Math.Sqrt(dx * dx + dy * dy);
                    if (segLen <= 1e-9)
                    {
                        continue;
                    }

                    var ux = dx / segLen;
                    var uy = dy / segLen;
                    var moldStepLength = GetStepLengthForDirection(ux, uy);
                    if (segLen <= moldStepLength + 1e-9)
                    {
                        continue;
                    }

                    for (var traveled = moldStepLength; traveled < segLen - 1e-9; traveled += moldStepLength)
                    {
                        var p = (a.X + ux * traveled, a.Y + uy * traveled);
                        points.Add(new HoleFeature(
                            $"ContourPath:{passName}",
                            p,
                            mold1.Feature.Width,
                            mold1.Feature.Height,
                            Math.Max(mold1.Feature.Area, 1.0),
                            Math.Max(mold1.Feature.Perimeter, 1.0),
                            0,
                            mold1.Feature.Signature));
                    }
                }
            }

            foreach (var contourPath in contourPaths)
            {
                var pts = contourPath.Points;
                if (pts is null || pts.Count < 2)
                {
                    continue;
                }

                var chain = pts
                    .DistinctBy(p => ($"{Math.Round(p.X, 4)}|{Math.Round(p.Y, 4)}"))
                    .ToList();
                if (chain.Count < 2)
                {
                    continue;
                }

                var passSource = OffsetPolylineOutward(chain, project.OuterRectangle, initialOffsetX, initialOffsetY);
                if (passSource.Count < 2)
                {
                    continue;
                }

                var endpointExtend = Math.Max(moldEdgeLength * 0.5, 0.8);
                var passIndex = 1;
                var currentPass = passSource;
                while (currentPass.Count >= 2 && passIndex <= 12)
                {
                    var extendedPass = ExtendPolylineEndpoints(currentPass, endpointExtend, project.OuterRectangle);
                    if (extendedPass.Count < 2)
                    {
                        break;
                    }

                    var passName = BuildPassName(contourPath.CornerName, passIndex);
                    guidePaths.Add(new CornerStepPath(passName, extendedPass));
                    EmitPassPoints(passName, extendedPass);

                    if (IsPolylineOutsideWorkingZone(extendedPass, project.OuterRectangle, moldOutlineWidth, moldOutlineHeight))
                    {
                        break;
                    }

                    var nextPass = OffsetPolylineOutward(extendedPass, project.OuterRectangle, repeatOffsetX, repeatOffsetY);
                    if (nextPass.Count < 2)
                    {
                        break;
                    }

                    currentPass = nextPass;
                    passIndex++;
                }
            }

            // 规则：不要求“完整落模在板内”，只要模具与板材有交集即可冲。
            bool IntersectsBoard((double X, double Y) center)
            {
                var minX = center.X - moldOutlineWidth * 0.5;
                var maxX = center.X + moldOutlineWidth * 0.5;
                var minY = center.Y - moldOutlineHeight * 0.5;
                var maxY = center.Y + moldOutlineHeight * 0.5;

                return maxX >= project.OuterRectangle.MinX &&
                       minX <= project.OuterRectangle.MaxX &&
                       maxY >= project.OuterRectangle.MinY &&
                       minY <= project.OuterRectangle.MaxY;
            }

            var dedup = new List<HoleFeature>();
            const double minKeepDistance = 3.0; // 固定：相邻冲点中心距 >= 3mm 才保留
            foreach (var p in points)
            {
                if (!IntersectsBoard(p.Centroid))
                {
                    continue;
                }

                var tooClose = dedup.Any(d =>
                {
                    var dx = d.Centroid.X - p.Centroid.X;
                    var dy = d.Centroid.Y - p.Centroid.Y;
                    return Math.Sqrt(dx * dx + dy * dy) < minKeepDistance;
                });
                if (!tooClose)
                {
                    dedup.Add(p);
                }
            }

            return dedup;
        }

        private static double EstimateOutlineStep(IReadOnlyList<(double X, double Y)> outline)
        {
            // 目标：给连续冲压采样一个“像 CAD Offset 后沿轮廓走刀”的稳定步距。
            // 思路：
            // 1) 去掉重复点与极短噪声边，避免步距被噪点拉小；
            // 2) 优先统计主方向（水平/垂直）边长；
            // 3) 用稳健分位值（Q40）并做 IQR 去极值，减少异常短边/长边影响。
            if (outline is null || outline.Count < 2)
            {
                return 10.0;
            }

            var eps = 1e-6;
            var clean = new List<(double X, double Y)>();
            foreach (var p in outline)
            {
                if (clean.Count == 0)
                {
                    clean.Add(p);
                    continue;
                }

                var prev = clean[^1];
                var d = Math.Sqrt((p.X - prev.X) * (p.X - prev.X) + (p.Y - prev.Y) * (p.Y - prev.Y));
                if (d > eps)
                {
                    clean.Add(p);
                }
            }

            if (clean.Count >= 2)
            {
                var first = clean[0];
                var last = clean[^1];
                var close = Math.Sqrt((first.X - last.X) * (first.X - last.X) + (first.Y - last.Y) * (first.Y - last.Y));
                if (close <= eps)
                {
                    clean.RemoveAt(clean.Count - 1);
                }
            }

            if (clean.Count < 2)
            {
                return 10.0;
            }

            var allLens = new List<double>();
            var axisLens = new List<double>();

            for (var i = 1; i < clean.Count; i++)
            {
                var dx = clean[i].X - clean[i - 1].X;
                var dy = clean[i].Y - clean[i - 1].Y;
                var len = Math.Sqrt(dx * dx + dy * dy);
                if (len <= eps)
                {
                    continue;
                }

                allLens.Add(len);

                // 主方向边（近似水平/垂直）优先，符合冲压路径多为正交折线的实际。
                var minComp = Math.Min(Math.Abs(dx), Math.Abs(dy));
                var maxComp = Math.Max(Math.Abs(dx), Math.Abs(dy));
                if (maxComp > eps && minComp / maxComp <= 0.08)
                {
                    axisLens.Add(len);
                }
            }

            var baseLens = axisLens.Count >= 3 ? axisLens : allLens;
            if (baseLens.Count == 0)
            {
                return 10.0;
            }

            baseLens.Sort();

            // IQR 去极值（稳健）
            double PickQuantile(List<double> arr, double q)
            {
                if (arr.Count == 1)
                {
                    return arr[0];
                }

                var pos = (arr.Count - 1) * q;
                var i0 = (int)Math.Floor(pos);
                var i1 = Math.Min(i0 + 1, arr.Count - 1);
                var t = pos - i0;
                return arr[i0] * (1 - t) + arr[i1] * t;
            }

            var q1 = PickQuantile(baseLens, 0.25);
            var q3 = PickQuantile(baseLens, 0.75);
            var iqr = Math.Max(q3 - q1, eps);
            var lo = Math.Max(q1 - 1.5 * iqr, eps);
            var hi = q3 + 1.5 * iqr;

            var trimmed = baseLens.Where(v => v >= lo && v <= hi).ToList();
            if (trimmed.Count == 0)
            {
                trimmed = baseLens;
            }
            trimmed.Sort();

            // Q40：比中位数略偏小，既能覆盖拐角短段，又不会被极短段主导。
            var step = PickQuantile(trimmed, 0.40);

            // 限幅，避免异常数据导致采样过密/过疏。
            var maxLen = trimmed[^1];
            step = Compat.Clamp(step, 2.0, Math.Max(2.0, maxLen * 0.85));
            return step;
        }

        private static List<(double X, double Y)> OffsetPolylineOutward(
            IReadOnlyList<(double X, double Y)> chain,
            RectBounds outer,
            double offsetX,
            double offsetY)
        {
            var result = new List<(double X, double Y)>();
            if (chain.Count < 2)
            {
                return result;
            }

            var center = ((outer.MinX + outer.MaxX) * 0.5, (outer.MinY + outer.MaxY) * 0.5);
            const double eps = 1e-9;

            // 先清理重复点，避免零长度段造成角点错位。
            var clean = new List<(double X, double Y)>();
            foreach (var p in chain)
            {
                if (clean.Count == 0)
                {
                    clean.Add(p);
                    continue;
                }

                var last = clean[^1];
                if (Math.Sqrt((p.X - last.X) * (p.X - last.X) + (p.Y - last.Y) * (p.Y - last.Y)) > eps)
                {
                    clean.Add(p);
                }
            }

            if (clean.Count < 2)
            {
                return result;
            }

            // 关键：偏移前先净化路径，去掉“极短边 + 近共线伪拐点”，避免产生额外角点。
            // 这能消除图中 2~6 一带那种由微小抖动引入的多余折线。
            var axisMergeTol = Math.Max(Math.Min(outer.Width, outer.Height) * 0.002, 0.6);
            bool changed;
            do
            {
                changed = false;

                // 1) 删除极短边中间点
                for (var i = 1; i < clean.Count - 1; i++)
                {
                    var a = clean[i - 1];
                    var b = clean[i];
                    var c = clean[i + 1];
                    var ab = Math.Sqrt((b.X - a.X) * (b.X - a.X) + (b.Y - a.Y) * (b.Y - a.Y));
                    var bc = Math.Sqrt((c.X - b.X) * (c.X - b.X) + (c.Y - b.Y) * (c.Y - b.Y));
                    if (ab <= axisMergeTol || bc <= axisMergeTol)
                    {
                        clean.RemoveAt(i);
                        changed = true;
                        break;
                    }
                }
                if (changed)
                {
                    continue;
                }

                // 2) 删除近共线点（水平/竖直折线场景）
                for (var i = 1; i < clean.Count - 1; i++)
                {
                    var a = clean[i - 1];
                    var b = clean[i];
                    var c = clean[i + 1];

                    var abx = b.X - a.X;
                    var aby = b.Y - a.Y;
                    var bcx = c.X - b.X;
                    var bcy = c.Y - b.Y;

                    var lab = Math.Sqrt(abx * abx + aby * aby);
                    var lbc = Math.Sqrt(bcx * bcx + bcy * bcy);
                    if (lab <= eps || lbc <= eps)
                    {
                        clean.RemoveAt(i);
                        changed = true;
                        break;
                    }

                    var cross = Math.Abs(abx * bcy - aby * bcx) / (lab * lbc);
                    if (cross <= 0.01)
                    {
                        clean.RemoveAt(i);
                        changed = true;
                        break;
                    }
                }
            } while (changed && clean.Count >= 3);

            if (clean.Count < 2)
            {
                return result;
            }

            ((double NX, double NY) Normal, double Dist) OutwardNormal((double X, double Y) a, (double X, double Y) b)
            {
                var dx = b.X - a.X;
                var dy = b.Y - a.Y;
                var len = Math.Sqrt(dx * dx + dy * dy);
                if (len <= eps)
                {
                    return ((0, 0), 0);
                }

                var tx = dx / len;
                var ty = dy / len;
                var left = (-ty, tx);
                var right = (ty, -tx);
                var mid = ((a.X + b.X) * 0.5, (a.Y + b.Y) * 0.5);

                // 正交轮廓优先：按“更靠近哪条外框边”确定外侧方向，避免中心距离法在拐角处翻转。
                var axisTol = 1e-4;
                var isHorizontal = Math.Abs(dy) <= axisTol;
                var isVertical = Math.Abs(dx) <= axisTol;

                if (isHorizontal)
                {
                    var distTop = Math.Abs(outer.MaxY - mid.Item2);
                    var distBottom = Math.Abs(mid.Item2 - outer.MinY);
                    var ny = distTop <= distBottom ? 1.0 : -1.0;
                    return ((0.0, ny), offsetY);
                }

                if (isVertical)
                {
                    var distLeft = Math.Abs(mid.Item1 - outer.MinX);
                    var distRight = Math.Abs(outer.MaxX - mid.Item1);
                    var nx = distLeft <= distRight ? -1.0 : 1.0;
                    return ((nx, 0.0), offsetX);
                }

                // 非正交段：按法向分量分别使用 X/Y 偏移，比较外侧方向。
                var ldx = mid.Item1 + left.Item1 * offsetX - center.Item1;
                var ldy = mid.Item2 + left.Item2 * offsetY - center.Item2;
                var rdx = mid.Item1 + right.Item1 * offsetX - center.Item1;
                var rdy = mid.Item2 + right.Item2 * offsetY - center.Item2;
                var dl = ldx * ldx + ldy * ldy;
                var dr = rdx * rdx + rdy * rdy;

                var chosen = dl >= dr ? left : right;
                // 对斜段给一个合成偏移距离，用于后续交点求解。
                var dist = Math.Sqrt((chosen.Item1 * offsetX) * (chosen.Item1 * offsetX) + (chosen.Item2 * offsetY) * (chosen.Item2 * offsetY));
                return (chosen, Math.Max(dist, 0.8));
            }

            (double X, double Y)? LineIntersection(
                (double X, double Y) p,
                (double X, double Y) r,
                (double X, double Y) q,
                (double X, double Y) s)
            {
                var rxs = r.X * s.Y - r.Y * s.X;
                if (Math.Abs(rxs) <= eps)
                {
                    return null;
                }

                var qmp = (q.X - p.X, q.Y - p.Y);
                var t = (qmp.Item1 * s.Y - qmp.Item2 * s.X) / rxs;
                return (p.X + t * r.X, p.Y + t * r.Y);
            }

            var segNormals = new List<((double NX, double NY) Normal, double Dist)>();
            for (var i = 1; i < clean.Count; i++)
            {
                segNormals.Add(OutwardNormal(clean[i - 1], clean[i]));
            }

            // 起点 = 首段起点偏移
            result.Add((
                clean[0].X + segNormals[0].Normal.NX * segNormals[0].Dist,
                clean[0].Y + segNormals[0].Normal.NY * segNormals[0].Dist));

            // 中间角点：相邻偏移线求交；失败时走“正交桥接点”，避免斜线短接。
            for (var i = 1; i < clean.Count - 1; i++)
            {
                var prev = clean[i - 1];
                var curr = clean[i];
                var next = clean[i + 1];

                var n1 = segNormals[i - 1];
                var n2 = segNormals[i];

                var p1 = (prev.X + n1.Normal.NX * n1.Dist, prev.Y + n1.Normal.NY * n1.Dist);
                var p2 = (curr.X + n1.Normal.NX * n1.Dist, curr.Y + n1.Normal.NY * n1.Dist);
                var r = (p2.Item1 - p1.Item1, p2.Item2 - p1.Item2);

                var p3 = (curr.X + n2.Normal.NX * n2.Dist, curr.Y + n2.Normal.NY * n2.Dist);
                var p4 = (next.X + n2.Normal.NX * n2.Dist, next.Y + n2.Normal.NY * n2.Dist);
                var s = (p4.Item1 - p3.Item1, p4.Item2 - p3.Item2);

                var cross = LineIntersection((p1.Item1, p1.Item2), (r.Item1, r.Item2), (p3.Item1, p3.Item2), (s.Item1, s.Item2));
                if (cross.HasValue)
                {
                    // miter 限幅，防止尖角飞点。
                    var dx = cross.Value.X - curr.X;
                    var dy = cross.Value.Y - curr.Y;
                    var dist = Math.Sqrt(dx * dx + dy * dy);
                    if (dist <= Math.Max(Math.Max(n1.Dist, n2.Dist) * 5.0, 1.0))
                    {
                        result.Add(cross.Value);
                        continue;
                    }
                }

                // 退化：只保留“当前拐点对应的一个角点”，然后按角点顺序连线。
                // 不再插入 p2/p3，避免产生额外折返段。
                var b1 = (p2.Item1, p3.Item2);
                var b2 = (p3.Item1, p2.Item2);

                var d1x = b1.Item1 - center.Item1;
                var d1y = b1.Item2 - center.Item2;
                var d2x = b2.Item1 - center.Item1;
                var d2y = b2.Item2 - center.Item2;
                var bridge = (d1x * d1x + d1y * d1y) >= (d2x * d2x + d2y * d2y) ? b1 : b2;

                result.Add(bridge);
            }

            // 终点 = 尾段终点偏移
            var lastNormal = segNormals[^1];
            result.Add((
                clean[^1].X + lastNormal.Normal.NX * lastNormal.Dist,
                clean[^1].Y + lastNormal.Normal.NY * lastNormal.Dist));

            // 去除相邻重复点
            var cleaned = new List<(double X, double Y)>();
            foreach (var p in result)
            {
                if (cleaned.Count == 0)
                {
                    cleaned.Add(p);
                    continue;
                }

                var last = cleaned[^1];
                if (Math.Sqrt((p.X - last.X) * (p.X - last.X) + (p.Y - last.Y) * (p.Y - last.Y)) > 1e-6)
                {
                    cleaned.Add(p);
                }
            }

            return cleaned;
        }

        private static List<(double X, double Y)> ExtendPolylineEndpoints(
            IReadOnlyList<(double X, double Y)> polyline,
            double extendDistance,
            RectBounds outer)
        {
            var result = polyline.ToList();
            if (result.Count < 2 || extendDistance <= 1e-9)
            {
                return result;
            }

            var eps = 1e-9;

            (double X, double Y) ExtendByOutward((double X, double Y) anchor, (double X, double Y) neighbor, bool isStart)
            {
                var dx = neighbor.X - anchor.X;
                var dy = neighbor.Y - anchor.Y;
                var len = Math.Sqrt(dx * dx + dy * dy);
                if (len <= eps)
                {
                    return anchor;
                }

                var ux = dx / len;
                var uy = dy / len;

                // 两个候选：沿线方向 / 反方向，选“更外侧”的那个
                var c1 = isStart
                    ? (anchor.X - ux * extendDistance, anchor.Y - uy * extendDistance)
                    : (anchor.X + ux * extendDistance, anchor.Y + uy * extendDistance);
                var c2 = isStart
                    ? (anchor.X + ux * extendDistance, anchor.Y + uy * extendDistance)
                    : (anchor.X - ux * extendDistance, anchor.Y - uy * extendDistance);

                var d1 = DistanceToRectCenter(c1, outer);
                var d2 = DistanceToRectCenter(c2, outer);
                return d1 >= d2 ? c1 : c2;
            }

            // 首端：自动判外侧，修复“0点反向”
            result[0] = ExtendByOutward(result[0], result[1], isStart: true);

            // 末端：自动判外侧
            result[^1] = ExtendByOutward(result[^1], result[^2], isStart: false);

            return result;
        }

        private static double DistanceToRectCenter((double X, double Y) p, RectBounds rect)
        {
            var cx = (rect.MinX + rect.MaxX) * 0.5;
            var cy = (rect.MinY + rect.MaxY) * 0.5;
            var dx = p.X - cx;
            var dy = p.Y - cy;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        private sealed record PathSample(double X, double Y, double TX, double TY);

        private static List<PathSample> SampleAlongPolyline(IReadOnlyList<(double X, double Y)> polyline, double step)
        {
            return SampleAlongPolylineWithMinSegmentLength(polyline, step, minSegmentLength: 0.0);
        }

        private static List<PathSample> SampleAlongPolylineWithMinSegmentLength(
            IReadOnlyList<(double X, double Y)> polyline,
            double step,
            double minSegmentLength)
        {
            var result = new List<PathSample>();
            if (polyline.Count < 2)
            {
                return result;
            }

            var firstDx = polyline[1].X - polyline[0].X;
            var firstDy = polyline[1].Y - polyline[0].Y;
            var firstLen = Math.Sqrt(firstDx * firstDx + firstDy * firstDy);
            if (firstLen > Math.Max(1e-9, minSegmentLength))
            {
                result.Add(new PathSample(polyline[0].X, polyline[0].Y, firstDx / firstLen, firstDy / firstLen));
            }

            var remain = step;

            for (var i = 1; i < polyline.Count; i++)
            {
                var a = polyline[i - 1];
                var b = polyline[i];
                var dx = b.X - a.X;
                var dy = b.Y - a.Y;
                var segLen = Math.Sqrt(dx * dx + dy * dy);
                if (segLen <= 1e-9)
                {
                    continue;
                }

                if (segLen < minSegmentLength)
                {
                    // 小于模具边长的短边：不做连续点采样，仅由拐点命中兜底。
                    continue;
                }

                var ux = dx / segLen;
                var uy = dy / segLen;
                var progressed = 0.0;

                while (progressed + remain <= segLen + 1e-9)
                {
                    progressed += remain;
                    var px = a.X + ux * progressed;
                    var py = a.Y + uy * progressed;
                    result.Add(new PathSample(px, py, ux, uy));
                    remain = step;
                }

                remain -= (segLen - progressed);
                if (remain <= 1e-9)
                {
                    remain = step;
                }
            }

            var tail = polyline[^1];
            if (result.Count == 0)
            {
                result.Add(new PathSample(tail.X, tail.Y, 1, 0));
                return result;
            }

            var last = result[^1];
            if (Math.Abs(last.X - tail.X) > 1e-6 || Math.Abs(last.Y - tail.Y) > 1e-6)
            {
                var prev = polyline[^2];
                var dx = tail.X - prev.X;
                var dy = tail.Y - prev.Y;
                var len = Math.Sqrt(dx * dx + dy * dy);
                if (len <= 1e-9)
                {
                    result.Add(new PathSample(tail.X, tail.Y, last.TX, last.TY));
                }
                else if (len >= minSegmentLength)
                {
                    result.Add(new PathSample(tail.X, tail.Y, dx / len, dy / len));
                }
            }

            return result;
        }

        private static IReadOnlyList<HoleAssignment> DeduplicateAssignments(IReadOnlyList<HoleAssignment> source)
        {
            if (source.Count <= 1)
            {
                return source;
            }

            var bestByHole = new Dictionary<string, HoleAssignment>(StringComparer.Ordinal);
            foreach (var row in source)
            {
                var key = BuildHoleKey(row.Hole);
                if (!bestByHole.TryGetValue(key, out var current))
                {
                    bestByHole[key] = row;
                    continue;
                }

                var currentScore = AssignmentPriority(current);
                var newScore = AssignmentPriority(row);
                if (newScore < currentScore)
                {
                    bestByHole[key] = row;
                }
            }

            return bestByHole.Values.ToList();
        }

        private static string BuildHoleKey(HoleFeature hole)
        {
            return $"{Math.Round(hole.Centroid.X, 4):F4}|{Math.Round(hole.Centroid.Y, 4):F4}|{Math.Round(hole.Width, 4):F4}|{Math.Round(hole.Height, 4):F4}|{hole.HoleType}";
        }

        private static double AssignmentPriority(HoleAssignment row)
        {
            var moldPenalty = row.MoldId == 1 ? 0.0 : 1.0;
            var edgePenalty = row.Hole.HoleType.StartsWith("EdgePartial:", StringComparison.Ordinal) ? 0.4 : 0.0;
            var cornerBonus = row.IsCornerCandidate ? -0.1 : 0.0;
            return moldPenalty + edgePenalty + cornerBonus;
        }

        private static bool IsInCornerMissingZone(HoleFeature hole, RectBounds rect, RectCorner corner)
        {
            var zoneX = rect.Width * CornerZoneRatio;
            var zoneY = rect.Height * CornerZoneRatio;
            var nearX = corner.X <= (rect.MinX + rect.MaxX) * 0.5
                ? hole.Centroid.X <= rect.MinX + zoneX
                : hole.Centroid.X >= rect.MaxX - zoneX;
            var nearY = corner.Y <= (rect.MinY + rect.MaxY) * 0.5
                ? hole.Centroid.Y <= rect.MinY + zoneY
                : hole.Centroid.Y >= rect.MaxY - zoneY;
            if (!(nearX && nearY))
            {
                return false;
            }

            // Corner hole must also be near the two adjacent outer edges.
            var edgeThreshold = Math.Max(Math.Min(rect.Width, rect.Height) * 0.05, 1.0);
            var nearVertical = corner.X <= (rect.MinX + rect.MaxX) * 0.5
                ? Math.Abs(hole.Centroid.X - rect.MinX) <= edgeThreshold
                : Math.Abs(rect.MaxX - hole.Centroid.X) <= edgeThreshold;
            var nearHorizontal = corner.Y <= (rect.MinY + rect.MaxY) * 0.5
                ? Math.Abs(hole.Centroid.Y - rect.MinY) <= edgeThreshold
                : Math.Abs(rect.MaxY - hole.Centroid.Y) <= edgeThreshold;
            return nearVertical && nearHorizontal;
        }

        private static bool IsAnyCornerZone(HoleFeature hole, RectBounds rect)
        {
            return rect.Corners.Any(c => IsInCornerMissingZone(hole, rect, c));
        }

        private static double ScoreForHoleAgainstMold(HoleFeature hole, HoleFeature mold, RectBounds rect)
        {
            var fullScore = SimilarityScore(hole, mold);
            if (!IsNearOuterEdge(hole, rect))
            {
                return fullScore;
            }

            var partialScore = PartialEdgeScore(hole, mold);
            return Math.Min(fullScore, partialScore);
        }

        private static bool IsNearOuterEdge(HoleFeature hole, RectBounds rect)
        {
            var minEdge = Math.Min(rect.Width, rect.Height);
            var threshold = Math.Max(minEdge * EdgePartialDistanceRatio, 1.0);
            var dx = Math.Min(Math.Abs(hole.Centroid.X - rect.MinX), Math.Abs(rect.MaxX - hole.Centroid.X));
            var dy = Math.Min(Math.Abs(hole.Centroid.Y - rect.MinY), Math.Abs(rect.MaxY - hole.Centroid.Y));
            return Math.Min(dx, dy) <= threshold;
        }

        private static double PartialEdgeScore(HoleFeature hole, HoleFeature mold)
        {
            // Edge holes can be partially stamped: allow smaller area/perimeter while
            // keeping shape ratio and signature close.
            var wRatio = hole.Width / Math.Max(mold.Width, 1e-6);
            var hRatio = hole.Height / Math.Max(mold.Height, 1e-6);
            var aRatio = hole.Area / Math.Max(mold.Area, 1e-6);
            var pRatio = hole.Perimeter / Math.Max(mold.Perimeter, 1e-6);

            var dw = Math.Abs(1.0 - Math.Min(wRatio, 1.0));
            var dh = Math.Abs(1.0 - Math.Min(hRatio, 1.0));
            var da = Math.Abs(1.0 - Math.Min(aRatio, 1.0));
            var dp = Math.Abs(1.0 - Math.Min(pRatio, 1.0));
            var dr = Math.Abs((hole.Width / Math.Max(hole.Height, 1e-6)) - (mold.Width / Math.Max(mold.Height, 1e-6)));
            var ds = SignatureDistance(hole.Signature, mold.Signature);
            return 0.12 * dw + 0.12 * dh + 0.16 * da + 0.1 * dp + 0.15 * dr + 0.1 * ds;
        }

        private static string BuildTopCandidates(HoleFeature hole, IEnumerable<MoldProfile> molds, RectBounds rect, int topN, bool isStage1)
        {
            var prefix = isStage1 ? "M" : "N";
            var tops = molds
                .Select(m => new
                {
                    m.MoldId,
                    Score = ScoreForHoleAgainstMold(hole, m.Feature, rect)
                })
                .OrderBy(x => x.Score)
                .Take(topN)
                .Select(x => $"{prefix}{x.MoldId:D2}:{x.Score:F1}");
            return string.Join(" | ", tops);
        }

        private static string BuildTopCandidatesByAreaRatio(HoleFeature hole, IEnumerable<MoldProfile> molds, int topN, bool isStage1)
        {
            var prefix = isStage1 ? "M" : "N";
            var tops = molds
                .SelectMany(m =>
                {
                    var features = (m.CandidateFeatures is { Count: > 0 } ? m.CandidateFeatures : [m.Feature]);
                    return features.Select(f => new
                    {
                        m.MoldId,
                        AreaRatio = hole.Area / Math.Max(f.Area, 1e-6),
                        Signature = SignatureDistance(hole.Signature, f.Signature)
                    });
                })
                .OrderBy(x => Math.Abs(x.AreaRatio - 1.0))
                .ThenBy(x => x.Signature)
                .Take(topN)
                .Select(x => $"{prefix}{x.MoldId:D2}:{x.AreaRatio:F1}");
            return string.Join(" | ", tops);
        }

        private sealed record PartialMatchCandidate(
            int MoldId,
            string CornerName,
            (double X, double Y) Placement,
            double Score,
            double Coverage,
            double TrimmedDistance,
            double WidthRatio,
            double HeightRatio);

        private sealed record PartialMatchAnchor(string Name, (double X, double Y) Point);

        private sealed record PartialMoldCache(
            MoldProfile Mold,
            IReadOnlyList<(double X, double Y)> Points,
            IReadOnlyList<PartialMatchAnchor> Anchors,
            RectBounds Bounds);

        private static readonly string EdgePartialDebugLogPath = System.IO.Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "edge-partial-debug.log");

        private static void AppendEdgePartialDebugLog(string message)
        {
            try
            {
                File.AppendAllText(EdgePartialDebugLogPath, message + Environment.NewLine, Encoding.UTF8);
            }
            catch
            {
                // 调试日志不能影响正常识别流程。
            }
        }

        private static PartialMatchCandidate? TryPartialBBoxContourMatch(HoleFeature hole, IReadOnlyList<MoldProfile> molds, bool isStage1)
        {
            var holePoints = NormalizeFeaturePoints(hole);
            var isEdgePartial = hole.HoleType.StartsWith("EdgePartial:", StringComparison.Ordinal);
            if (isEdgePartial)
            {
                var box = holePoints.Count > 0 ? BoundsOf(holePoints) : new RectBounds(0, 0, 0, 0);
                AppendEdgePartialDebugLog(
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] Begin {hole.HoleType}, Stage={(isStage1 ? "M" : "N")}, " +
                    $"Center=({hole.Centroid.X:F3},{hole.Centroid.Y:F3}), Size=({hole.Width:F3},{hole.Height:F3}), " +
                    $"Area={hole.Area:F3}, Perimeter={hole.Perimeter:F3}, Points={holePoints.Count}, " +
                    $"Box=({box.MinX:F3},{box.MinY:F3})-({box.MaxX:F3},{box.MaxY:F3}), Molds={molds.Count}");
            }

            var minHolePointCount = isEdgePartial ? 3 : 8;
            if (holePoints.Count < minHolePointCount || hole.Perimeter < 10.0)
            {
                if (isEdgePartial)
                {
                    AppendEdgePartialDebugLog($"  Reject before mold scan: Points={holePoints.Count} (<{minHolePointCount}) or Perimeter={hole.Perimeter:F3} (<10).");
                }
                return null;
            }

            var holeBounds = BoundsOf(holePoints);
            var holeAnchors = BuildPartialMatchAnchors(holePoints, "孔", isEdgePartial ? 10 : 18);
            var moldCaches = molds
                .Select(m =>
                {
                    var points = NormalizeMoldPoints(m);
                    return new PartialMoldCache(
                        m,
                        points,
                        BuildPartialMatchAnchors(points, "模", isEdgePartial ? 10 : 18),
                        points.Count > 0 ? BoundsOf(points) : new RectBounds(0, 0, 0, 0));
                })
                .ToList();

            PartialMatchCandidate? best = null;
            var scannedMoldCount = 0;
            foreach (var moldCache in moldCaches)
            {
                var mold = moldCache.Mold;
                var moldPoints = moldCache.Points;
                var rejectReason = GetPartialMatchRejectReason(hole, mold.Feature, moldPoints.Count, isEdgePartial);
                if (rejectReason is null && isEdgePartial)
                {
                    rejectReason = GetEdgePartialPrefilterRejectReason(hole, holeBounds, mold.Feature, moldCache.Bounds);
                }

                if (rejectReason is not null)
                {
                    if (isEdgePartial)
                    {
                        AppendEdgePartialDebugLog($"  Mold M{mold.MoldId:D2} skipped: {rejectReason}; MoldType={mold.Feature.HoleType}, Size=({mold.Feature.Width:F3},{mold.Feature.Height:F3}), Area={mold.Feature.Area:F3}, MoldPoints={moldPoints.Count}.");
                    }
                    continue;
                }

                scannedMoldCount++;
                var candidate = TryPartialBBoxContourMatch(hole, holePoints, holeAnchors, moldCache, isEdgePartial);
                if (candidate is null)
                {
                    if (isEdgePartial)
                    {
                        AppendEdgePartialDebugLog($"  Mold M{mold.MoldId:D2} no pass: all anchor pairs failed thresholds Cover>=80%, D<=1.8.");
                    }
                    continue;
                }

                if (isEdgePartial)
                {
                    AppendEdgePartialDebugLog($"  Mold M{mold.MoldId:D2} PASS: Anchor={candidate.CornerName}, Score={candidate.Score:F3}, Cover={candidate.Coverage:P1}, D={candidate.TrimmedDistance:F3}, Place=({candidate.Placement.X:F3},{candidate.Placement.Y:F3}).");
                }

                if (best is null || candidate.Score < best.Score)
                {
                    best = candidate;
                    if (isEdgePartial && candidate.Coverage >= 0.995 && candidate.TrimmedDistance <= 0.05)
                    {
                        break;
                    }
                }
            }

            if (isEdgePartial)
            {
                AppendEdgePartialDebugLog($"  ScannedMolds={scannedMoldCount}/{molds.Count} after prefilter.");
            }

            if (isEdgePartial)
            {
                AppendEdgePartialDebugLog(best is null
                    ? $"End {hole.HoleType}: NO MATCH. LogPath={EdgePartialDebugLogPath}"
                    : $"End {hole.HoleType}: BEST M{best.MoldId:D2}, Anchor={best.CornerName}, Score={best.Score:F3}, Cover={best.Coverage:P1}, D={best.TrimmedDistance:F3}, Place=({best.Placement.X:F3},{best.Placement.Y:F3}). LogPath={EdgePartialDebugLogPath}");
            }

            return best;
        }

        private static bool CanAttemptPartialMatch(HoleFeature hole, HoleFeature mold)
        {
            var isEdgePartial = hole.HoleType.StartsWith("EdgePartial:", StringComparison.Ordinal);
            return GetPartialMatchRejectReason(hole, mold, moldPointCount: 8, isEdgePartial) is null;
        }

        private static string? GetPartialMatchRejectReason(HoleFeature hole, HoleFeature mold, int moldPointCount, bool isEdgePartial)
        {
            var minMoldPointCount = isEdgePartial ? 3 : 8;
            if (moldPointCount < minMoldPointCount)
            {
                return $"mold points too few ({moldPointCount}<{minMoldPointCount})";
            }

            if (!IsShapeFamilyCompatible(hole, mold))
            {
                return $"shape family incompatible: hole={hole.HoleType}, mold={mold.HoleType}";
            }

            const double sizeTolerance = 0.01;
            if (hole.Width > mold.Width + sizeTolerance || hole.Height > mold.Height + sizeTolerance)
            {
                return $"size too large: hole=({hole.Width:F3},{hole.Height:F3}), mold=({mold.Width:F3},{mold.Height:F3}), tol={sizeTolerance:F1}";
            }

            if (hole.Area > mold.Area * 1.001)
            {
                return $"area too large: hole={hole.Area:F3}, moldLimit={mold.Area * 1.1:F3}";
            }

            return null;
        }

        private static string? GetEdgePartialPrefilterRejectReason(HoleFeature hole, RectBounds holeBounds, HoleFeature mold, RectBounds moldBounds)
        {
            var holeLong = Math.Max(holeBounds.Width, holeBounds.Height);
            var holeShort = Math.Max(Math.Min(holeBounds.Width, holeBounds.Height), 1e-6);
            var moldLong = Math.Max(moldBounds.Width, moldBounds.Height);
            var moldShort = Math.Max(Math.Min(moldBounds.Width, moldBounds.Height), 1e-6);
            if (holeLong > moldLong + 0.01 || holeShort > moldShort + 0.01)
            {
                return $"bbox long/short too large: hole=({holeLong:F3},{holeShort:F3}), mold=({moldLong:F3},{moldShort:F3})";
            }

            if (hole.Perimeter > mold.Perimeter * 1.001)
            {
                return $"perimeter too large: hole={hole.Perimeter:F3}, moldLimit={mold.Perimeter * 1.35:F3}";
            }

            return null;
        }

        private static PartialMatchCandidate? TryPartialBBoxContourMatch(
            HoleFeature hole,
            IReadOnlyList<(double X, double Y)> holePoints,
            IReadOnlyList<PartialMatchAnchor> holeAnchors,
            PartialMoldCache moldCache,
            bool enableDebugLog)
        {
            var mold = moldCache.Mold;
            var moldPoints = moldCache.Points;
            var moldAnchors = moldCache.Anchors;
            if (enableDebugLog)
            {
                AppendEdgePartialDebugLog($"  Mold M{mold.MoldId:D2} scan: HoleAnchors={holeAnchors.Count}, MoldAnchors={moldAnchors.Count}, AnchorPairs={holeAnchors.Count * moldAnchors.Count}.");
            }

            PartialMatchCandidate? best = null;
            string? bestFailedAnchor = null;
            var bestFailedCoverage = 0.0;
            var bestFailedDistance = double.PositiveInfinity;
            var bestFailedScore = double.PositiveInfinity;
            foreach (var holeAnchor in holeAnchors)
            {
                foreach (var moldAnchor in moldAnchors)
                {
                    var placement = (X: holeAnchor.Point.X - moldAnchor.Point.X, Y: holeAnchor.Point.Y - moldAnchor.Point.Y);
                    var refined = RefinePartialPlacement(holePoints, moldPoints, placement, out var coverage, out var trimmedDistance);
                    var score = trimmedDistance + (1.0 - coverage) * 5.0;
                    var anchorName = $"{holeAnchor.Name}->{moldAnchor.Name}";
                    if (coverage < 0.95 || trimmedDistance > 0.5)
                    {
                        if (score < bestFailedScore)
                        {
                            bestFailedScore = score;
                            bestFailedAnchor = anchorName;
                            bestFailedCoverage = coverage;
                            bestFailedDistance = trimmedDistance;
                        }
                        continue;
                    }

                    var widthRatio = hole.Width / Math.Max(mold.Feature.Width, 1e-6);
                    var heightRatio = hole.Height / Math.Max(mold.Feature.Height, 1e-6);
                    var candidate = new PartialMatchCandidate(mold.MoldId, anchorName, refined, score, coverage, trimmedDistance, widthRatio, heightRatio);
                    if (best is null || candidate.Score < best.Score)
                    {
                        best = candidate;
                    }
                }
            }

            if (enableDebugLog && best is null)
            {
                AppendEdgePartialDebugLog(bestFailedAnchor is null
                    ? $"    Best failed anchor: none evaluated."
                    : $"    Best failed anchor: {bestFailedAnchor}, Score={bestFailedScore:F3}, Cover={bestFailedCoverage:P1}, D={bestFailedDistance:F3}.");
            }

            // 轮廓对齐精度：残孔必须轮廓完全一致
            if (best is not null && best.TrimmedDistance > 0.02)
            {
                if (enableDebugLog)
                {
                    AppendEdgePartialDebugLog($"  REJECT: 轮廓距离超差 D={best.TrimmedDistance:F3} (需≤0.02)");
                }
                return null;
            }

            return best;
        }

        private static (double X, double Y) RefinePartialPlacement(
            IReadOnlyList<(double X, double Y)> holePoints,
            IReadOnlyList<(double X, double Y)> moldPoints,
            (double X, double Y) initialPlacement,
            out double bestCoverage,
            out double bestTrimmedDistance)
        {
            var bestPlacement = initialPlacement;
            var bestScore = double.PositiveInfinity;
            bestCoverage = 0.0;
            bestTrimmedDistance = double.PositiveInfinity;

            var offsets = new (double X, double Y)[]
            {
                (0, 0),
                (-2, 0), (2, 0), (0, -2), (0, 2),
                (-2, -2), (-2, 2), (2, -2), (2, 2),
                (-1, 0), (1, 0), (0, -1), (0, 1)
            };

            foreach (var offset in offsets)
            {
                var placement = (initialPlacement.X + offset.X, initialPlacement.Y + offset.Y);
                var trimmed = TrimmedDirectedChamfer(holePoints, moldPoints, placement, distanceTolerance: 1.5, trimRatio: 0.8, out var coverage);
                var score = trimmed + (1.0 - coverage) * 5.0;
                if (score < bestScore)
                {
                    bestScore = score;
                    bestPlacement = placement;
                    bestCoverage = coverage;
                    bestTrimmedDistance = trimmed;
                }
            }

            return bestPlacement;
        }

        private static IReadOnlyList<PartialMatchAnchor> BuildPartialMatchAnchors(IReadOnlyList<(double X, double Y)> points, string prefix, int maxAnchors)
        {
            if (points.Count == 0)
            {
                return [];
            }

            var raw = new List<PartialMatchAnchor>();
            void Add(string name, (double X, double Y) point)
            {
                if (raw.Any(a => Math.Abs(a.Point.X - point.X) < 1e-6 && Math.Abs(a.Point.Y - point.Y) < 1e-6))
                {
                    return;
                }

                raw.Add(new PartialMatchAnchor(name, point));
            }

            Add($"{prefix}起点", points[0]);
            if (points.Count > 1)
            {
                Add($"{prefix}终点", points[^1]);
            }

            var box = BoundsOf(points);
            var leftBottom = points.OrderBy(p => p.X + p.Y).First();
            var leftTop = points.OrderBy(p => p.X - p.Y).First();
            var rightBottom = points.OrderByDescending(p => p.X - p.Y).First();
            var rightTop = points.OrderByDescending(p => p.X + p.Y).First();
            var left = points.OrderBy(p => p.X).ThenBy(p => p.Y).First();
            var right = points.OrderByDescending(p => p.X).ThenByDescending(p => p.Y).First();
            var bottom = points.OrderBy(p => p.Y).ThenBy(p => p.X).First();
            var top = points.OrderByDescending(p => p.Y).ThenByDescending(p => p.X).First();
            var center = ((box.MinX + box.MaxX) * 0.5, (box.MinY + box.MaxY) * 0.5);

            Add($"{prefix}左下", leftBottom);
            Add($"{prefix}左上", leftTop);
            Add($"{prefix}右下", rightBottom);
            Add($"{prefix}右上", rightTop);
            Add($"{prefix}最左", left);
            Add($"{prefix}最右", right);
            Add($"{prefix}最下", bottom);
            Add($"{prefix}最上", top);
            Add($"{prefix}中心", center);

            var diagonal = Math.Sqrt(box.Width * box.Width + box.Height * box.Height);
            var minSegmentLength = Math.Max(diagonal * 0.003, 0.2);
            var minTurnSin = Math.Sin(Math.PI / 180.0 * 12.0);
            for (var i = 1; i < points.Count - 1; i++)
            {
                var prev = points[i - 1];
                var current = points[i];
                var next = points[i + 1];
                var vx1 = current.X - prev.X;
                var vy1 = current.Y - prev.Y;
                var vx2 = next.X - current.X;
                var vy2 = next.Y - current.Y;
                var len1 = Math.Sqrt(vx1 * vx1 + vy1 * vy1);
                var len2 = Math.Sqrt(vx2 * vx2 + vy2 * vy2);
                if (len1 < minSegmentLength || len2 < minSegmentLength)
                {
                    continue;
                }

                var turnSin = Math.Abs(vx1 * vy2 - vy1 * vx2) / Math.Max(len1 * len2, 1e-9);
                if (turnSin >= minTurnSin)
                {
                    Add($"{prefix}拐角{i}", current);
                }
            }

            if (raw.Count <= maxAnchors)
            {
                return raw;
            }

            var prioritized = raw
                .OrderBy(a =>
                {
                    var p = a.Point;
                    var dx = Math.Min(p.X - box.MinX, box.MaxX - p.X);
                    var dy = Math.Min(p.Y - box.MinY, box.MaxY - p.Y);
                    var edgeScore = Math.Min(dx, dy);
                    return edgeScore;
                })
                .ThenBy(a => a.Name)
                .Take(maxAnchors)
                .ToList();

            return prioritized;
        }

        private static double TrimmedDirectedChamfer(
            IReadOnlyList<(double X, double Y)> sourcePoints,
            IReadOnlyList<(double X, double Y)> targetPoints,
            (double X, double Y) targetPlacement,
            double distanceTolerance,
            double trimRatio,
            out double coverage)
        {
            var distances = new double[sourcePoints.Count];
            var covered = 0;
            var targetClosed = IsClosedPointChain(targetPoints);
            var segmentCount = targetClosed ? targetPoints.Count : Math.Max(targetPoints.Count - 1, 0);
            for (var i = 0; i < sourcePoints.Count; i++)
            {
                var hp = sourcePoints[i];
                var best = double.PositiveInfinity;
                for (var j = 0; j < segmentCount; j++)
                {
                    var a = targetPoints[j];
                    var b = targetPoints[(j + 1) % targetPoints.Count];
                    var shiftedA = (X: a.X + targetPlacement.X, Y: a.Y + targetPlacement.Y);
                    var shiftedB = (X: b.X + targetPlacement.X, Y: b.Y + targetPlacement.Y);
                    var dist = PointToSegmentDistance(hp, shiftedA, shiftedB);
                    if (dist < best)
                    {
                        best = dist;
                    }
                }

                distances[i] = best;
                if (best <= distanceTolerance)
                {
                    covered++;
                }
            }

            coverage = sourcePoints.Count == 0 ? 0.0 : (double)covered / sourcePoints.Count;
            Array.Sort(distances);
            var keep = Compat.Clamp((int)Math.Round(distances.Length * trimRatio), 1, distances.Length);
            var sum = 0.0;
            for (var i = 0; i < keep; i++)
            {
                sum += distances[i];
            }
            return sum / keep;
        }

        private static bool IsClosedPointChain(IReadOnlyList<(double X, double Y)> points)
        {
            if (points.Count < 3)
            {
                return false;
            }

            var first = points[0];
            var last = points[^1];
            var closeDistance = Math.Sqrt((first.X - last.X) * (first.X - last.X) + (first.Y - last.Y) * (first.Y - last.Y));
            if (closeDistance <= 1e-6)
            {
                return true;
            }

            var box = BoundsOf(points);
            var diagonal = Math.Sqrt(box.Width * box.Width + box.Height * box.Height);
            return closeDistance <= Math.Max(diagonal * 0.001, 0.1);
        }

        private static double PointToSegmentDistance((double X, double Y) p, (double X, double Y) a, (double X, double Y) b)
        {
            var vx = b.X - a.X;
            var vy = b.Y - a.Y;
            var wx = p.X - a.X;
            var wy = p.Y - a.Y;
            var len2 = vx * vx + vy * vy;
            if (len2 <= 1e-12)
            {
                var dx = p.X - a.X;
                var dy = p.Y - a.Y;
                return Math.Sqrt(dx * dx + dy * dy);
            }

            var t = Compat.Clamp((wx * vx + wy * vy) / len2, 0.0, 1.0);
            var projX = a.X + t * vx;
            var projY = a.Y + t * vy;
            var pdx = p.X - projX;
            var pdy = p.Y - projY;
            return Math.Sqrt(pdx * pdx + pdy * pdy);
        }

        private static IReadOnlyList<(double X, double Y)> NormalizeFeaturePoints(HoleFeature feature)
        {
            if (feature.Points is { Count: >= 2 })
            {
                return feature.Points;
            }

            if (IsCircleLike(feature))
            {
                var radius = Math.Max(Math.Min(feature.Width, feature.Height) * 0.5, 0.5);
                var points = new List<(double X, double Y)>(36);
                for (var i = 0; i < 36; i++)
                {
                    var angle = 2.0 * Math.PI * i / 36.0;
                    points.Add((feature.Centroid.X + radius * Math.Cos(angle), feature.Centroid.Y + radius * Math.Sin(angle)));
                }
                return points;
            }

            return BBoxPoints(feature.Centroid, feature.Width, feature.Height);
        }

        private static IReadOnlyList<(double X, double Y)> NormalizeMoldPoints(MoldProfile mold)
        {
            if (mold.OutlinePoints is { Count: >= 2 })
            {
                return mold.OutlinePoints;
            }

            return NormalizeFeaturePoints(mold.Feature);
        }

        private static IReadOnlyList<(double X, double Y)> BBoxPoints((double X, double Y) center, double width, double height)
        {
            var minX = center.X - width * 0.5;
            var maxX = center.X + width * 0.5;
            var minY = center.Y - height * 0.5;
            var maxY = center.Y + height * 0.5;
            return [(minX, minY), (minX, maxY), (maxX, maxY), (maxX, minY)];
        }

        private static RectBounds BoundsOf(IReadOnlyList<(double X, double Y)> points)
        {
            return new RectBounds(points.Min(p => p.X), points.Min(p => p.Y), points.Max(p => p.X), points.Max(p => p.Y));
        }

        private static bool IsSameShapeType(HoleFeature hole, HoleFeature mold)
        {
            var hCircle = IsCircleLike(hole);
            var mCircle = IsCircleLike(mold);
            if (hCircle || mCircle)
            {
                return hCircle == mCircle;
            }

            var hPoly = hole.HoleType.ContainsIgnoreCase("Polyline") ||
                        hole.HoleType.ContainsIgnoreCase("EntityComposite") ||
                        hole.HoleType.ContainsIgnoreCase("MixedArcLine") ||
                        hole.HoleType.ContainsIgnoreCase("EdgePartial");
            var mPoly = mold.HoleType.ContainsIgnoreCase("Polyline") ||
                        mold.HoleType.ContainsIgnoreCase("EntityComposite") ||
                        mold.HoleType.ContainsIgnoreCase("MixedArcLine") ||
                        mold.HoleType.ContainsIgnoreCase("EdgePartial");
            if (hPoly || mPoly)
            {
                return hPoly == mPoly;
            }

            return true;
        }

        private static bool IsShapeFamilyCompatible(HoleFeature hole, HoleFeature mold)
        {
            var hCircle = IsCircleLike(hole);
            var mCircle = IsCircleLike(mold);
            if (hCircle || mCircle)
            {
                // 圆孔只允许圆族模具；但允许“圆形EntityComposite”进入圆族，避免圆孔全丢。
                return hCircle && mCircle;
            }

            var hPolyFamily = hole.HoleType.ContainsIgnoreCase("Polyline")
                              || hole.HoleType.ContainsIgnoreCase("EntityComposite")
                              || hole.HoleType.ContainsIgnoreCase("MixedArcLine")
                              || hole.HoleType.ContainsIgnoreCase("EdgePartial");
            var mPolyFamily = mold.HoleType.ContainsIgnoreCase("Polyline")
                              || mold.HoleType.ContainsIgnoreCase("EntityComposite")
                              || mold.HoleType.ContainsIgnoreCase("MixedArcLine")
                              || mold.HoleType.ContainsIgnoreCase("EdgePartial");
            if (hPolyFamily || mPolyFamily)
            {
                return hPolyFamily && mPolyFamily;
            }

            return true;
        }

        public  static bool IsCircleLike(HoleFeature f)
        {
            if (f.HoleType.ContainsIgnoreCase("Circle"))
            {
                return true;
            }

            // 先判“椭圆/腰孔”再判圆孔：
            // 对近圆但存在明显长短轴差异的形状（如 19*18）优先按非圆处理，避免被 φ20 抢分类。
            var longSide = Math.Max(f.Width, f.Height);
            var shortSide = Math.Max(Math.Min(f.Width, f.Height), 1e-6);
            var axisRatio = longSide / shortSide;
            // 放宽圆孔几何门限：实图中圆孔经抽样/离散后常出现 3%~7% 的轴向误差。
            // 仍保留上限，避免明显腰孔/椭圆误入圆孔族。
            if (axisRatio >= 1.08)
            {
                return false;
            }

            // 对 EntityComposite / 其他类型做几何判定：允许更小的宽高误差容忍。
            var maxWh = Math.Max(longSide, 1e-6);
            var whRatio = Math.Abs(f.Width - f.Height) / maxWh;
            if (whRatio > 0.08)
            {
                return false;
            }

            var circularity = 4.0 * Math.PI * f.Area / Math.Max(f.Perimeter * f.Perimeter, 1e-6);
            return circularity >= 0.76;
        }

        private static double SimilarityScore(HoleFeature h, HoleFeature m)
        {
            var dw = Math.Abs(h.Width - m.Width) / Math.Max(h.Width, 1e-6);
            var dh = Math.Abs(h.Height - m.Height) / Math.Max(h.Height, 1e-6);
            var da = Math.Abs(h.Area - m.Area) / Math.Max(h.Area, 1e-6);
            var dp = Math.Abs(h.Perimeter - m.Perimeter) / Math.Max(h.Perimeter, 1e-6);
            var dr = Math.Abs((h.Width / Math.Max(h.Height, 1e-6)) - (m.Width / Math.Max(m.Height, 1e-6)));
            var ds = SignatureDistance(h.Signature, m.Signature);
            return 0.15 * dw + 0.15 * dh + 0.2 * da + 0.15 * dp + 0.1 * dr + 0.25 * ds;
        }

        private static double SignatureDistance(double[] a, double[] b)
        {
            if (a.Length == 0 || b.Length == 0)
            {
                return 1.0;
            }
            var n = Math.Min(a.Length, b.Length);
            if (a.Length != n)
            {
                a = Resample(a, n);
            }
            if (b.Length != n)
            {
                b = Resample(b, n);
            }

            var forward = MinCyclicRmse(a, b);
            var mirroredSource = b.ToArray();
            Array.Reverse(mirroredSource);
            var mirrored = MinCyclicRmse(a, mirroredSource);
            return Math.Min(forward, mirrored);
        }

        private static double MinCyclicRmse(double[] a, double[] b)
        {
            var n = a.Length;
            var best = double.MaxValue;
            for (var shift = 0; shift < n; shift++)
            {
                double sum = 0;
                for (var i = 0; i < n; i++)
                {
                    var j = (i + shift) % n;
                    var d = a[i] - b[j];
                    sum += d * d;
                }
                var rmse = Math.Sqrt(sum / n);
                if (rmse < best)
                {
                    best = rmse;
                }
            }
            return best;
        }

        private static double[] Resample(double[] source, int n)
        {
            if (source.Length == n)
            {
                return source;
            }
            var result = new double[n];
            for (var i = 0; i < n; i++)
            {
                var idx = (double)i * source.Length / n;
                var i0 = (int)Math.Floor(idx) % source.Length;
                var i1 = (i0 + 1) % source.Length;
                var t = idx - Math.Floor(idx);
                result[i] = source[i0] * (1 - t) + source[i1] * t;
            }
            return result;
        }
    }

    public sealed class PlcRegisterRow : INotifyPropertyChanged
    {
        private string _address;
        private string _value;
        private string _info;

        public PlcRegisterRow(string address, string value, string info)
        {
            _address = address;
            _value = value;
            _info = info;
        }

        public string Address
        {
            get => _address;
            set
            {
                if (_address == value) return;
                _address = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Address)));
            }
        }

        public string Value
        {
            get => _value;
            set
            {
                if (_value == value) return;
                _value = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Value)));
            }
        }

        public string Info
        {
            get => _info;
            set
            {
                if (_info == value) return;
                _info = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Info)));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }
}

