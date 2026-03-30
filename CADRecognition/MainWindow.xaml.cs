using Microsoft.Win32;
using netDxf;
using netDxf.Entities;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;
using WinOpenFileDialog = Microsoft.Win32.OpenFileDialog;
using WpfBrushes = System.Windows.Media.Brushes;
using WpfColor = System.Windows.Media.Color;
using WpfEllipse = System.Windows.Shapes.Ellipse;
using WpfLine = System.Windows.Shapes.Line;

namespace CADRecognition
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private readonly IDxfPreviewPlugin _previewPlugin = new BasicCanvasPreviewPlugin();
        private readonly Dictionary<string, DxfDocument> _documentCache = [];
        private readonly Dictionary<string, ImageSource> _moldPreviewCache = [];
        private readonly InteractiveDxfPreview _viewer = new();
        private readonly ObservableCollection<MoldRow> _moldRows = [];
        private readonly ObservableCollection<PositionRow> _positionRows = [];
        private readonly List<string> _moldFiles = [];
        private MatchResult? _lastMatchResult;
        private ProjectProfile? _lastProjectProfile;
        private List<MoldProfile> _lastMolds = [];

        private string? _projectFile;
        private DxfDocument? _projectDoc;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            MoldCountText.Text = "0";
            ProjectFileText.Text = "未加载";
            PreviewHost.Content = _viewer;
            FileTreeView.Items.Clear();
        }

        public ObservableCollection<MoldRow> MoldRows => _moldRows;
        public ObservableCollection<PositionRow> PositionRows => _positionRows;

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        private void ImportProjectDxf_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new WinOpenFileDialog
            {
                Filter = "DXF 文件 (*.dxf)|*.dxf",
                Multiselect = false
            };
            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            _projectFile = dialog.FileName;
            _projectDoc = DxfDocument.Load(_projectFile);
            _documentCache[_projectFile] = _projectDoc;
            ProjectFileText.Text = System.IO.Path.GetFileName(_projectFile);
            RefreshFileList();
            RenderPreview(_projectDoc, _projectFile, withAnnotation: _lastMatchResult is not null);
            StatusText.Text = "工程 DXF 已加载。";
        }

        private void ImportMoldsDxf_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new WinOpenFileDialog
            {
                Filter = "DXF 文件 (*.dxf)|*.dxf",
                Multiselect = true
            };
            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            _moldFiles.Clear();
            _moldFiles.AddRange(dialog.FileNames);
            foreach (var file in _moldFiles)
            {
                _documentCache[file] = DxfDocument.Load(file);
            }
            MoldCountText.Text = _moldFiles.Count.ToString(CultureInfo.InvariantCulture);
            RefreshFileList();
            StatusText.Text = $"已导入 {_moldFiles.Count} 张模具 DXF。";
        }

        private void Recognize_Click(object sender, RoutedEventArgs e)
        {
            if (_projectDoc is null || string.IsNullOrWhiteSpace(_projectFile))
            {
                StatusText.Text = "请先导入工程 DXF。";
                return;
            }
            if (_moldFiles.Count == 0)
            {
                StatusText.Text = "请先导入模具 DXF。";
                return;
            }

            var project = DxfAnalyzer.ExtractProject(_projectDoc);
            _lastProjectProfile = project;
            var molds = _moldFiles
                .Select((f, idx) => DxfAnalyzer.ExtractMold(idx + 1, f))
                .ToList();
            _lastMolds = molds;

            var matcher = new MoldMatcher();
            var result = matcher.Match(project, molds);
            _lastMatchResult = result;
            RenderResult(result, molds);
            RenderPreview(_projectDoc, _projectFile, withAnnotation: true);
            StatusText.Text = $"识别完成：外轮廓 {project.OuterRectangle.Width:F2} x {project.OuterRectangle.Height:F2}，孔洞 {result.HoleAssignments.Count} 个。";
        }

        private void RenderResult(MatchResult result, IReadOnlyList<MoldProfile> molds)
        {
            _moldRows.Clear();
            _positionRows.Clear();
            RenderLegend(molds);

            var useCounter = result.HoleAssignments
                .Where(x => !x.Hole.HoleType.StartsWith("EdgeNotch:", StringComparison.Ordinal))
                .GroupBy(x => x.MoldId)
                .ToDictionary(g => g.Key, g => g.Count());

            foreach (var mold in molds.OrderBy(x => x.MoldId))
            {
                useCounter.TryGetValue(mold.MoldId, out var count);
                _moldRows.Add(new MoldRow
                {
                    MoldPreview = BuildMoldPreview(mold.FilePath),
                    MoldCode = $"M{mold.MoldId:D2}",
                    MoldName = System.IO.Path.GetFileNameWithoutExtension(mold.FilePath),
                    UsedCount = count,
                    MatchType = mold.MoldId == 1 ? "角落连续冲压" : "单次冲压",
                    Remark = mold.MoldId == 1 ? "仅四角孔洞" : "普通孔洞"
                });
            }

            var i = 1;
            foreach (var row in result.HoleAssignments
                         .Where(x => !x.Hole.HoleType.StartsWith("EdgeNotch:", StringComparison.Ordinal))
                         .OrderBy(x => x.Hole.Centroid.Y).ThenBy(x => x.Hole.Centroid.X))
            {
                _positionRows.Add(new PositionRow
                {
                    Index = i++,
                    HoleType = row.Hole.HoleType,
                    MoldId = row.MoldId,
                    MoldCode = $"M{row.MoldId:D2}",
                    PosX = Math.Round(row.Hole.Centroid.X, 3),
                    PosY = Math.Round(row.Hole.Centroid.Y, 3),
                    PositionRelation = row.PositionRelation,
                    IsCornerCandidate = row.IsCornerCandidate ? "是" : "否",
                    IsEdgeHole = row.IsEdgeHole ? "是" : "否",
                    TopCandidates = row.TopCandidates
                });
            }
        }

        private void RenderLegend(IReadOnlyList<MoldProfile> molds)
        {
            LegendPanel.Children.Clear();
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
                panel.Children.Add(new TextBlock
                {
                    Text = $"M{mold.MoldId:D2}",
                    Foreground = WpfBrushes.White,
                    FontSize = 11,
                    VerticalAlignment = VerticalAlignment.Center
                });
                chip.Child = panel;
                LegendPanel.Children.Add(chip);
            }
        }

        private void RefreshFileList()
        {
            FileTreeView.Items.Clear();

            var root = new TreeViewItem
            {
                Header = "图纸列表",
                IsExpanded = true
            };

            var projectNode = new TreeViewItem
            {
                Header = "工程图",
                IsExpanded = true
            };
            if (!string.IsNullOrWhiteSpace(_projectFile))
            {
                projectNode.Items.Add(new TreeViewItem
                {
                    Header = System.IO.Path.GetFileName(_projectFile),
                    Tag = _projectFile
                });
            }

            var moldNode = new TreeViewItem
            {
                Header = "模具库",
                IsExpanded = true
            };
            foreach (var file in _moldFiles)
            {
                moldNode.Items.Add(new TreeViewItem
                {
                    Header = System.IO.Path.GetFileName(file),
                    Tag = file
                });
            }
            root.Items.Add(projectNode);
            root.Items.Add(moldNode);
            FileTreeView.Items.Add(root);
        }

        private void RenderPreview(DxfDocument doc, string? path, bool withAnnotation)
        {
            _previewPlugin.CreatePreview(doc, _viewer);
            if (withAnnotation && !string.IsNullOrWhiteSpace(path) && path == _projectFile && _lastMatchResult is not null)
            {
                if (_lastProjectProfile is not null)
                {
                    _viewer.RenderCornerZones(_lastProjectProfile.OuterRectangle, 0.22);
                }
                _viewer.RenderAnnotations(_lastMatchResult.HoleAssignments, _lastMolds);
            }
            else
            {
                _viewer.RenderCornerZones(null, 0);
                _viewer.RenderAnnotations([], []);
            }
            PreviewHintText.Visibility = Visibility.Collapsed;
        }

        private void FileTreeView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (FileTreeView.SelectedItem is not TreeViewItem item || item.Tag is not string path)
            {
                return;
            }
            if (!_documentCache.TryGetValue(path, out var doc))
            {
                return;
            }
            var showAnnotation = _lastMatchResult is not null && path == _projectFile;
            RenderPreview(doc, path, showAnnotation);
            StatusText.Text = $"预览图纸：{System.IO.Path.GetFileName(path)}";
        }

        private void PositionGrid_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (PositionGrid.SelectedItem is not PositionRow row)
            {
                return;
            }
            _viewer.FocusHole(row.PosX, row.PosY, row.MoldId);
            StatusText.Text = $"已定位孔位 #{row.Index}（{row.MoldCode}），角候选={row.IsCornerCandidate}，边缘孔={row.IsEdgeHole}，Top3={row.TopCandidates}";
        }

        private async void PositionGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (PositionGrid.SelectedItem is not PositionRow row)
            {
                return;
            }
            _viewer.FocusHole(row.PosX, row.PosY, row.MoldId, targetZoom: 4.0);
            await _viewer.BlinkFocusAsync(row.PosX, row.PosY, row.MoldId);
            StatusText.Text = $"已放大定位孔位 #{row.Index}（{row.MoldCode}），角候选={row.IsCornerCandidate}，边缘孔={row.IsEdgeHole}，Top3={row.TopCandidates}";
        }

        private ImageSource? BuildMoldPreview(string path)
        {
            if (_moldPreviewCache.TryGetValue(path, out var cached))
            {
                return cached;
            }

            if (!_documentCache.TryGetValue(path, out var doc))
            {
                if (!System.IO.File.Exists(path))
                {
                    return null;
                }
                doc = DxfDocument.Load(path);
                _documentCache[path] = doc;
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

                foreach (var c in doc.Entities.Circles)
                {
                    var center = Map(c.Center.X, c.Center.Y);
                    dc.DrawEllipse(null, circlePen, center, c.Radius * scale, c.Radius * scale);
                }

                foreach (var p in doc.Entities.Polylines2D)
                {
                    var pts = DxfAnalyzer.ExpandPolyline2D(p, 12);
                    if (pts.Count < 2)
                    {
                        continue;
                    }
                    var geo = new StreamGeometry();
                    using var g = geo.Open();
                    g.BeginFigure(Map(pts[0].X, pts[0].Y), false, p.IsClosed);
                    g.PolyLineTo(pts.Skip(1).Select(t => Map(t.X, t.Y)).ToList(), true, false);
                    geo.Freeze();
                    dc.DrawGeometry(null, polyPen, geo);
                }

                foreach (var a in doc.Entities.Arcs)
                {
                    var pts = DxfAnalyzer.SampleArc(a, 18);
                    if (pts.Count < 2)
                    {
                        continue;
                    }
                    var geo = new StreamGeometry();
                    using var g = geo.Open();
                    g.BeginFigure(Map(pts[0].X, pts[0].Y), false, false);
                    g.PolyLineTo(pts.Skip(1).Select(t => Map(t.X, t.Y)).ToList(), true, false);
                    geo.Freeze();
                    dc.DrawGeometry(null, arcPen, geo);
                }
            }

            var bmp = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(dv);
            bmp.Freeze();
            _moldPreviewCache[path] = bmp;
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

            foreach (var line in document.Entities.Lines)
            {
                canvas.Children.Add(new WpfLine
                {
                    X1 = (line.StartPoint.X - bounds.MinX) * scale + margin,
                    Y1 = viewHeight - ((line.StartPoint.Y - bounds.MinY) * scale + margin),
                    X2 = (line.EndPoint.X - bounds.MinX) * scale + margin,
                    Y2 = viewHeight - ((line.EndPoint.Y - bounds.MinY) * scale + margin),
                    Stroke = WpfBrushes.LimeGreen,
                    StrokeThickness = 1
                });
            }

            foreach (var circle in document.Entities.Circles)
            {
                var r = circle.Radius * scale;
                var x = (circle.Center.X - bounds.MinX) * scale + margin - r;
                var y = viewHeight - ((circle.Center.Y - bounds.MinY) * scale + margin) - r;
                var el = new WpfEllipse
                {
                    Width = r * 2,
                    Height = r * 2,
                    Stroke = WpfBrushes.Gold,
                    StrokeThickness = 1
                };
                Canvas.SetLeft(el, x);
                Canvas.SetTop(el, y);
                canvas.Children.Add(el);
            }

            foreach (var poly in document.Entities.Polylines2D.Where(p => p.IsClosed))
            {
                var sampled = DxfAnalyzer.ExpandPolyline2D(poly, 24);
                var points = new PointCollection(sampled.Select(v =>
                    new System.Windows.Point(
                        (v.X - bounds.MinX) * scale + margin,
                        viewHeight - ((v.Y - bounds.MinY) * scale + margin))));
                canvas.Children.Add(new Polygon
                {
                    Points = points,
                    Stroke = WpfBrushes.DeepSkyBlue,
                    StrokeThickness = 1,
                    Fill = WpfBrushes.Transparent
                });
            }

            foreach (var arc in document.Entities.Arcs)
            {
                var sampled = DxfAnalyzer.SampleArc(arc, 32);
                var points = new PointCollection(sampled.Select(p =>
                    new System.Windows.Point(
                        (p.X - bounds.MinX) * scale + margin,
                        viewHeight - ((p.Y - bounds.MinY) * scale + margin))));
                canvas.Children.Add(new Polyline
                {
                    Points = points,
                    Stroke = WpfBrushes.Orange,
                    StrokeThickness = 1.2
                });
            }

            viewer.LoadScene(canvas, bounds, viewWidth, viewHeight, scale, margin);
        }
    }

    public sealed class InteractiveDxfPreview : Border
    {
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

        public void RenderAnnotations(IReadOnlyList<HoleAssignment> assignments, IReadOnlyList<MoldProfile> molds)
        {
            _markCanvas.Children.Clear();
            _focusRing = null;
            var moldMap = molds.ToDictionary(m => m.MoldId, m => m);

            foreach (var ass in assignments)
            {
                var c = ass.Hole.Centroid;
                var p = ModelToCanvas(c.X, c.Y);
                var color = GetMoldColor(ass.MoldId);
                var brush = new SolidColorBrush(color);

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
                        StrokeThickness = 1.6
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
                        StrokeThickness = 1
                    };
                    Canvas.SetLeft(mark, p.X - 6);
                    Canvas.SetTop(mark, p.Y - 6);
                    _markCanvas.Children.Add(mark);
                }

                var text = new TextBlock
                {
                    Text = $"M{ass.MoldId:D2}",
                    Foreground = brush,
                    FontWeight = FontWeights.Bold,
                    FontSize = 11
                };
                Canvas.SetLeft(text, p.X + 8);
                Canvas.SetTop(text, p.Y - 8);
                _markCanvas.Children.Add(text);
            }
        }

        public void RenderCornerZones(RectBounds? rect, double ratio)
        {
            _zoneCanvas.Children.Clear();
            if (rect is null || ratio <= 0)
            {
                return;
            }

            var zoneW = rect.Width * ratio;
            var zoneH = rect.Height * ratio;
            var zones = new[]
            {
                (rect.MinX, rect.MinY),
                (rect.MinX, rect.MaxY - zoneH),
                (rect.MaxX - zoneW, rect.MinY),
                (rect.MaxX - zoneW, rect.MaxY - zoneH)
            };

            foreach (var z in zones)
            {
                var p1 = ModelToCanvas(z.Item1, z.Item2);
                var p2 = ModelToCanvas(z.Item1 + zoneW, z.Item2 + zoneH);
                var left = Math.Min(p1.X, p2.X);
                var top = Math.Min(p1.Y, p2.Y);
                var w = Math.Abs(p2.X - p1.X);
                var h = Math.Abs(p2.Y - p1.Y);

                var box = new Border
                {
                    Width = w,
                    Height = h,
                    Background = new SolidColorBrush(WpfColor.FromArgb(35, 255, 152, 0)),
                    BorderBrush = new SolidColorBrush(WpfColor.FromArgb(160, 255, 152, 0)),
                    BorderThickness = new Thickness(1)
                };
                Canvas.SetLeft(box, left);
                Canvas.SetTop(box, top);
                _zoneCanvas.Children.Add(box);
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

        public void FocusHole(double modelX, double modelY, int moldId, double? targetZoom = null)
        {
            var p = ModelToCanvas(modelX, modelY);
            var color = new SolidColorBrush(GetMoldColor(moldId));
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

            if (targetZoom.HasValue)
            {
                var z = Math.Clamp(targetZoom.Value, 0.2, 30);
                _scale.ScaleX = z;
                _scale.ScaleY = z;
            }
            var scale = _scale.ScaleX;
            var centerX = ActualWidth > 0 ? ActualWidth / 2 : _sceneCanvas.Width / 2;
            var centerY = ActualHeight > 0 ? ActualHeight / 2 : _sceneCanvas.Height / 2;
            _translate.X = centerX - p.X * scale;
            _translate.Y = centerY - p.Y * scale;
        }

        public async Task BlinkFocusAsync(double modelX, double modelY, int moldId)
        {
            if (_focusRing is null)
            {
                FocusHole(modelX, modelY, moldId);
            }
            if (_focusRing is null)
            {
                return;
            }
            var c = new SolidColorBrush(GetMoldColor(moldId));
            _focusRing.Stroke = c;
            for (var i = 0; i < 3; i++)
            {
                _focusRing.Visibility = Visibility.Hidden;
                await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Background);
                await Task.Delay(120);
                _focusRing.Visibility = Visibility.Visible;
                await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Background);
                await Task.Delay(120);
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
            var target = Math.Clamp(old * factor, 0.2, 30);
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

    public static class DxfAnalyzer
    {
        private const int SignatureSamples = 72;

        public static MoldProfile ExtractMold(int moldId, string path)
        {
            var doc = DxfDocument.Load(path);
            var holes = ExtractHoles(doc);
            var outer = DetectOuterRectangle(doc);
            var feature = holes
                .Where(h => h.Area <= Math.Max(outer.Area * 0.75, 10.0))
                .OrderByDescending(h => h.Area)
                .FirstOrDefault();

            if (feature is null)
            {
                feature = BuildFeatureFromEntities(doc);
            }

            feature ??= new HoleFeature("Unknown", (0, 0), 1, 1, 1, 1, 0, CreateCircleSignature(1, SignatureSamples));
            var outline = ExtractMoldOutline(doc);
            return new MoldProfile(moldId, path, feature, outline);
        }

        public static ProjectProfile ExtractProject(DxfDocument doc)
        {
            var holes = ExtractHoles(doc);
            var outer = DetectOuterRectangle(doc);
            holes.AddRange(ExtractCornerMissingFeatures(doc, outer));

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
            return new ProjectProfile(outer, DeduplicateHoles(innerHoles));
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
            points.AddRange(doc.Entities.Arcs.SelectMany(a => SampleArc(a, 24)));

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

        private static List<HoleFeature> ExtractHoles(DxfDocument doc)
        {
            var holes = new List<HoleFeature>();
            var allBounds = GetRawBounds(doc);
            var allArea = Math.Max((allBounds.MaxX - allBounds.MinX) * (allBounds.MaxY - allBounds.MinY), 1.0);

            foreach (var c in doc.Entities.Circles)
            {
                var d = c.Radius * 2;
                var area = Math.PI * c.Radius * c.Radius;
                var perimeter = 2 * Math.PI * c.Radius;
                holes.Add(new HoleFeature(
                    "Circle",
                    (c.Center.X, c.Center.Y),
                    d,
                    d,
                    area,
                    perimeter,
                    0,
                    CreateCircleSignature(c.Radius, SignatureSamples)));
            }

            foreach (var a in doc.Entities.Arcs)
            {
                var sweep = NormalizeArcSweep(a.StartAngle, a.EndAngle);
                if (sweep < 350.0)
                {
                    continue;
                }
                var d = a.Radius * 2;
                var area = Math.PI * a.Radius * a.Radius;
                var perimeter = 2 * Math.PI * a.Radius;
                holes.Add(new HoleFeature(
                    "ArcCircle",
                    (a.Center.X, a.Center.Y),
                    d,
                    d,
                    area,
                    perimeter,
                    0,
                    CreateCircleSignature(a.Radius, SignatureSamples)));
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
                    (pts.Average(p => p.Item1), pts.Average(p => p.Item2)),
                    maxX - minX,
                    maxY - minY,
                    Math.Abs(area),
                    perimeter,
                    pl.Elevation,
                    CreatePolylineSignature(pts, SignatureSamples)));
            }

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
                    (pts.Average(p => p.X), pts.Average(p => p.Y)),
                    width,
                    height,
                    pseudoArea,
                    perimeter,
                    pl.Elevation,
                    CreatePolylineSignature(pts, SignatureSamples)));
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
            var result = new List<HoleFeature>();
            foreach (var h in source.OrderBy(x => x.Area))
            {
                var exists = result.Any(r =>
                    Math.Abs(r.Centroid.X - h.Centroid.X) < 1e-3 &&
                    Math.Abs(r.Centroid.Y - h.Centroid.Y) < 1e-3 &&
                    Math.Abs(r.Width - h.Width) < 1e-3 &&
                    Math.Abs(r.Height - h.Height) < 1e-3);
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
                    sig));
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
                        CreatePolylineSignature(g, SignatureSamples)));
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
                signature);
        }

        private static List<(double X, double Y)> ExtractMoldOutline(DxfDocument doc)
        {
            var pts = CollectGeometryPoints(doc);
            if (pts.Count < 2)
            {
                return [];
            }

            var cx = pts.Average(p => p.X);
            var cy = pts.Average(p => p.Y);
            var ordered = pts
                .Select(p => (X: p.X - cx, Y: p.Y - cy))
                .OrderBy(p => Math.Atan2(p.Y, p.X))
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

        public MatchResult Match(ProjectProfile project, IReadOnlyList<MoldProfile> molds)
        {
            var rows = new List<HoleAssignment>();
            if (molds.Count == 0 || project.Holes.Count == 0)
            {
                return new MatchResult(rows);
            }

            var corners = project.OuterRectangle.Corners;
            var holePool = project.Holes.ToList();
            var mold1 = molds.FirstOrDefault(m => m.MoldId == 1) ?? molds[0];
            var nonCornerMolds = molds.Where(m => m.MoldId != mold1.MoldId).ToList();
            if (nonCornerMolds.Count == 0)
            {
                nonCornerMolds.Add(mold1);
            }

            var diag = Math.Sqrt(project.OuterRectangle.Width * project.OuterRectangle.Width +
                                 project.OuterRectangle.Height * project.OuterRectangle.Height);
            var safeDiag = Math.Max(diag, 1e-6);

            var usedCorner = new HashSet<HoleFeature>();
            foreach (var corner in corners)
            {
                var synthetic = holePool
                    .Where(h => !usedCorner.Contains(h) && h.HoleType == $"CornerMissing:{corner.Name}")
                    .OrderBy(h => Dist2(h.Centroid, corner))
                    .FirstOrDefault();
                if (synthetic is not null)
                {
                    usedCorner.Add(synthetic);
                    rows.Add(new HoleAssignment(
                        synthetic,
                        mold1.MoldId,
                        $"角落连续冲压({corner.Name})",
                        true,
                        true,
                        BuildTopCandidates(synthetic, molds, project.OuterRectangle, 3)));
                    continue;
                }

                var nearest = holePool
                    .Where(h => !usedCorner.Contains(h))
                    .Where(h => IsInCornerMissingZone(h, project.OuterRectangle, corner))
                    .Select(h =>
                    {
                        var m1 = ScoreForHoleAgainstMold(h, mold1.Feature, project.OuterRectangle);
                        var bestOther = nonCornerMolds.Min(m => ScoreForHoleAgainstMold(h, m.Feature, project.OuterRectangle));
                        var distNorm = Math.Sqrt(Dist2(h.Centroid, corner)) / safeDiag;
                        var combined = distNorm * 0.55 + m1 * 0.45;
                        return new
                        {
                            Hole = h,
                            Mold1Score = m1,
                            BestOtherScore = bestOther,
                            DistNorm = distNorm,
                            Combined = combined
                        };
                    })
                    .Where(x => x.DistNorm <= CornerMaxNormalizedDistance &&
                                x.Mold1Score <= Mold1AbsoluteScoreThreshold &&
                                x.Mold1Score <= x.BestOtherScore * Mold1MarginFactor)
                    .OrderBy(x => x.Combined)
                    .FirstOrDefault();
                if (nearest is null)
                {
                    continue;
                }
                usedCorner.Add(nearest.Hole);
                rows.Add(new HoleAssignment(
                    nearest.Hole,
                    mold1.MoldId,
                    $"角落连续冲压({corner.Name})",
                    true,
                    IsNearOuterEdge(nearest.Hole, project.OuterRectangle),
                    BuildTopCandidates(nearest.Hole, molds, project.OuterRectangle, 3)));
            }

            foreach (var hole in holePool.Where(h => !usedCorner.Contains(h)))
            {
                var best = nonCornerMolds
                    .Select(m => new
                    {
                        MoldId = m.MoldId,
                        Score = ScoreForHoleAgainstMold(hole, m.Feature, project.OuterRectangle)
                    })
                    .OrderBy(x => x.Score)
                    .First();
                rows.Add(new HoleAssignment(
                    hole,
                    best.MoldId,
                    "单次冲压",
                    IsAnyCornerZone(hole, project.OuterRectangle),
                    IsNearOuterEdge(hole, project.OuterRectangle),
                    BuildTopCandidates(hole, nonCornerMolds, project.OuterRectangle, 3)));
            }

            return new MatchResult(rows);
        }

        private static double Dist2((double X, double Y) a, RectCorner b)
        {
            var dx = a.X - b.X;
            var dy = a.Y - b.Y;
            return dx * dx + dy * dy;
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
            return 0.12 * dw + 0.12 * dh + 0.16 * da + 0.1 * dp + 0.15 * dr + 0.35 * ds;
        }

        private static string BuildTopCandidates(HoleFeature hole, IEnumerable<MoldProfile> molds, RectBounds rect, int topN)
        {
            var tops = molds
                .Select(m => new
                {
                    m.MoldId,
                    Score = ScoreForHoleAgainstMold(hole, m.Feature, rect)
                })
                .OrderBy(x => x.Score)
                .Take(topN)
                .Select(x => $"M{x.MoldId:D2}:{x.Score:F3}");
            return string.Join(" | ", tops);
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
            var mirrored = MinCyclicRmse(a, b.Reverse().ToArray());
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

    public sealed record ProjectProfile(RectBounds OuterRectangle, IReadOnlyList<HoleFeature> Holes);
    public sealed record MoldProfile(int MoldId, string FilePath, HoleFeature Feature, IReadOnlyList<(double X, double Y)> OutlinePoints);
    public sealed record MatchResult(IReadOnlyList<HoleAssignment> HoleAssignments);
    public sealed record HoleAssignment(
        HoleFeature Hole,
        int MoldId,
        string PositionRelation,
        bool IsCornerCandidate,
        bool IsEdgeHole,
        string TopCandidates,
        double RotationDeg = 0,
        bool IsMirrored = false);
    public sealed record HoleFeature(string HoleType, (double X, double Y) Centroid, double Width, double Height, double Area, double Perimeter, double Rotation, double[] Signature);

    public sealed record RawBounds(double MinX, double MinY, double MaxX, double MaxY)
    {
        public double Width => MaxX - MinX;
        public double Height => MaxY - MinY;
    }

    public sealed record RectBounds(double MinX, double MinY, double MaxX, double MaxY)
    {
        public double Width => MaxX - MinX;
        public double Height => MaxY - MinY;
        public double Area => Width * Height;
        public IReadOnlyList<RectCorner> Corners => new[]
        {
            new RectCorner("左下", MinX, MinY),
            new RectCorner("左上", MinX, MaxY),
            new RectCorner("右下", MaxX, MinY),
            new RectCorner("右上", MaxX, MaxY)
        };
    }

    public sealed record RectCorner(string Name, double X, double Y);

    public sealed class MoldRow
    {
        public ImageSource? MoldPreview { get; set; }
        public string MoldCode { get; set; } = string.Empty;
        public string MoldName { get; set; } = string.Empty;
        public int UsedCount { get; set; }
        public string MatchType { get; set; } = string.Empty;
        public string Remark { get; set; } = string.Empty;
    }

    public sealed class PositionRow
    {
        public int Index { get; set; }
        public string HoleType { get; set; } = string.Empty;
        public int MoldId { get; set; }
        public string MoldCode { get; set; } = string.Empty;
        public double PosX { get; set; }
        public double PosY { get; set; }
        public string PositionRelation { get; set; } = string.Empty;
        public string IsCornerCandidate { get; set; } = string.Empty;
        public string IsEdgeHole { get; set; } = string.Empty;
        public string TopCandidates { get; set; } = string.Empty;
    }
}