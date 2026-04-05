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
using WpfRectangle = System.Windows.Shapes.Rectangle;

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
        private IReadOnlyList<(double X, double Y)> _lastOuterContourPoints = [];

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
            _lastOuterContourPoints = DxfAnalyzer.ExtractOuterContourForDebug(_projectDoc);
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
            if (withAnnotation && !string.IsNullOrWhiteSpace(path) && path == _projectFile)
            {
                _viewer.RenderCornerContours(
                    _lastProjectProfile?.OuterRectangle,
                    _lastOuterContourPoints,
                    _lastMatchResult?.GuidePaths,
                    _lastProjectProfile?.CornerCandidates);

                if (_lastMatchResult is not null)
                {
                    _viewer.RenderAnnotations(_lastMatchResult.HoleAssignments, _lastMolds);
                }
                else
                {
                    _viewer.RenderAnnotations([], []);
                }
            }
            else
            {
                _viewer.RenderCornerContours(null, null, null, null);
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

                // Draw a tiny cross at the reported centroid for coordinate verification.
                var cross = new Polyline
                {
                    Stroke = WpfBrushes.White,
                    StrokeThickness = 1,
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
                    StrokeThickness = 1,
                    Points = new PointCollection
                    {
                        new System.Windows.Point(p.X, p.Y - 4),
                        new System.Windows.Point(p.X, p.Y + 4),
                    }
                };
                _markCanvas.Children.Add(cross2);

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

        public void RenderCornerContours(
            RectBounds? rect,
            IReadOnlyList<(double X, double Y)>? outerContourPoints,
            IReadOnlyList<CornerStepPath>? cornerPaths,
            IReadOnlyList<HoleFeature>? cornerHints)
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
                StrokeThickness = 1.8,
                StrokeDashArray = new DoubleCollection([4, 3]),
                Fill = WpfBrushes.Transparent
            };
            Canvas.SetLeft(rectBox, Math.Min(r1.X, r2.X));
            Canvas.SetTop(rectBox, Math.Min(r1.Y, r2.Y));
            _zoneCanvas.Children.Add(rectBox);

            // 2) 画当前识别到的真实外轮廓（红色）
            if (outerContourPoints is not null && outerContourPoints.Count >= 2)
            {
                var polyAll = new Polyline
                {
                    Stroke = new SolidColorBrush(WpfColor.FromArgb(245, 255, 82, 82)),
                    StrokeThickness = 1.8,
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

            var edgeTol = Math.Clamp(Math.Min(rect.Width, rect.Height) * 0.0004, 0.05, 0.35);
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
                var innerSafe = Math.Max(edgeTol * 1.6, 0.35);
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
                    StrokeThickness = 2.4,
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

            // 辅助线（紫色）：M01 连续冲压使用的外偏移路径。
            if (cornerPaths is not null)
            {
                foreach (var gp in cornerPaths)
                {
                    if (gp.Points is null || gp.Points.Count < 2)
                    {
                        continue;
                    }

                    var guide = new Polyline
                    {
                        Stroke = new SolidColorBrush(WpfColor.FromArgb(235, 186, 104, 200)),
                        StrokeThickness = 2.0,
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
                    StrokeThickness = 2.0,
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
                FontSize = 11,
                FontWeight = FontWeights.Bold,
                Background = new SolidColorBrush(WpfColor.FromArgb(120, 0, 0, 0))
            };
            Canvas.SetLeft(label, labelAnchor.X + 6);
            Canvas.SetTop(label, labelAnchor.Y - 18);
            _zoneCanvas.Children.Add(label);
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
            var outline = ExtractMoldOutline(doc, feature.Centroid);
            return new MoldProfile(moldId, path, feature, outline);
        }

        public static ProjectProfile ExtractProject(DxfDocument doc)
        {
            var holes = ExtractHoles(doc);
            var outer = DetectOuterRectangle(doc);
            var edgeCandidates = ExtractEdgePartialCandidates(doc, outer);
            var cornerCandidates = ExtractCornerMissingFeatures(doc, outer);
            var cornerStepPaths = ExtractCornerStepPaths(doc, outer);
            var contourPaths = ExtractContourDifferencePaths(doc, outer);

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
            return new ProjectProfile(
                outer,
                DeduplicateHoles(innerHoles),
                cornerCandidates,
                edgeCandidates,
                cornerStepPaths,
                contourPaths);
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

            var edgeTol = Math.Clamp(Math.Min(rect.Width, rect.Height) * 0.0004, 0.05, 0.35);
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
            var tol = Math.Clamp(diag * 0.0005, 1e-4, 0.2);

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
            var pts = CollectGeometryPoints(doc);
            var result = new List<EdgeCandidate>();
            if (pts.Count < 10)
            {
                // still try polyline-based extraction below
            }

            // Edge band (exclude corner zones)
            var cornerX = outer.Width * 0.22;
            var cornerY = outer.Height * 0.22;
            var band = Math.Max(Math.Min(outer.Width, outer.Height) * 0.06, 15.0);
            var depth = Math.Max(Math.Min(outer.Width, outer.Height) * 0.02, 5.0);
            var gap = Math.Max(Math.Min(outer.Width, outer.Height) * 0.03, 25.0);

            IEnumerable<(double X, double Y)> top = pts.Where(p =>
                p.Y <= outer.MaxY - depth && p.Y >= outer.MaxY - band &&
                p.X > outer.MinX + cornerX && p.X < outer.MaxX - cornerX);
            IEnumerable<(double X, double Y)> bottom = pts.Where(p =>
                p.Y >= outer.MinY + depth && p.Y <= outer.MinY + band &&
                p.X > outer.MinX + cornerX && p.X < outer.MaxX - cornerX);
            IEnumerable<(double X, double Y)> left = pts.Where(p =>
                p.X >= outer.MinX + depth && p.X <= outer.MinX + band &&
                p.Y > outer.MinY + cornerY && p.Y < outer.MaxY - cornerY);
            IEnumerable<(double X, double Y)> right = pts.Where(p =>
                p.X <= outer.MaxX - depth && p.X >= outer.MaxX - band &&
                p.Y > outer.MinY + cornerY && p.Y < outer.MaxY - cornerY);

            void AddGroups(IEnumerable<(double X, double Y)> bandPts, bool sortByX, string side)
            {
                var sorted = (sortByX ? bandPts.OrderBy(p => p.X) : bandPts.OrderBy(p => p.Y)).ToList();
                if (sorted.Count < 8)
                {
                    return;
                }

                var current = new List<(double X, double Y)>();
                for (var i = 0; i < sorted.Count; i++)
                {
                    if (current.Count == 0)
                    {
                        current.Add(sorted[i]);
                        continue;
                    }
                    var prev = current[^1];
                    var d = Math.Sqrt((sorted[i].X - prev.X) * (sorted[i].X - prev.X) + (sorted[i].Y - prev.Y) * (sorted[i].Y - prev.Y));
                    if (d <= gap)
                    {
                        current.Add(sorted[i]);
                    }
                    else
                    {
                        if (current.Count >= 12)
                        {
                            result.Add(BuildEdgeCandidate(side, current));
                        }
                        current = [sorted[i]];
                    }
                }
                if (current.Count >= 12)
                {
                    result.Add(BuildEdgeCandidate(side, current));
                }
            }

            AddGroups(top, true, "Top");
            AddGroups(bottom, true, "Bottom");
            AddGroups(left, false, "Left");
            AddGroups(right, false, "Right");

            // Extra: use open polylines as edge-notch candidates (common DXF export).
            foreach (var pl in doc.Entities.Polylines2D.Where(p => !IsPolylineClosedLike(p) && p.Vertexes.Count >= 3))
            {
                var polyPts = ExpandPolyline2D(pl, 24).ToList();
                if (polyPts.Count < 6)
                {
                    continue;
                }
                var minX = polyPts.Min(p => p.X);
                var maxX = polyPts.Max(p => p.X);
                var minY = polyPts.Min(p => p.Y);
                var maxY = polyPts.Max(p => p.Y);
                var w = maxX - minX;
                var h = maxY - minY;
                if (w < 1e-6 || h < 1e-6)
                {
                    continue;
                }

                var cx = polyPts.Average(p => p.X);
                var cy = polyPts.Average(p => p.Y);
                var nearLeft = Math.Abs(cx - outer.MinX) <= band;
                var nearRight = Math.Abs(outer.MaxX - cx) <= band;
                var nearBottom = Math.Abs(cy - outer.MinY) <= band;
                var nearTop = Math.Abs(outer.MaxY - cy) <= band;

                // Exclude corner zones
                var inCornerZone = (cx <= outer.MinX + cornerX || cx >= outer.MaxX - cornerX) &&
                                   (cy <= outer.MinY + cornerY || cy >= outer.MaxY - cornerY);
                if (inCornerZone)
                {
                    continue;
                }

                string? side = null;
                if (nearLeft) side = "Left";
                else if (nearRight) side = "Right";
                else if (nearBottom) side = "Bottom";
                else if (nearTop) side = "Top";
                if (side is null)
                {
                    continue;
                }

                // Avoid duplicates near existing candidates
                var thresh = Math.Max(Math.Min(w, h) * 0.8, 40.0);
                var dup = result.Any(r =>
                {
                    var dx = r.Centroid.X - cx;
                    var dy = r.Centroid.Y - cy;
                    return Math.Sqrt(dx * dx + dy * dy) <= thresh;
                });
                if (dup)
                {
                    continue;
                }

                result.Add(BuildEdgeCandidate(side, polyPts));
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
                var tolPos = Math.Max(Math.Min(h.Width, h.Height) * 0.15, 2.0);
                var tolSize = Math.Max(Math.Min(h.Width, h.Height) * 0.12, 2.0);
                var exists = result.Any(r =>
                    Math.Abs(r.Centroid.X - h.Centroid.X) <= tolPos &&
                    Math.Abs(r.Centroid.Y - h.Centroid.Y) <= tolPos &&
                    Math.Abs(r.Width - h.Width) <= tolSize &&
                    Math.Abs(r.Height - h.Height) <= tolSize);
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

        private static List<(double X, double Y)> ExtractMoldOutline(DxfDocument doc, (double X, double Y) referenceCenter)
        {
            var pts = CollectGeometryPoints(doc);
            if (pts.Count < 2)
            {
                return [];
            }
            var ordered = pts
                .Select(p => (X: p.X - referenceCenter.X, Y: p.Y - referenceCenter.Y))
                .OrderBy(p => Math.Atan2(p.Y, p.X))
                .ToList();

            var step = Math.Max(1, ordered.Count / 80);
            var outline = new List<(double X, double Y)>();
            for (var i = 0; i < ordered.Count; i += step)
            {
                outline.Add((ordered[i].X, ordered[i].Y));
            }
            // Recenter again after downsampling to remove any residual bias.
            if (outline.Count > 2)
            {
                var ox = outline.Average(p => p.X);
                var oy = outline.Average(p => p.Y);
                for (var i = 0; i < outline.Count; i++)
                {
                    outline[i] = (outline[i].X - ox, outline[i].Y - oy);
                }
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
        private const int CornerPathMaxPointsPerCorner = 24;

        public MatchResult Match(ProjectProfile project, IReadOnlyList<MoldProfile> molds)
        {
            var rows = new List<HoleAssignment>();
            var guidePaths = new List<CornerStepPath>();
            if (molds.Count == 0 || project.Holes.Count == 0)
            {
                return new MatchResult(rows, guidePaths);
            }

            var corners = project.OuterRectangle.Corners;
            var holePool = project.Holes.ToList();
            var mold1 = molds.FirstOrDefault(m => m.MoldId == 1) ?? molds[0];
            var nonCornerMolds = molds.Where(m => m.MoldId != mold1.MoldId).ToList();
            if (nonCornerMolds.Count == 0)
            {
                nonCornerMolds.Add(mold1);
            }

            // M01：沿“青色差集线的外偏移路径”做连续冲压。
            var contourStamps = GenerateContinuousContourStampCenters(project, mold1, guidePaths);
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

            var diag = Math.Sqrt(project.OuterRectangle.Width * project.OuterRectangle.Width +
                                 project.OuterRectangle.Height * project.OuterRectangle.Height);
            var safeDiag = Math.Max(diag, 1e-6);

            foreach (var hole in holePool)
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

            // Edge partial stamping: translation-only, trimmed chamfer to allow partial imprint
            var addedEdge = new List<HoleFeature>();
            foreach (var cand in project.EdgeCandidates)
            {
                // If this edge candidate overlaps a real hole, skip it to avoid "one hole, two molds".
                var overlapThreshold = Math.Max(Math.Min(cand.Width, cand.Height) * 0.6, 25.0);
                var overlapsRealHole = holePool.Any(h =>
                {
                    var dx = h.Centroid.X - cand.Centroid.X;
                    var dy = h.Centroid.Y - cand.Centroid.Y;
                    return Math.Sqrt(dx * dx + dy * dy) <= overlapThreshold;
                });
                if (overlapsRealHole)
                {
                    continue;
                }

                var best = nonCornerMolds
                    .Select(m =>
                    {
                        var score = TrimmedChamferScore(m.OutlinePoints, cand.Points, out var placement);
                        return new { m.MoldId, Score = score, Placement = placement };
                    })
                    .OrderBy(x => x.Score)
                    .FirstOrDefault();
                if (best is null || double.IsInfinity(best.Score))
                {
                    continue;
                }

                var placementHole = new HoleFeature(
                    $"EdgePartial:{cand.Side}",
                    best.Placement,
                    cand.Width,
                    cand.Height,
                    Math.Max(cand.Width * cand.Height * 0.45, 1.0),
                    Math.Max(cand.Perimeter, 1.0),
                    0,
                    cand.Signature);

                // De-duplicate edge partials near each other (one notch -> one mold).
                var dup = addedEdge.Any(h =>
                {
                    var dx = h.Centroid.X - placementHole.Centroid.X;
                    var dy = h.Centroid.Y - placementHole.Centroid.Y;
                    return Math.Sqrt(dx * dx + dy * dy) <= overlapThreshold;
                });
                if (dup)
                {
                    continue;
                }
                addedEdge.Add(placementHole);
                rows.Add(new HoleAssignment(
                    placementHole,
                    best.MoldId,
                    "边缘局部冲压",
                    false,
                    true,
                    BuildTopCandidates(placementHole, nonCornerMolds, project.OuterRectangle, 3)));
            }

            var cleaned = DeduplicateAssignments(rows);
            return new MatchResult(cleaned, guidePaths);
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

        private static IReadOnlyList<HoleFeature> GenerateContinuousContourStampCenters(ProjectProfile project, MoldProfile mold1, List<CornerStepPath> guidePaths)
        {
            if (project.ContourPaths is null || project.ContourPaths.Count == 0)
            {
                return [];
            }

            var outline = mold1.OutlinePoints;
            if (outline is null || outline.Count < 2)
            {
                return [];
            }

            // CAD 偏移量：向外偏移半个模具宽度。
            var offsetDist = Math.Max(Math.Min(mold1.Feature.Width, mold1.Feature.Height) * 0.5, 0.8);
            var moldStep = Math.Max(EstimateOutlineStep(outline) * 0.55, 2.5);
            var points = new List<HoleFeature>();

            foreach (var contourPath in project.ContourPaths)
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

                // 先对整条线做外偏移（类似 CAD Offset），再沿偏移后的线采样。
                var offsetChain = OffsetPolylineOutward(chain, project.OuterRectangle, offsetDist);
                if (offsetChain.Count < 2)
                {
                    continue;
                }

                // 输出辅助线：用于界面可视化核对。
                guidePaths.Add(new CornerStepPath(contourPath.CornerName, offsetChain));

                var sampled = SampleAlongPolyline(offsetChain, moldStep);
                foreach (var s in sampled)
                {
                    points.Add(new HoleFeature(
                        $"ContourPath:{contourPath.CornerName}",
                        (s.X, s.Y),
                        mold1.Feature.Width,
                        mold1.Feature.Height,
                        Math.Max(mold1.Feature.Area, 1.0),
                        Math.Max(mold1.Feature.Perimeter, 1.0),
                        0,
                        mold1.Feature.Signature));
                }

                // 内拐点必须命中：偏移线每个折点都强制补一个冲压中心。
                for (var vi = 1; vi < offsetChain.Count - 1; vi++)
                {
                    var v = offsetChain[vi];
                    points.Add(new HoleFeature(
                        $"ContourCornerHit:{contourPath.CornerName}",
                        v,
                        mold1.Feature.Width,
                        mold1.Feature.Height,
                        Math.Max(mold1.Feature.Area, 1.0),
                        Math.Max(mold1.Feature.Perimeter, 1.0),
                        0,
                        mold1.Feature.Signature));
                }
            }

            // 去重：先保留“内拐点强制命中”点，再并入普通连续点。
            var dedup = new List<HoleFeature>();

            var mandatoryHits = points
                .Where(p => p.HoleType.StartsWith("ContourCornerHit", StringComparison.Ordinal))
                .OrderBy(p => p.Centroid.Y)
                .ThenBy(p => p.Centroid.X)
                .ToList();

            foreach (var p in mandatoryHits)
            {
                var tol = Math.Max(Math.Min(p.Width, p.Height) * 0.12, 0.6);
                var exists = dedup.Any(d =>
                {
                    var dx = d.Centroid.X - p.Centroid.X;
                    var dy = d.Centroid.Y - p.Centroid.Y;
                    return Math.Sqrt(dx * dx + dy * dy) <= tol;
                });
                if (!exists)
                {
                    dedup.Add(p);
                }
            }

            foreach (var p in points.Where(p => !p.HoleType.StartsWith("ContourCornerHit", StringComparison.Ordinal))
                                    .OrderBy(p => p.Centroid.Y)
                                    .ThenBy(p => p.Centroid.X))
            {
                var tol = Math.Max(Math.Min(p.Width, p.Height) * 0.25, 1.2);
                var exists = dedup.Any(d =>
                {
                    var dx = d.Centroid.X - p.Centroid.X;
                    var dy = d.Centroid.Y - p.Centroid.Y;
                    return Math.Sqrt(dx * dx + dy * dy) <= tol;
                });
                if (!exists)
                {
                    dedup.Add(p);
                }
            }

            return dedup;
        }

        private static double EstimateOutlineStep(IReadOnlyList<(double X, double Y)> outline)
        {
            var lengths = new List<double>();
            for (var i = 1; i < outline.Count; i++)
            {
                var dx = outline[i].X - outline[i - 1].X;
                var dy = outline[i].Y - outline[i - 1].Y;
                var len = Math.Sqrt(dx * dx + dy * dy);
                if (len > 1e-6)
                {
                    lengths.Add(len);
                }
            }
            if (lengths.Count == 0)
            {
                return 10.0;
            }
            lengths.Sort();
            return lengths[Math.Max(0, lengths.Count / 4)];
        }

        private static List<(double X, double Y)> OffsetPolylineOutward(
            IReadOnlyList<(double X, double Y)> chain,
            RectBounds outer,
            double offset)
        {
            var result = new List<(double X, double Y)>();
            if (chain.Count < 2)
            {
                return result;
            }

            var center = ((outer.MinX + outer.MaxX) * 0.5, (outer.MinY + outer.MaxY) * 0.5);

            (double NX, double NY) OutwardNormal((double X, double Y) a, (double X, double Y) b)
            {
                var dx = b.X - a.X;
                var dy = b.Y - a.Y;
                var len = Math.Sqrt(dx * dx + dy * dy);
                if (len <= 1e-9)
                {
                    return (0, 0);
                }

                var tx = dx / len;
                var ty = dy / len;
                var n1 = (-ty, tx);
                var n2 = (ty, -tx);
                var mid = ((a.X + b.X) * 0.5, (a.Y + b.Y) * 0.5);
                var toCenter = (center.Item1 - mid.Item1, center.Item2 - mid.Item2);
                var dot = n1.Item1 * toCenter.Item1 + n1.Item2 * toCenter.Item2;
                var inward = dot >= 0 ? n1 : n2;
                return (-inward.Item1, -inward.Item2);
            }

            (double X, double Y)? LineIntersection(
                (double X, double Y) p,
                (double X, double Y) r,
                (double X, double Y) q,
                (double X, double Y) s)
            {
                var rxs = r.X * s.Y - r.Y * s.X;
                if (Math.Abs(rxs) <= 1e-9)
                {
                    return null;
                }
                var qmp = (q.X - p.X, q.Y - p.Y);
                var t = (qmp.Item1 * s.Y - qmp.Item2 * s.X) / rxs;
                return (p.X + t * r.X, p.Y + t * r.Y);
            }

            // 起点
            {
                var n = OutwardNormal(chain[0], chain[1]);
                result.Add((chain[0].X + n.NX * offset, chain[0].Y + n.NY * offset));
            }

            // 中间点：相邻两条偏移线求交
            for (var i = 1; i < chain.Count - 1; i++)
            {
                var a0 = chain[i - 1];
                var a1 = chain[i];
                var b0 = chain[i];
                var b1 = chain[i + 1];

                var na = OutwardNormal(a0, a1);
                var nb = OutwardNormal(b0, b1);

                (double X, double Y) pa = (a0.X + na.NX * offset, a0.Y + na.NY * offset);
                (double X, double Y) qa = (a1.X + na.NX * offset, a1.Y + na.NY * offset);
                (double X, double Y) pb = (b0.X + nb.NX * offset, b0.Y + nb.NY * offset);
                (double X, double Y) qb = (b1.X + nb.NX * offset, b1.Y + nb.NY * offset);

                (double X, double Y) ra = (qa.X - pa.X, qa.Y - pa.Y);
                (double X, double Y) rb = (qb.X - pb.X, qb.Y - pb.Y);
                var cross = LineIntersection(pa, ra, pb, rb);

                if (cross.HasValue)
                {
                    result.Add(cross.Value);
                }
                else
                {
                    // 平行或数值不稳定时退化为点法线偏移。
                    var nx = na.NX + nb.NX;
                    var ny = na.NY + nb.NY;
                    var nl = Math.Sqrt(nx * nx + ny * ny);
                    if (nl <= 1e-9)
                    {
                        nx = na.NX;
                        ny = na.NY;
                        nl = Math.Sqrt(nx * nx + ny * ny);
                    }
                    nx /= nl;
                    ny /= nl;
                    result.Add((chain[i].X + nx * offset, chain[i].Y + ny * offset));
                }
            }

            // 终点
            {
                var n = OutwardNormal(chain[^2], chain[^1]);
                result.Add((chain[^1].X + n.NX * offset, chain[^1].Y + n.NY * offset));
            }

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

        private sealed record PathSample(double X, double Y, double TX, double TY);

        private static List<PathSample> SampleAlongPolyline(IReadOnlyList<(double X, double Y)> polyline, double step)
        {
            var result = new List<PathSample>();
            if (polyline.Count < 2)
            {
                return result;
            }

            var firstDx = polyline[1].X - polyline[0].X;
            var firstDy = polyline[1].Y - polyline[0].Y;
            var firstLen = Math.Sqrt(firstDx * firstDx + firstDy * firstDy);
            if (firstLen > 1e-9)
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
                else
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

            var result = new List<HoleAssignment>();

            // 先保留“内拐点强制命中”点，避免后续去重把它们消掉。
            var cornerHits = source
                .Where(r => r.Hole.HoleType.StartsWith("ContourCornerHit", StringComparison.Ordinal))
                .OrderBy(r => r.MoldId)
                .ThenBy(r => r.Hole.Centroid.Y)
                .ThenBy(r => r.Hole.Centroid.X);

            foreach (var row in cornerHits)
            {
                var tol = Math.Max(Math.Min(row.Hole.Width, row.Hole.Height) * 0.08, 0.35);
                var dup = result.Any(existing =>
                {
                    if (existing.MoldId != row.MoldId)
                    {
                        return false;
                    }
                    var dx = existing.Hole.Centroid.X - row.Hole.Centroid.X;
                    var dy = existing.Hole.Centroid.Y - row.Hole.Centroid.Y;
                    return Math.Sqrt(dx * dx + dy * dy) <= tol;
                });

                if (!dup)
                {
                    result.Add(row);
                }
            }

            foreach (var row in source
                         .Where(r => !r.Hole.HoleType.StartsWith("ContourCornerHit", StringComparison.Ordinal))
                         .OrderBy(r => r.MoldId)
                         .ThenBy(r => r.Hole.Centroid.Y)
                         .ThenBy(r => r.Hole.Centroid.X))
            {
                var sizeRef = Math.Max(Math.Min(row.Hole.Width, row.Hole.Height), 1.0);
                var tol = Math.Max(sizeRef * 0.35, 3.0);

                var dup = result.Any(existing =>
                {
                    if (existing.MoldId != row.MoldId)
                    {
                        return false;
                    }
                    var dx = existing.Hole.Centroid.X - row.Hole.Centroid.X;
                    var dy = existing.Hole.Centroid.Y - row.Hole.Centroid.Y;
                    return Math.Sqrt(dx * dx + dy * dy) <= tol;
                });

                if (!dup)
                {
                    result.Add(row);
                }
            }
            return result;
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

}