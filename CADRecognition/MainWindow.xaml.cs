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
        private readonly InteractiveDxfPreview _viewer = new();
        private readonly ObservableCollection<MoldRow> _moldRows = [];
        private readonly ObservableCollection<PositionRow> _positionRows = [];
        private readonly List<string> _moldFiles = [];
        private MatchResult? _lastMatchResult;

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
            var molds = _moldFiles
                .Select((f, idx) => DxfAnalyzer.ExtractMold(idx + 1, f))
                .ToList();

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
                .GroupBy(x => x.MoldId)
                .ToDictionary(g => g.Key, g => g.Count());

            foreach (var mold in molds.OrderBy(x => x.MoldId))
            {
                useCounter.TryGetValue(mold.MoldId, out var count);
                _moldRows.Add(new MoldRow
                {
                    MoldCode = $"M{mold.MoldId:D2}",
                    MoldName = System.IO.Path.GetFileNameWithoutExtension(mold.FilePath),
                    UsedCount = count,
                    MatchType = mold.MoldId == 1 ? "角落连续冲压" : "单次冲压",
                    Remark = mold.MoldId == 1 ? "仅四角孔洞" : "普通孔洞"
                });
            }

            var i = 1;
            foreach (var row in result.HoleAssignments.OrderBy(x => x.Hole.Centroid.Y).ThenBy(x => x.Hole.Centroid.X))
            {
                _positionRows.Add(new PositionRow
                {
                    Index = i++,
                    HoleType = row.Hole.HoleType,
                    MoldId = row.MoldId,
                    MoldCode = $"M{row.MoldId:D2}",
                    PosX = Math.Round(row.Hole.Centroid.X, 3),
                    PosY = Math.Round(row.Hole.Centroid.Y, 3),
                    PositionRelation = row.PositionRelation
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
                _viewer.RenderAnnotations(_lastMatchResult.HoleAssignments);
            }
            else
            {
                _viewer.RenderAnnotations([]);
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
            StatusText.Text = $"已定位孔位 #{row.Index}（{row.MoldCode}）";
        }

        private async void PositionGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (PositionGrid.SelectedItem is not PositionRow row)
            {
                return;
            }
            _viewer.FocusHole(row.PosX, row.PosY, row.MoldId, targetZoom: 4.0);
            await _viewer.BlinkFocusAsync(row.PosX, row.PosY, row.MoldId);
            StatusText.Text = $"已放大定位孔位 #{row.Index}（{row.MoldCode}）";
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
            _markCanvas.RenderTransform = _group;

            _root.Children.Add(_sceneCanvas);
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
            _markCanvas.Width = viewWidth;
            _markCanvas.Height = viewHeight;
            _sceneCanvas.Children.Clear();
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

        public void RenderAnnotations(IReadOnlyList<HoleAssignment> assignments)
        {
            _markCanvas.Children.Clear();
            _focusRing = null;

            foreach (var ass in assignments)
            {
                var c = ass.Hole.Centroid;
                var p = ModelToCanvas(c.X, c.Y);
                var color = GetMoldColor(ass.MoldId);
                var brush = new SolidColorBrush(color);

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
            var feature = holes.OrderByDescending(h => h.Area).FirstOrDefault()
                ?? new HoleFeature("Unknown", (0, 0), 1, 1, 1, 1, 0, CreateCircleSignature(1, SignatureSamples));
            return new MoldProfile(moldId, path, feature);
        }

        public static ProjectProfile ExtractProject(DxfDocument doc)
        {
            var holes = ExtractHoles(doc);
            var outer = DetectOuterRectangle(doc);

            var maxHoleArea = Math.Max(outer.Area * 0.2, 1.0);
            var innerHoles = holes
                .Where(h => h.Centroid.X >= outer.MinX - 1e-6 &&
                            h.Centroid.X <= outer.MaxX + 1e-6 &&
                            h.Centroid.Y >= outer.MinY - 1e-6 &&
                            h.Centroid.Y <= outer.MaxY + 1e-6 &&
                            h.Area <= maxHoleArea)
                .ToList();
            return new ProjectProfile(outer, innerHoles);
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
                return bestRect;
            }

            var all = GetRawBounds(doc);
            return new RectBounds(all.MinX, all.MinY, all.MaxX, all.MaxY);
        }

        private static List<HoleFeature> ExtractHoles(DxfDocument doc)
        {
            var holes = new List<HoleFeature>();

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

            foreach (var pl in doc.Entities.Polylines2D.Where(p => p.IsClosed && p.Vertexes.Count >= 3))
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

            return holes;
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
            var compatibleForCorner = holePool
                .Select(h => new { Hole = h, Score = SimilarityScore(h, mold1.Feature) })
                .OrderBy(x => x.Score)
                .Take(Math.Min(12, holePool.Count))
                .ToList();

            var usedCorner = new HashSet<HoleFeature>();
            foreach (var corner in corners)
            {
                var nearest = compatibleForCorner
                    .Where(x => !usedCorner.Contains(x.Hole))
                    .OrderBy(x => Dist2(x.Hole.Centroid, corner))
                    .FirstOrDefault();
                if (nearest is null)
                {
                    continue;
                }
                usedCorner.Add(nearest.Hole);
                rows.Add(new HoleAssignment(nearest.Hole, mold1.MoldId, $"角落连续冲压({corner.Name})"));
            }

            var nonCornerMolds = molds.Where(m => m.MoldId != mold1.MoldId).ToList();
            if (nonCornerMolds.Count == 0)
            {
                nonCornerMolds.Add(mold1);
            }

            foreach (var hole in holePool.Where(h => !usedCorner.Contains(h)))
            {
                var best = nonCornerMolds
                    .Select(m => new { MoldId = m.MoldId, Score = SimilarityScore(hole, m.Feature) })
                    .OrderBy(x => x.Score)
                    .First();
                rows.Add(new HoleAssignment(hole, best.MoldId, "单次冲压"));
            }

            return new MatchResult(rows);
        }

        private static double Dist2((double X, double Y) a, RectCorner b)
        {
            var dx = a.X - b.X;
            var dy = a.Y - b.Y;
            return dx * dx + dy * dy;
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
    public sealed record MoldProfile(int MoldId, string FilePath, HoleFeature Feature);
    public sealed record MatchResult(IReadOnlyList<HoleAssignment> HoleAssignments);
    public sealed record HoleAssignment(HoleFeature Hole, int MoldId, string PositionRelation);
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
    }
}