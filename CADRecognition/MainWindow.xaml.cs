using Microsoft.Win32;
using netDxf;
using netDxf.Entities;
using CADRecognition;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
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
        private string? _selectedM01File;
        private MatchResult? _lastMatchResult;
        private ProjectProfile? _lastProjectProfile;
        private List<MoldProfile> _lastMolds = [];
        private IReadOnlyList<(double X, double Y)> _lastOuterContourPoints = [];

        private string? _projectFile;
        private DxfDocument? _projectDoc;
        private bool _compactAnnotation = false;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            MoldCountText.Text = "0";
            ProjectFileText.Text = "未加载";
            PreviewHost.Content = _viewer;
            _viewer.SetCompactMode(_compactAnnotation);
            FileTreeView.Items.Clear();
        }

        public ObservableCollection<MoldRow> MoldRows => _moldRows;
        public ObservableCollection<PositionRow> PositionRows => _positionRows;

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
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

        private void ImportProjectDxf_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new WinOpenFileDialog
            {
                Filter = "CAD 鏂囦欢 (*.dxf;*.dwg)|*.dxf;*.dwg|DXF 鏂囦欢 (*.dxf)|*.dxf|DWG 鏂囦欢 (*.dwg)|*.dwg",
                Multiselect = false
            };
            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            _projectFile = dialog.FileName;
            _projectDoc = LoadCadDocument(_projectFile);
            var removedProjectLines = 0;
            _documentCache[_projectFile] = _projectDoc;

            // 瀵煎叆鏂板伐绋嬫椂娓呯┖涓婁竴寮犲浘绾哥殑璇嗗埆/鏍囨敞灞曠ず鐘舵€併€?
            _lastMatchResult = null;
            _lastProjectProfile = null;
            _lastOuterContourPoints = [];
            _moldRows.Clear();
            _positionRows.Clear();
            LegendPanel.Children.Clear();

            ProjectFileText.Text = System.IO.Path.GetFileName(_projectFile);
            RefreshFileList();
            RenderPreview(_projectDoc, _projectFile, withAnnotation: false);
            StatusText.Text = removedProjectLines > 0
                ? $"工程 DXF 已加载，已去重 {removedProjectLines} 条。"
                : "工程 DXF 已加载。";
        }

        private void ImportMoldsDxf_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new WinOpenFileDialog
            {
                Filter = "CAD 鏂囦欢 (*.dxf;*.dwg)|*.dxf;*.dwg|DXF 鏂囦欢 (*.dxf)|*.dxf|DWG 鏂囦欢 (*.dwg)|*.dwg",
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
                var moldDoc = LoadCadDocument(file);
                _documentCache[file] = moldDoc;
            }

            M01MoldComboBox.ItemsSource = _moldFiles.Select(System.IO.Path.GetFileName).ToList();
            _selectedM01File = _moldFiles.FirstOrDefault();
            M01MoldComboBox.SelectedIndex = _selectedM01File is null ? -1 : 0;

            MoldCountText.Text = _moldFiles.Count.ToString(CultureInfo.InvariantCulture);
            RefreshFileList();
            StatusText.Text = $"已导入 {_moldFiles.Count} 张模具 CAD 图。";
        }

        private DxfDocument LoadCadDocument(string path)
        {
            try
            {
                return CadDocumentLoader.Load(path);
            }
            catch (Exception ex)
            {
                StatusText.Text = $"读取失败：{System.IO.Path.GetFileName(path)}，{ex.Message}";
                throw;
            }
        }


        private string? ResolveSelectedM01File()
        {
            if (_moldFiles.Count == 0)
            {
                return null;
            }

            if (M01MoldComboBox.SelectedItem is string selectedName)
            {
                var matched = _moldFiles.FirstOrDefault(f =>
                    string.Equals(System.IO.Path.GetFileName(f), selectedName, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrWhiteSpace(matched))
                {
                    return matched;
                }
            }

            return _moldFiles.First();
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

            _selectedM01File = ResolveSelectedM01File();
            var orderedFiles = _moldFiles
                .OrderBy(f => string.Equals(f, _selectedM01File, StringComparison.OrdinalIgnoreCase) ? 0 : 1)
                .ThenBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var molds = orderedFiles
                .Select((f, idx) => DxfAnalyzer.ExtractMold(idx + 1, f))
                .ToList();
            _lastMolds = molds;

            var matcher = new MoldMatcher();
            var result = matcher.Match(project, molds);
            _lastMatchResult = result;
            RenderResult(result, molds, project.OuterRectangle);
            RenderPreview(_projectDoc, _projectFile, withAnnotation: true);
            StatusText.Text = $"识别完成：外轮廓 {project.OuterRectangle.Width:F2} x {project.OuterRectangle.Height:F2}，孔洞 {result.HoleAssignments.Count} 个。";
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            if (_lastMatchResult is null || _projectFile is null)
            {
                StatusText.Text = "请先完成识图后再导出。";
                return;
            }

            var dialog = new TcpExportDialog(BuildTcpExportModel())
            {
                Owner = this
            };

            if (dialog.ShowDialog() == true)
            {
                StatusText.Text = "已导出并发送 JSON。";
            }
        }

        private TcpExportModel BuildTcpExportModel()
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
            model.Stage1PunchCount = _positionRows.Count;
            model.Stage2PunchCount = 0;
            model.Spare2 = 0;
            model.PlateLength = _lastProjectProfile?.OuterRectangle.Width ?? 0;
            model.PlateWidth = _lastProjectProfile?.OuterRectangle.Height ?? 0;
            model.PlateThickness = 0;
            model.Spare3 = 0;
            model.Spare4 = 0;
            model.CustomContent = string.Empty;

            var boundary = GetRecognitionBoundary();
            var stage1Rows = SplitRowsByBoundary(_positionRows, boundary, true).ToList();
            var stage2Rows = SplitRowsByBoundary(_positionRows, boundary, false).ToList();
            model.Stage1DiagramCoordinates = stage1Rows.Select(r => new TcpCoordinateRow { X = r.PosX, Y = r.PosY }).ToList();
            model.Stage2DiagramCoordinates = stage2Rows.Select(r => new TcpCoordinateRow { X = r.PosX, Y = r.PosY }).ToList();
            model.Stage1PositionMoldIds = stage1Rows.Select(r => r.MoldId).ToList();
            model.Stage2PositionMoldIds = stage2Rows.Select(r => r.MoldId).ToList();
            model.Stage1PunchMoldIds = stage1Rows.Select(r => r.MoldId).ToList();
            model.Stage2PunchMoldIds = stage2Rows.Select(r => r.MoldId).ToList();
            return model;
        }

        private RectBounds GetRecognitionBoundary()
        {
            if (_lastProjectProfile is not null)
            {
                return _lastProjectProfile.OuterRectangle;
            }

            if (_positionRows.Count > 0)
            {
                var minX = _positionRows.Min(x => x.PosX);
                var minY = _positionRows.Min(x => x.PosY);
                var maxX = _positionRows.Max(x => x.PosX);
                var maxY = _positionRows.Max(x => x.PosY);
                return new RectBounds(minX, minY, maxX, maxY);
            }

            return new RectBounds(0, 0, 1, 1);
        }

        private IEnumerable<PositionRow> SplitRowsByBoundary(IEnumerable<PositionRow> rows, RectBounds boundary, bool upperHalf)
        {
            var midpoint = boundary.MinY + boundary.Height / 2.0;
            foreach (var row in rows.OrderBy(x => x.PosY).ThenBy(x => x.PosX))
            {
                var isUpper = row.PosY >= midpoint;
                if (upperHalf == isUpper)
                {
                    yield return row;
                }
            }
        }

        private void RenderResult(MatchResult result, IReadOnlyList<MoldProfile> molds, RectBounds outer)
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
                    MoldCode = row.MoldId > 0 ? $"M{row.MoldId:D2}" : "未匹配",
                    PosX = Math.Round(row.Hole.Centroid.X - outer.MinX, 0),
                    PosY = Math.Round(row.Hole.Centroid.Y - outer.MinY, 0),
                    AbsX = row.Hole.Centroid.X,
                    AbsY = row.Hole.Centroid.Y,
                    PositionRelation = row.PositionRelation,
                    IsCornerCandidate = row.IsCornerCandidate ? "是" : "否",
                    IsEdgeHole = row.IsEdgeHole ? "是" : "否",
                    TopCandidates = row.TopCandidates,
                    AreaRatio = row.AreaRatioInfo,
                    FailureReason = row.FailureReason
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
                    FontSize = 9,
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
            RenderCadPreview(doc, _viewer, path, withAnnotation);
            PreviewHintText.Visibility = Visibility.Collapsed;
        }

        private void RenderCadPreview(DxfDocument doc, InteractiveDxfPreview viewer, string? path, bool withAnnotation)
        {
            _previewPlugin.CreatePreview(doc, viewer);
            if (withAnnotation && !string.IsNullOrWhiteSpace(path) && path == _projectFile)
            {
                viewer.RenderCornerContours(
                    _lastProjectProfile?.OuterRectangle,
                    _lastOuterContourPoints,
                    _lastMatchResult?.GuidePaths,
                    _lastProjectProfile?.CornerCandidates);

                if (_lastMatchResult is not null)
                {
                    viewer.RenderAnnotations(_lastMatchResult.HoleAssignments, _lastMolds);
                }
                else
                {
                    viewer.RenderAnnotations([], []);
                }
            }
            else
            {
                viewer.RenderCornerContours(null, null, null, null);
                viewer.RenderAnnotations([], []);
            }
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
            _viewer.FocusHole(row.AbsX, row.AbsY, row.MoldId);
            StatusText.Text = $"已定位孔位 #{row.Index}（{row.MoldCode}），角候选={row.IsCornerCandidate}，边缘孔={row.IsEdgeHole}，Top3={row.TopCandidates}";
        }

        private async void PositionGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (PositionGrid.SelectedItem is not PositionRow row)
            {
                return;
            }
            _viewer.FocusHole(row.AbsX, row.AbsY, row.MoldId, targetZoom: 4.0);
            await _viewer.BlinkFocusAsync(row.AbsX, row.AbsY, row.MoldId);
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
                doc = LoadCadDocument(path);
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
                CadPreviewRenderService.DrawToDrawingContext(dc, doc, bounds, width, height, scale, margin);
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

            CadPreviewRenderService.DrawToCanvas(canvas, document, bounds, viewWidth, viewHeight, scale, margin);
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

                var shouldShowLabel = !_compactMode && ass.MoldId != 1;

                if (shouldShowLabel)
                {
                    var text = new TextBlock
                    {
                        Text = $"M{ass.MoldId:D2}",
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
            IReadOnlyList<HoleFeature>? cornerHints)
        {
            _zoneCanvas.Children.Clear();
            if (rect is null)
            {
                return;
            }

            // 1) 鐢绘渶灏忓鍖呯煩褰紙榛勮壊锛?
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

            // 2) 鐢诲綋鍓嶈瘑鍒埌鐨勭湡瀹炲杞粨锛堢孩鑹诧級
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

            // 3) 闈掕壊绾?= 绾㈣壊澶栬疆寤?- 鐭╁舰澶栬疆寤擄紙鎸夆€滅嚎娈靛樊闆嗏€濈粯鍒讹紝閬垮厤璺ㄦ璇繛锛?
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

            // 鍘婚櫎鐩搁偦閲嶅鐐?鏋佺煭杈癸紝閬垮厤宸泦鍒囨鏃朵涪澶辩煭鍙伴樁銆?
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

            // 鎸夆€滅嚎娈碘€濆仛宸泦锛氬彧瑕佺嚎娈典腑鐐逛笉鍦ㄧ煩褰㈣竟涓婏紝灏变繚鐣欒娈点€?
            // 娉ㄦ剰瑕佸寘鍚灏鹃棴鍚堟锛岄伩鍏嶆紡鎺夎疆寤撹捣鐐归檮杩戠殑涓€娈点€?
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

                // 杈逛繚鎶わ細鑻ョ嚎娈垫槑鏄句綅浜庣煩褰㈠唴閮紙绂讳换浣曞杈规湁瀹夊叏闂磋窛锛夛紝寮哄埗淇濈暀銆?
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
                // 杈呭姪绾匡紙绱壊锛夛細M01 杩炵画鍐插帇浣跨敤鐨勫鍋忕Щ璺緞銆?
                if (cornerPaths is not null)
                {
                    foreach (var gp in cornerPaths)
                    {
                        if (gp.Points is null || gp.Points.Count < 2)
                        {
                            continue;
                        }

                        // 绱壊绾匡細鏄剧ず瀹屾暣 offset 璺緞锛堜笉鍋氱鐐?鎷愮偣鍘嬬缉锛夛紝渚夸簬鏍稿鍑犱綍鏈韩銆?
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
                    // 璋冭瘯灞傦細琚垽瀹氫负鈥滅煩褰㈣竟鑰屽垹闄も€濈殑绾挎锛堟鑹诧級
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
                        Text = "寰呭啿杞粨",
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



}
