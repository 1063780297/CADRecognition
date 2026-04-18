using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using TextBox = System.Windows.Controls.TextBox;

namespace CADRecognition
{
    public partial class TcpExportDialog : Window
    {
        private readonly string _customContentFilePath;
        private readonly string _tcpHistoryFilePath;
        private readonly TcpCustomContentStore _customContentStore;
        private readonly TcpConnectionHistoryStore _tcpHistoryStore;
        private readonly TcpCommService _tcpCommService = new();
        private readonly ObservableCollection<TcpGridRow> _stage1Rows = new();
        private readonly ObservableCollection<TcpGridRow> _stage2Rows = new();
        private string _clipboard = string.Empty;
        private bool _isUpdatingView;

        public TcpExportDialog(TcpExportModel model)
        {
            InitializeComponent();
            Model = model;
            _customContentFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CADRecognition", "tcp-custom-content.json");
            _tcpHistoryFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CADRecognition", "tcp-history.json");
            _customContentStore = LoadCustomContentStore(_customContentFilePath);
            _tcpHistoryStore = LoadTcpHistoryStore(_tcpHistoryFilePath);

            Loaded += TcpExportDialog_Loaded;
            BindModelToView();
            HookAutoSave(ProgramNameTextBox);
            HookAutoSave(ProgramNoTextBox);
            HookAutoSave(LeftRightDoorTextBox);
            HookAutoSave(MaterialTextBox);
            HookAutoSave(TypeTextBox);
            HookAutoSave(FormingLengthTextBox);
            HookAutoSave(FormingWidthTextBox);
            HookAutoSave(FormingThicknessTextBox);
            HookAutoSave(Stage1PunchCountTextBox);
            HookAutoSave(Stage2PunchCountTextBox);
            HookAutoSave(Spare2TextBox);
            HookAutoSave(PlateLengthTextBox);
            HookAutoSave(PlateWidthTextBox);
            HookAutoSave(PlateThicknessTextBox);
            HookAutoSave(Spare3TextBox);
            HookAutoSave(Spare4TextBox);

            LoadTcpHistoryToComboBoxes();
        }

        public TcpExportModel Model { get; }
        public string CustomContent => Model.CustomContent;

        private void HookAutoSave(TextBox textBox)
        {
            textBox.TextChanged += (_, __) =>
            {
                if (_isUpdatingView) return;
                OnFieldChanged();
            };
        }

        private void TcpExportDialog_Loaded(object sender, RoutedEventArgs e)
        {
            RefreshStageGrids();
            StageTabControl.SelectedIndex = _stage1Rows.Count > 0 ? 0 : 1;
        }

        private void LoadTcpHistoryToComboBoxes()
        {
            TcpHostComboBox.ItemsSource = _tcpHistoryStore.Hosts;
            TcpPortComboBox.ItemsSource = _tcpHistoryStore.Ports;

            if (!string.IsNullOrWhiteSpace(_tcpHistoryStore.LastHost)) TcpHostComboBox.Text = _tcpHistoryStore.LastHost;
            else TcpHostComboBox.Text = "127.0.0.1";

            if (!string.IsNullOrWhiteSpace(_tcpHistoryStore.LastPort)) TcpPortComboBox.Text = _tcpHistoryStore.LastPort;
            else TcpPortComboBox.Text = "9000";
        }

        private void SaveTcpHistory(string host, string port)
        {
            if (!string.IsNullOrWhiteSpace(host) && !_tcpHistoryStore.Hosts.Contains(host)) _tcpHistoryStore.Hosts.Insert(0, host);
            if (!string.IsNullOrWhiteSpace(port) && !_tcpHistoryStore.Ports.Contains(port)) _tcpHistoryStore.Ports.Insert(0, port);
            _tcpHistoryStore.LastHost = host;
            _tcpHistoryStore.LastPort = port;
            SaveTcpHistoryStore();
        }

        private void RefreshStageGrids()
        {
            _stage1Rows.Clear();
            _stage2Rows.Clear();
            foreach (var row in BuildStageRows(Model.Stage1DiagramCoordinates, Model.Stage1PositionMoldIds, Model.Stage1PunchMoldIds)) _stage1Rows.Add(row);
            foreach (var row in BuildStageRows(Model.Stage2DiagramCoordinates, Model.Stage2PositionMoldIds, Model.Stage2PunchMoldIds)) _stage2Rows.Add(row);
            Stage1Grid.ItemsSource = _stage1Rows;
            Stage2Grid.ItemsSource = _stage2Rows;
        }

        private static ObservableCollection<TcpGridRow> BuildStageRows(IReadOnlyList<TcpCoordinateRow> coords, IReadOnlyList<int> pos, IReadOnlyList<int> punch)
        {
            var count = Math.Max(coords.Count, Math.Max(pos.Count, punch.Count));
            var result = new ObservableCollection<TcpGridRow>();
            for (var i = 0; i < count; i++)
            {
                result.Add(new TcpGridRow
                {
                    RowIndex = i + 1,
                    X = i < coords.Count ? coords[i].X.ToString("0.###") : string.Empty,
                    Y = i < coords.Count ? coords[i].Y.ToString("0.###") : string.Empty,
                    PositionMoldId = i < pos.Count ? pos[i].ToString() : string.Empty,
                    PunchMoldId = i < punch.Count ? punch[i].ToString() : string.Empty
                });
            }
            return result;
        }

        private void BindModelToView()
        {
            _isUpdatingView = true;
            ProgramNameTextBox.Text = Model.ProgramName;
            ProgramNoTextBox.Text = Model.ProgramNo.ToString();
            LeftRightDoorTextBox.Text = Model.LeftRightDoor.ToString();
            MaterialTextBox.Text = Model.Material.ToString();
            TypeTextBox.Text = Model.Type.ToString();
            FormingLengthTextBox.Text = Model.FormingLength.ToString();
            FormingWidthTextBox.Text = Model.FormingWidth.ToString();
            FormingThicknessTextBox.Text = Model.FormingThickness.ToString();
            Stage1PunchCountTextBox.Text = Model.Stage1PunchCount.ToString();
            Stage2PunchCountTextBox.Text = Model.Stage2PunchCount.ToString();
            Spare2TextBox.Text = Model.Spare2.ToString();
            PlateLengthTextBox.Text = Model.PlateLength.ToString("0.###");
            PlateWidthTextBox.Text = Model.PlateWidth.ToString("0.###");
            PlateThicknessTextBox.Text = Model.PlateThickness.ToString("0.###");
            Spare3TextBox.Text = Model.Spare3.ToString("0.###");
            Spare4TextBox.Text = Model.Spare4.ToString();
            ApplyPermanentCustomFieldValues();
            _isUpdatingView = false;
        }

        private void ApplyPermanentCustomFieldValues()
        {
            SetIfSaved("ProgramNo", ProgramNoTextBox);
            SetIfSaved("LeftRightDoor", LeftRightDoorTextBox);
            SetIfSaved("Material", MaterialTextBox);
            SetIfSaved("Type", TypeTextBox);
            SetIfSaved("FormingLength", FormingLengthTextBox);
            SetIfSaved("FormingWidth", FormingWidthTextBox);
            SetIfSaved("FormingThickness", FormingThicknessTextBox);
            SetIfSaved("Spare2", Spare2TextBox);
            SetIfSaved("PlateThickness", PlateThicknessTextBox);
            SetIfSaved("Spare3", Spare3TextBox);
            SetIfSaved("Spare4", Spare4TextBox);
        }

        private void SetIfSaved(string key, TextBox box)
        {
            if (_customContentStore.Values.TryGetValue(key, out var saved) && !string.IsNullOrWhiteSpace(saved))
            {
                box.Text = saved;
            }
        }

        private void SavePermanentCustomFieldValues()
        {
            SaveIfCustom("ProgramNo", ProgramNoTextBox.Text);
            SaveIfCustom("LeftRightDoor", LeftRightDoorTextBox.Text);
            SaveIfCustom("Material", MaterialTextBox.Text);
            SaveIfCustom("Type", TypeTextBox.Text);
            SaveIfCustom("FormingLength", FormingLengthTextBox.Text);
            SaveIfCustom("FormingWidth", FormingWidthTextBox.Text);
            SaveIfCustom("FormingThickness", FormingThicknessTextBox.Text);
            SaveIfCustom("Spare2", Spare2TextBox.Text);
            SaveIfCustom("PlateThickness", PlateThicknessTextBox.Text);
            SaveIfCustom("Spare3", Spare3TextBox.Text);
            SaveIfCustom("Spare4", Spare4TextBox.Text);
        }

        private void SaveIfCustom(string key, string? value)
        {
            _customContentStore.Values[key] = value?.Trim() ?? string.Empty;
        }

        private static TcpCustomContentStore LoadCustomContentStore(string filePath)
        {
            try
            {
                if (!File.Exists(filePath)) return new TcpCustomContentStore();
                var json = File.ReadAllText(filePath);
                return JsonSerializer.Deserialize<TcpCustomContentStore>(json) ?? new TcpCustomContentStore();
            }
            catch
            {
                return new TcpCustomContentStore();
            }
        }

        private static TcpConnectionHistoryStore LoadTcpHistoryStore(string filePath)
        {
            try
            {
                if (!File.Exists(filePath)) return new TcpConnectionHistoryStore();
                var json = File.ReadAllText(filePath);
                return JsonSerializer.Deserialize<TcpConnectionHistoryStore>(json) ?? new TcpConnectionHistoryStore();
            }
            catch
            {
                return new TcpConnectionHistoryStore();
            }
        }

        private void SaveTcpHistoryStore()
        {
            try
            {
                var dir = Path.GetDirectoryName(_tcpHistoryFilePath);
                if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
                var json = JsonSerializer.Serialize(_tcpHistoryStore, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_tcpHistoryFilePath, json);
            }
            catch
            {
            }
        }

        private void SaveCustomContentStore()
        {
            try
            {
                var dir = Path.GetDirectoryName(_customContentFilePath);
                if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
                var json = JsonSerializer.Serialize(_customContentStore, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_customContentFilePath, json);
            }
            catch
            {
            }
        }

        private void OnFieldChanged()
        {
            UpdateModelFromRows();
            StatusTextBlock.Text = "已自动保存";
        }

        private void UpdateModelFromRows()
        {
            Model.ProgramName = ProgramNameTextBox.Text?.Trim() ?? string.Empty;
            Model.ProgramNo = ParseInt(ProgramNoTextBox.Text);
            Model.LeftRightDoor = ParseInt(LeftRightDoorTextBox.Text);
            Model.Material = ParseInt(MaterialTextBox.Text);
            Model.Type = ParseInt(TypeTextBox.Text);
            Model.FormingLength = ParseInt(FormingLengthTextBox.Text);
            Model.FormingWidth = ParseInt(FormingWidthTextBox.Text);
            Model.FormingThickness = ParseInt(FormingThicknessTextBox.Text);
            Model.Stage1PunchCount = ParseInt(Stage1PunchCountTextBox.Text);
            Model.Stage2PunchCount = ParseInt(Stage2PunchCountTextBox.Text);
            Model.Spare2 = ParseInt(Spare2TextBox.Text);
            Model.PlateLength = ParseDouble(PlateLengthTextBox.Text);
            Model.PlateWidth = ParseDouble(PlateWidthTextBox.Text);
            Model.PlateThickness = ParseDouble(PlateThicknessTextBox.Text);
            Model.Spare3 = ParseDouble(Spare3TextBox.Text);
            Model.Spare4 = ParseInt(Spare4TextBox.Text);

            SavePermanentCustomFieldValues();
            SaveCustomContentStore();
        }

        private async void SendTcp_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var host = TcpHostComboBox.Text?.Trim();
                var portText = TcpPortComboBox.Text?.Trim();
                if (!int.TryParse(portText, out var port))
                {
                    StatusTextBlock.Text = "TCP 端口格式不正确。";
                    return;
                }

                SaveTcpHistory(host ?? string.Empty, portText ?? string.Empty);
                await _tcpCommService.SendJsonAsync(host ?? string.Empty, port, Model);
                StatusTextBlock.Text = $"TCP 已发送到 {host}:{port}。";
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"TCP 发送失败：{ex.Message}";
            }
        }

        private static int ParseInt(string? text) => int.TryParse(text, out var value) ? value : 0;
        private static double ParseDouble(string? text) => double.TryParse(text, out var value) ? value : 0;

        private void InsertRow(ObservableCollection<TcpGridRow> rows, System.Windows.Controls.DataGrid grid)
        {
            var index = grid.SelectedIndex < 0 ? rows.Count : grid.SelectedIndex;
            rows.Insert(index, new TcpGridRow { RowIndex = index + 1 });
            RenumberRows(rows);
            grid.Items.Refresh();
        }

        private void DeleteRow(ObservableCollection<TcpGridRow> rows, System.Windows.Controls.DataGrid grid)
        {
            if (grid.SelectedItem is not TcpGridRow row) return;
            var idx = row.RowIndex - 1;
            if (idx < 0 || idx >= rows.Count) return;
            rows.RemoveAt(idx);
            RenumberRows(rows);
            grid.Items.Refresh();
        }

        private void CopyRow(System.Windows.Controls.DataGrid grid)
        {
            if (grid.SelectedItem is not TcpGridRow row) return;
            _clipboard = string.Join("\t", row.RowIndex, row.X, row.Y, row.PositionMoldId, row.PunchMoldId);
        }

        private void PasteRow(ObservableCollection<TcpGridRow> rows, System.Windows.Controls.DataGrid grid)
        {
            if (string.IsNullOrWhiteSpace(_clipboard)) return;
            var row = grid.SelectedItem as TcpGridRow;
            if (row is null)
            {
                row = new TcpGridRow { RowIndex = rows.Count + 1 };
                rows.Add(row);
            }
            var parts = _clipboard.Split('\t');
            while (parts.Length < 5) Array.Resize(ref parts, 5);
            row.X = parts[1];
            row.Y = parts[2];
            row.PositionMoldId = parts[3];
            row.PunchMoldId = parts[4];
            RenumberRows(rows);
            grid.Items.Refresh();
        }

        private static void RenumberRows(ObservableCollection<TcpGridRow> rows)
        {
            for (var i = 0; i < rows.Count; i++) rows[i].RowIndex = i + 1;
        }

        private void Stage1Insert_Click(object sender, RoutedEventArgs e) => InsertRow(_stage1Rows, Stage1Grid);
        private void Stage1Delete_Click(object sender, RoutedEventArgs e) => DeleteRow(_stage1Rows, Stage1Grid);
        private void Stage1Copy_Click(object sender, RoutedEventArgs e) => CopyRow(Stage1Grid);
        private void Stage1Paste_Click(object sender, RoutedEventArgs e) => PasteRow(_stage1Rows, Stage1Grid);
        private void Stage2Insert_Click(object sender, RoutedEventArgs e) => InsertRow(_stage2Rows, Stage2Grid);
        private void Stage2Delete_Click(object sender, RoutedEventArgs e) => DeleteRow(_stage2Rows, Stage2Grid);
        private void Stage2Copy_Click(object sender, RoutedEventArgs e) => CopyRow(Stage2Grid);
        private void Stage2Paste_Click(object sender, RoutedEventArgs e) => PasteRow(_stage2Rows, Stage2Grid);

        private void Stage1Grid_PreviewKeyDown(object sender, KeyEventArgs e) => HandleGridKeyDown(e, Stage1Grid, _stage1Rows);
        private void Stage2Grid_PreviewKeyDown(object sender, KeyEventArgs e) => HandleGridKeyDown(e, Stage2Grid, _stage2Rows);

        private void HandleGridKeyDown(KeyEventArgs e, System.Windows.Controls.DataGrid grid, ObservableCollection<TcpGridRow> rows)
        {
            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == System.Windows.Input.Key.C)
            {
                CopyRow(grid);
                e.Handled = true;
                return;
            }

            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == System.Windows.Input.Key.V)
            {
                PasteRow(rows, grid);
                e.Handled = true;
                return;
            }

            if (e.Key == System.Windows.Input.Key.Delete)
            {
                DeleteRow(rows, grid);
                e.Handled = true;
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            UpdateModelFromRows();
            DialogResult = true;
            Close();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
