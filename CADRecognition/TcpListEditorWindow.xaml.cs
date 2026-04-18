using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace CADRecognition
{
    public partial class TcpListEditorWindow : Window
    {
        private readonly bool _isStage1;
        private readonly TcpExportModel _model;
        private readonly ObservableCollection<TcpGridRow> _rows;
        private string _clipboard = string.Empty;

        public TcpListEditorWindow(bool isStage1, TcpExportModel model)
        {
            InitializeComponent();
            _isStage1 = isStage1;
            _model = model;
            TitleText.Text = isStage1 ? "编辑：台1合并数据表" : "编辑：台2合并数据表";
            _rows = LoadRows(isStage1, model);
            GridView.ItemsSource = _rows;
            GridView.PreviewKeyDown += GridView_PreviewKeyDown;
            GridView.CellEditEnding += GridView_CellEditEnding;
        }

        public ObservableCollection<TcpGridRow> Rows => _rows;

        private static ObservableCollection<TcpGridRow> LoadRows(bool isStage1, TcpExportModel model)
        {
            var coords = isStage1 ? model.Stage1DiagramCoordinates : model.Stage2DiagramCoordinates;
            var pos = isStage1 ? model.Stage1PositionMoldIds : model.Stage2PositionMoldIds;
            var punch = isStage1 ? model.Stage1PunchMoldIds : model.Stage2PunchMoldIds;
            var count = new[] { coords.Count, pos.Count, punch.Count }.Max();

            var rows = new ObservableCollection<TcpGridRow>();
            for (var i = 0; i < count; i++)
            {
                rows.Add(new TcpGridRow
                {
                    RowIndex = i + 1,
                    X = i < coords.Count ? coords[i].X.ToString("0.###") : "0",
                    Y = i < coords.Count ? coords[i].Y.ToString("0.###") : "0",
                    PositionMoldId = i < pos.Count ? pos[i].ToString() : "0",
                    PunchMoldId = i < punch.Count ? punch[i].ToString() : "0"
                });
            }
            return rows;
        }

        private void GridView_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
        }

        private void GridView_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                DeleteRow_Click(sender, e);
                e.Handled = true;
            }
        }

        private void InsertRow_Click(object sender, RoutedEventArgs e)
        {
            var index = GridView.SelectedIndex < 0 ? _rows.Count : GridView.SelectedIndex;
            _rows.Insert(index, new TcpGridRow { RowIndex = index + 1, X = "0", Y = "0", PositionMoldId = "0", PunchMoldId = "0" });
            RenumberRows();
        }

        private void DeleteRow_Click(object sender, RoutedEventArgs e)
        {
            if (GridView.SelectedItem is not TcpGridRow row) return;
            var idx = row.RowIndex - 1;
            if (idx < 0 || idx >= _rows.Count) return;
            _rows.RemoveAt(idx);
            _rows.Add(new TcpGridRow { X = "0", Y = "0", PositionMoldId = "0", PunchMoldId = "0" });
            RenumberRows();
        }

        private void Copy_Click(object sender, RoutedEventArgs e)
        {
            if (GridView.SelectedItem is not TcpGridRow row) return;
            _clipboard = string.Join(",", row.X, row.Y, row.PositionMoldId, row.PunchMoldId);
        }

        private void Paste_Click(object sender, RoutedEventArgs e)
        {
            if (GridView.SelectedItem is not TcpGridRow row || string.IsNullOrWhiteSpace(_clipboard)) return;
            var parts = _clipboard.Split(',');
            while (parts.Length < 4) Array.Resize(ref parts, 4);
            row.X = parts[0];
            row.Y = parts[1];
            row.PositionMoldId = parts[2];
            row.PunchMoldId = parts[3];
            GridView.Items.Refresh();
        }

        private void RenumberRows()
        {
            for (var i = 0; i < _rows.Count; i++) _rows[i].RowIndex = i + 1;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            var coords = new System.Collections.Generic.List<TcpCoordinateRow>();
            var pos = new System.Collections.Generic.List<int>();
            var punch = new System.Collections.Generic.List<int>();

            foreach (var r in _rows)
            {
                var x = ParseDoubleOrZero(r.X);
                var y = ParseDoubleOrZero(r.Y);
                var p = ParseIntOrZero(r.PositionMoldId);
                var q = ParseIntOrZero(r.PunchMoldId);

                if (x == 0 && y == 0 && p == 0 && q == 0)
                {
                    continue;
                }

                coords.Add(new TcpCoordinateRow { X = x, Y = y });
                pos.Add(p);
                punch.Add(q);
            }

            if (_isStage1)
            {
                _model.Stage1DiagramCoordinates = coords;
                _model.Stage1PositionMoldIds = pos;
                _model.Stage1PunchMoldIds = punch;
            }
            else
            {
                _model.Stage2DiagramCoordinates = coords;
                _model.Stage2PositionMoldIds = pos;
                _model.Stage2PunchMoldIds = punch;
            }

            DialogResult = true;
            Close();
        }

        private static int ParseIntOrZero(string? text) => int.TryParse(text, out var value) ? value : 0;
        private static double ParseDoubleOrZero(string? text) => double.TryParse(text, out var value) ? value : 0;
    }

    public sealed class TcpGridRow
    {
        public int RowIndex { get; set; }
        public string X { get; set; } = string.Empty;
        public string Y { get; set; } = string.Empty;
        public string PositionMoldId { get; set; } = string.Empty;
        public string PunchMoldId { get; set; } = string.Empty;
    }
}
