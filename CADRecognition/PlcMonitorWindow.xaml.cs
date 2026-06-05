using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Microsoft.Win32;

namespace CADRecognition
{
    public partial class PlcMonitorWindow : Window
    {
        private readonly MainWindow _owner;
        private readonly DispatcherTimer _refreshTimer;

        public PlcMonitorWindow(MainWindow owner)
        {
            InitializeComponent();
            _owner = owner;
            DataContext = _owner;
            FolderText.Text = _owner.CurrentProjectFolder ?? "未选择";

            _refreshTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(500)
            };
            _refreshTimer.Tick += RefreshTimer_Tick;
            Loaded += PlcMonitorWindow_Loaded;
            Unloaded += PlcMonitorWindow_Unloaded;
        }

        private async void PlcMonitorWindow_Loaded(object sender, RoutedEventArgs e)
        {
            await _owner.LoadPlcRegistersForMonitorAsync().ConfigureAwait(true);
            _refreshTimer.Start();
        }

        private void PlcMonitorWindow_Unloaded(object sender, RoutedEventArgs e)
        {
            _refreshTimer.Stop();
        }

        private async void RefreshTimer_Tick(object? sender, EventArgs e)
        {
            if (!IsVisible)
            {
                _refreshTimer.Stop();
                return;
            }

            FolderText.Text = _owner.CurrentProjectFolder ?? "未选择";
            await _owner.LoadPlcRegistersForMonitorAsync().ConfigureAwait(true);
        }

        private void ChooseFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "请选择图纸文件所在文件夹",
                SelectedPath = _owner.CurrentProjectFolder ?? string.Empty,
                ShowNewFolderButton = false
            };

            if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK || string.IsNullOrWhiteSpace(dialog.SelectedPath))
            {
                return;
            }

            _owner.SetProjectFolder(dialog.SelectedPath);
            FolderText.Text = dialog.SelectedPath;
        }

        private void PlcRegisterGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            _owner.HandlePlcRegisterCellEditEnding(sender, e);
        }
    }
}
