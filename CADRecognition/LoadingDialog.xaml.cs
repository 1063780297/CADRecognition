using System.Windows;

namespace CADRecognition
{
    public partial class LoadingDialog : Window
    {
        public LoadingDialog()
        {
            InitializeComponent();
        }

        public void SetMessage(string message)
        {
            Dispatcher.Invoke(() =>
            {
                LoadingText.Text = message;
            });
        }

        public void SetDetail(string detail)
        {
            Dispatcher.Invoke(() =>
            {
                DetailText.Text = detail;
            });
        }
    }
}
