using System.Configuration;
using System.Data;
using System.Windows;
using Application = System.Windows.Application;
using HslCommunication;

namespace CADRecognition
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            Authorization.SetAuthorizationCode(HslCommunicationAppLicense.AuthorizationCode);
            base.OnStartup(e);
        }
    }

}
