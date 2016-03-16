using System.Threading;
using System.Windows;

namespace YUBI_TOOL
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
           
            Mutex mutex = new Mutex(false, "YUBI-GROUP");
            if (mutex.WaitOne(0, false))
            {
                try
                {
                    base.OnStartup(e);
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }
            else
            {
                this.Shutdown();
            }
            
        }
    }

}
