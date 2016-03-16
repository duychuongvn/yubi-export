using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.ComponentModel.Composition.Primitives;
using System.Data.SqlClient;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;
using Caliburn.Micro;
namespace YUBI_TOOL
{
    public class AppBootstrapper : Bootstrapper<IShell>
    {

        private CompositionContainer _container;
        protected override void Configure()
        {

            var catalog = new AggregateCatalog(
                AssemblySource.Instance.Select(x => new AssemblyCatalog(x)).OfType<ComposablePartCatalog>()
                );

            _container = new CompositionContainer(catalog);
            var batch = new CompositionBatch();

            batch.AddExportedValue<IWindowManager>(new WindowManager());
            batch.AddExportedValue<IEventAggregator>(new EventAggregator());
            batch.AddExportedValue(_container);
            batch.AddExportedValue(catalog);
            _container.Compose(batch);
            var originalInvoke = ActionMessage.InvokeAction;
            ActionMessage.InvokeAction = context =>
            {
                Mouse.OverrideCursor = Cursors.Wait;
                try
                {
                    originalInvoke(context);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                Mouse.OverrideCursor = null;
            };

            if ( Properties.Settings.Default.Is_DB_Configed || TestConnection(Properties.Settings.Default.YUBITAROConnectionString))
            {
                Properties.Settings.Default.Is_DB_Configed = true;
                Properties.Settings.Default.Save();
            }
            else
            {
                Properties.Settings.Default.Is_DB_Configed = false;
                Properties.Settings.Default.Save();
            }
          
            //XmlConfigurator.Configure();
            Common.CommonUtil.ClearTemplate();
            ConventionManager.ApplyValidation = (binding, viewModelType, property) =>
            {
                binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
                binding.ValidatesOnExceptions = false;
                binding.ValidatesOnDataErrors = true;
            };
        }

        protected override object GetInstance(Type serviceType, string key)
        {
            string contract = string.IsNullOrEmpty(key) ? AttributedModelServices.GetContractName(serviceType) : key;
            var exports = _container.GetExportedValues<object>(contract);

            if (exports.Count() > 0)
                return exports.First();

            throw new Exception(string.Format("Could not locate any instances of contract {0}.", contract));
        }
        protected override IEnumerable<object> GetAllInstances(Type serviceType)
        {
            return _container.GetExportedValues<object>(AttributedModelServices.GetContractName(serviceType));
        }
        protected override void BuildUp(object instance)
        {
            _container.SatisfyImportsOnce(instance);
        }
        protected override void OnUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            
        }
        protected override void StartRuntime()
        {


            Thread newWindowThread = new Thread(new ThreadStart(LoadSplashScreen));
            newWindowThread.SetApartmentState(ApartmentState.STA);
            newWindowThread.IsBackground = true;
            newWindowThread.Start();
            base.StartRuntime();
            newWindowThread.Abort();
        }
        private void LoadSplashScreen()
        {
            Common.SplashScreen.SplashScreen splashScreen = new Common.SplashScreen.SplashScreen();
            splashScreen.Show();
            System.Windows.Threading.Dispatcher.Run();
        }
        private bool TestConnection(string connectString)
        {
            using (SqlConnection connection = new SqlConnection(connectString))
            {
                try
                {
                    connection.Open();
                    return true;
                }
                catch (SqlException)
                {
                    return false;
                }
                finally
                {
                    // not really necessary
                    connection.Close();
                }
            }
        }
    }
}
