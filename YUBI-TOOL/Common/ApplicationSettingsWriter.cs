using System.Configuration;

namespace YUBI_TOOL.Common
{
    /// <summary>
    /// Change application config with aplication scope
    /// </summary>
    public class ApplicationSettingsWriter
    {
        /// <summary>
        /// Change connectionString
        /// </summary>
        /// <param name="connectionString"></param>
        public void ChangeConnectionStrings(string connectionString)
        {
            // Get the application configuration file.
            System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
           
            string connectionName = "YUBI_TOOL.Properties.Settings.YUBITAROConnectionString";
            ConnectionStringSettings csSettings = new ConnectionStringSettings(connectionName, connectionString);
            ConnectionStringsSection csSection = config.ConnectionStrings;

            while (csSection.ConnectionStrings.Count > 0)
            {
                csSection.ConnectionStrings.RemoveAt(0);
            }
            csSection.ConnectionStrings.Add(csSettings);
           
            config.Save(ConfigurationSaveMode.Modified, true);
            ConfigurationManager.RefreshSection("connectionStrings");
            Properties.Settings.Default.Reload();
        }
    }
}
