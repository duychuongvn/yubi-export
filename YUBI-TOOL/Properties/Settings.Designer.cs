﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.261
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace YUBI_TOOL.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "10.0.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Dailly_Report_V1_{0}.xls")]
        public string XLS_Daily_Report {
            get {
                return ((string)(this["XLS_Daily_Report"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Monthly_Report_V1_{0}.xls")]
        public string XLS_Monthly_Report {
            get {
                return ((string)(this["XLS_Monthly_Report"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Personal_Report_V1_{0}.xls")]
        public string XLS_Personal_Report {
            get {
                return ((string)(this["XLS_Personal_Report"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string XLS_Export_Path {
            get {
                return ((string)(this["XLS_Export_Path"]));
            }
            set {
                this["XLS_Export_Path"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("True")]
        public bool XLS_Use_Multi_Language {
            get {
                return ((bool)(this["XLS_Use_Multi_Language"]));
            }
            set {
                this["XLS_Use_Multi_Language"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Personal Report, {0}, {1}.xls")]
        public string XLS_Out_Personal_File {
            get {
                return ((string)(this["XLS_Out_Personal_File"]));
            }
            set {
                this["XLS_Out_Personal_File"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Daily Report, {0}, {1}.xls")]
        public string XLS_Out_Daily_File {
            get {
                return ((string)(this["XLS_Out_Daily_File"]));
            }
            set {
                this["XLS_Out_Daily_File"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Monthly Report, {0}, {1}.xls")]
        public string XLS_Out_Monthly_File {
            get {
                return ((string)(this["XLS_Out_Monthly_File"]));
            }
            set {
                this["XLS_Out_Monthly_File"] = value;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.ConnectionString)]
        [global::System.Configuration.DefaultSettingValueAttribute("Data Source=localhost;Initial Catalog=YUBITARO;Integrated Security=True")]
        public string YUBITAROConnectionString {
            get {
                return ((string)(this["YUBITAROConnectionString"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("English")]
        public string SelectedLanguage {
            get {
                return ((string)(this["SelectedLanguage"]));
            }
            set {
                this["SelectedLanguage"] = value;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("True")]
        public bool AutoCorrectThirdShift {
            get {
                return ((bool)(this["AutoCorrectThirdShift"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Correct_Night_Shift_{0}.xls")]
        public string XLS_Correct_Night_Shift {
            get {
                return ((string)(this["XLS_Correct_Night_Shift"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("False")]
        public bool Is_DB_Configed {
            get {
                return ((bool)(this["Is_DB_Configed"]));
            }
            set {
                this["Is_DB_Configed"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Sheet1")]
        public string XLS_Default_Sheet_Name {
            get {
                return ((string)(this["XLS_Default_Sheet_Name"]));
            }
            set {
                this["XLS_Default_Sheet_Name"] = value;
            }
        }
    }
}
