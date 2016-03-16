
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using YUBI_TOOL.Model;
using YUBI_TOOL.Common;
using System.Text.RegularExpressions;
using System.Data.SqlClient;

namespace YUBI_TOOL.ViewModel
{
    public class DBServerViewModel : ViewModelBase
    {
        private const string FORM_ID = "DBServerMante";
        private const string DBTYPE_WINDOWS_AUTHENTICATION = "0";
        private const string DBTYPE_SQL_SERVER_AUTHENTICATION = "1";
        private List<SelectItemModel> authenticationTypeList;
        private string selectedAuthenthicationType;
        private string userName;
        private string password;
        private string serverName;
        private bool isFocused;
        private LanguageModel lblTitle;
        private LanguageModel lblServerName;
        private LanguageModel lblAuthentication;
        private LanguageModel lblUserName;
        private LanguageModel lblPassword;
        private LanguageModel lblMessageArea;
        private LanguageModel lblMessageText;
        private LanguageModel btnTest;
        private LanguageModel btnSave;
        private LanguageModel btnCancel;
        private LanguageModel btnClose;
        private LanguageModel lblGrpInput;
        public bool canEditUserNameAndPassword;

        public DBServerViewModel()
        {
            this.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(DBServerViewModel_PropertyChanged);
        }

        private void DBServerViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == GetPropertyName<DBServerViewModel, string>(x => x.SelectedAuthenthicationType))
            {
                NotifyOfPropertyChange(() => CanEditUserNameAndPassword);
            }
        }
        protected override void SetMultiLanguage()
        {
            LanguageModel form = ResourcesManager.GetLanguageForForm(FORM_ID);
            var display = ResourcesManager.GetLanguageForControlInForm(form, "Text");
            if (display != null)
            {
                SetDisplayName(display.Text);
            }

            LblTitle = ResourcesManager.GetLanguageForControlInForm(form, "Title");
            LblMessageArea = ResourcesManager.GetLanguageForControlInForm(form, "MessageArea");
            LblMessageText = ResourcesManager.GetLanguageForControlInForm(form, "MessageText");
            LblServerName = ResourcesManager.GetLanguageForControlInForm(form, "lblDBServerName");
            LblAuthentication = ResourcesManager.GetLanguageForControlInForm(form, "lblAuthentication");
            LblUserName = ResourcesManager.GetLanguageForControlInForm(form, "lblUserName");
            LblPassword = ResourcesManager.GetLanguageForControlInForm(form, "lblPassword");
            BtnTest = ResourcesManager.GetLanguageForControlInForm(form, "btnTest");
            BtnSave = ResourcesManager.GetLanguageForControlInForm(form, "btnUpdate");
            BtnCancel = ResourcesManager.GetLanguageForControlInForm(form, "btnCancel");

            BtnClose = ResourcesManager.GetLanguageForControlInForm(form, "btnEnd");
            LblGrpInput = ResourcesManager.GetLanguageForControlInForm(form, "grpInput");
            Message = null;

            var rdoWindowsAuthentication = ResourcesManager.GetLanguageForControlInForm(form, "rdoWindowsAuthentication");
            var rdoSQLServerAuthentication = ResourcesManager.GetLanguageForControlInForm(form, "rdoSQLServerAuthentication");
            AuthenticationTypeList = new List<SelectItemModel>()
            {
                new SelectItemModel() {
                    ItemCD = DBTYPE_WINDOWS_AUTHENTICATION,
                    ItemValue = rdoWindowsAuthentication.Text
                },
                new SelectItemModel() {
                    ItemCD = DBTYPE_SQL_SERVER_AUTHENTICATION,
                    ItemValue = rdoSQLServerAuthentication.Text
                },
            };
        }

        public void TestConnection()
        {
            System.Data.SqlClient.SqlConnectionStringBuilder builder = new System.Data.SqlClient.SqlConnectionStringBuilder();
            builder.DataSource = ServerName;
            builder.InitialCatalog = "YUBITARO";
            builder.IntegratedSecurity = true;
            if (CanEditUserNameAndPassword)
            {
                builder.UserID = userName;
                builder.Password = Password;
            }
            if (TestConnection(builder.ConnectionString))
            {
                Message = ResourcesManager.GetMessage(MessageConstant.I0035);
            }
            else
            {
                Message = ResourcesManager.GetMessage(MessageConstant.I0034, new object[] { serverName });
            }
           
        }

       
        public void Cancel()
        {
            TryClose();
        }
        public void Save()
        {
            System.Data.SqlClient.SqlConnectionStringBuilder builder = new System.Data.SqlClient.SqlConnectionStringBuilder();
            builder.DataSource = ServerName;
            builder.InitialCatalog = "YUBITARO";
            builder.IntegratedSecurity = true;
            if (CanEditUserNameAndPassword)
            {
                builder.UserID = userName;
                builder.Password = Password;
            }
            if (TestConnection(builder.ConnectionString))
            {
                Common.ApplicationSettingsWriter writer = new ApplicationSettingsWriter();
                writer.ChangeConnectionStrings(builder.ConnectionString);
                Properties.Settings.Default.Is_DB_Configed = true;
                Properties.Settings.Default.Save();
                Cancel();
            }
            else
            {
                Message = ResourcesManager.GetMessage(MessageConstant.I0034, new object[] { serverName });
            }
        }
        protected override void OnActivate()
        {
            base.OnActivate();
            Init();
        }

        public bool CanUserName
        {
            get { return false; }
        }
        private void Init()
        {

            System.Data.SqlClient.SqlConnectionStringBuilder builder = new System.Data.SqlClient.SqlConnectionStringBuilder();
            builder.ConnectionString = Properties.Settings.Default.YUBITAROConnectionString;
            ServerName = builder.DataSource;
            UserName = builder.UserID;
            Password = builder.Password;
            if (string.IsNullOrEmpty(userName))
            {
                SelectedAuthenthicationType = DBTYPE_WINDOWS_AUTHENTICATION;
            }
            else
            {
                SelectedAuthenthicationType = DBTYPE_SQL_SERVER_AUTHENTICATION;
            }

        }
        #region get/set
        public List<SelectItemModel> AuthenticationTypeList
        {
            get
            {
                return authenticationTypeList;
            }
            set
            {
                if (authenticationTypeList != value)
                {
                    authenticationTypeList = value;
                    NotifyOfPropertyChange(() => AuthenticationTypeList);
                }
            }
        }

        public string SelectedAuthenthicationType
        {
            get
            {
                return selectedAuthenthicationType;
            }
            set
            {
                if (selectedAuthenthicationType != value)
                {
                    selectedAuthenthicationType = value;
                    NotifyOfPropertyChange(() => SelectedAuthenthicationType);
                }
            }
        }

        public string UserName
        {
            get
            {
                return userName;
            }
            set
            {
                if (userName != value)
                {
                    userName = value;
                    NotifyOfPropertyChange(() => UserName);
                }
            }
        }

        public string Password
        {
            get
            {
                return password;
            }
            set
            {
                if (password != value)
                {
                    password = value;
                    NotifyOfPropertyChange(() => Password);
                }
            }
        }

        public string ServerName
        {
            get
            {
                return serverName;
            }
            set
            {
                if (serverName != value)
                {
                    serverName = value;
                    NotifyOfPropertyChange(() => ServerName);
                }
            }
        }

        public bool IsFocused
        {
            get
            {
                return isFocused;
            }
            set
            {
                if (isFocused != value)
                {
                    isFocused = value;
                    NotifyOfPropertyChange(() => IsFocused);
                }
            }
        }

        public LanguageModel LblServerName
        {
            get
            {
                return lblServerName;
            }
            set
            {
                if (lblServerName != value)
                {
                    lblServerName = value;
                    NotifyOfPropertyChange(() => LblServerName);
                }
            }
        }

        public LanguageModel LblAuthentication
        {
            get
            {
                return lblAuthentication;
            }
            set
            {
                if (lblAuthentication != value)
                {
                    lblAuthentication = value;
                    NotifyOfPropertyChange(() => LblAuthentication);
                }
            }
        }

        public LanguageModel LblUserName
        {
            get
            {
                return lblUserName;
            }
            set
            {
                if (lblUserName != value)
                {
                    lblUserName = value;
                    NotifyOfPropertyChange(() => LblUserName);
                }
            }
        }

        public LanguageModel LblPassword
        {
            get
            {
                return lblPassword;
            }
            set
            {
                if (lblPassword != value)
                {
                    lblPassword = value;
                    NotifyOfPropertyChange(() => LblPassword);
                }
            }
        }

        public LanguageModel LblMessageArea
        {
            get
            {
                return lblMessageArea;
            }
            set
            {
                if (lblMessageArea != value)
                {
                    lblMessageArea = value;
                    NotifyOfPropertyChange(() => LblMessageArea);
                }
            }
        }

        public LanguageModel LblMessageText
        {
            get
            {
                return lblMessageText;
            }
            set
            {
                if (lblMessageText != value)
                {
                    lblMessageText = value;
                    NotifyOfPropertyChange(() => LblMessageText);
                }
            }
        }

        public LanguageModel BtnTest
        {
            get
            {
                return btnTest;
            }
            set
            {
                if (btnTest != value)
                {
                    btnTest = value;
                    NotifyOfPropertyChange(() => BtnTest);
                }
            }
        }

        public LanguageModel BtnSave
        {
            get
            {
                return btnSave;
            }
            set
            {
                if (btnSave != value)
                {
                    btnSave = value;
                    NotifyOfPropertyChange(() => BtnSave);
                }
            }
        }

        public LanguageModel BtnClose
        {
            get
            {
                return btnClose;
            }
            set
            {
                if (btnClose != value)
                {
                    btnClose = value;
                    NotifyOfPropertyChange(() => BtnClose);
                }
            }
        }

        public LanguageModel LblGrpInput
        {
            get
            {
                return lblGrpInput;
            }
            set
            {
                if (lblGrpInput != value)
                {
                    lblGrpInput = value;
                    NotifyOfPropertyChange(() => LblGrpInput);
                }
            }
        }
        public LanguageModel LblTitle
        {
            get
            {
                return lblTitle;
            }
            set
            {
                if (lblTitle != value)
                {
                    lblTitle = value;
                    NotifyOfPropertyChange(() => LblTitle);
                }
            }
        }
        public LanguageModel BtnCancel
        {
            get
            {
                return btnCancel;
            }
            set
            {
                if (btnCancel != value)
                {
                    btnCancel = value;
                    NotifyOfPropertyChange(() => BtnCancel);
                }
            }
        }
        public bool CanEditUserNameAndPassword
        {
            get
            {
                return selectedAuthenthicationType == DBTYPE_SQL_SERVER_AUTHENTICATION;
            }

        }



        #endregion
    }
}
