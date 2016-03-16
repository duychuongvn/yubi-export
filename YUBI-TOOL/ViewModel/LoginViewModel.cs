using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Caliburn.Micro;
using System.ComponentModel.Composition;
using YUBI_TOOL.Model;
using YUBI_TOOL.Service;
using YUBI_TOOL.Common;
namespace YUBI_TOOL.ViewModel
{
    [Export(typeof(LoginViewModel))]
    public class LoginViewModel : ViewModelBase
    {
        private const string FORM_ID = "Login";
        private IEventAggregator eventAggregator;
        private ICompanyService companyService;

        private List<CompanyModel> companyList;
        private string userName;
        private string password;
        private bool isFocused;
        private LanguageModel lblLogin;
        private LanguageModel lblCompany;
        private LanguageModel lblUserName;
        private LanguageModel lblPassword;
        private LanguageModel lblMessageArea;
        private LanguageModel lblMessageText;
        private LanguageModel btnDBServer;
        private LanguageModel btnLogin;
        private LanguageModel btnClose;
        private LanguageModel lblGrpInput;
        private bool canLogin;


        private List<string> languageList = new List<string>() { "English", "Japanese", "Vietnamese" };


        [ImportingConstructor()]
        public LoginViewModel(IEventAggregator eventAggregator)
        {
            this.eventAggregator = eventAggregator;
            this.companyService = IoC.Get<ICompanyService>();
        }

        protected override void OnActivate()
        {
            base.OnActivate();
            UserName = null;
            Password = null;

            SetMultiLanguage();
            try
            {
                if (Properties.Settings.Default.Is_DB_Configed)
                {
                    CompanyList = companyService.SearchCompanyList();
                    CanLogin = true;
                }
                else
                {
                    CanLogin = false;
                }
            }
            catch
            {
                CanLogin = false;
            }
            IsFocused = true;
        }

        public void ChangeLanguage(string language)
        {
            if (language != Properties.Settings.Default.SelectedLanguage)
            {
                Properties.Settings.Default.SelectedLanguage = language;
                Properties.Settings.Default.Save();
                SetMultiLanguage();
            }
        }

        public void UpdateDB()
        {
            ShellViewModel shellViewModel = (ShellViewModel)IoC.Get<IShell>();

            shellViewModel.ActiveDBServerMainte();
        }
        public void Login()
        {
            if (string.IsNullOrEmpty(userName))
            {
                Message = ResourcesManager.GetMessage(MessageConstant.A0021, new string[] { CommonUtil.GetCaption(lblUserName.Text) });
                return;
            }
            else if (string.IsNullOrEmpty(Password))
            {
                Message = ResourcesManager.GetMessage(MessageConstant.A0021, new string[] { CommonUtil.GetCaption(lblPassword.Text) });
                return;
            }
            var autheticationManager = IoC.Get<Security.IAuthenticationManager>();
            autheticationManager.DoAuthentication(userName, password);
            if (autheticationManager.IsAuthenticated())
            {
                ShellViewModel shellViewModel = (ShellViewModel)IoC.Get<IShell>();
                TryClose(true);
                shellViewModel.ActiveMainWindow();
            }
            else
            {
                Message = ResourcesManager.GetMessage(MessageConstant.A0023);
            }
        }

        protected override void SetMultiLanguage()
        {
            LanguageModel loginForm = ResourcesManager.GetLanguageForForm(FORM_ID);
            var display = ResourcesManager.GetLanguageForControlInForm(loginForm, "Text");
            if (display != null)
            {
                SetDisplayName(display.Text);
            }
            LblLogin = ResourcesManager.GetLanguageForControlInForm(loginForm, "Title");
            LblMessageArea = ResourcesManager.GetLanguageForControlInForm(loginForm, "MessageArea");
            LblMessageText = ResourcesManager.GetLanguageForControlInForm(loginForm, "MessageText");
            LblCompany = ResourcesManager.GetLanguageForControlInForm(loginForm, "lblCompany");
            LblUserName = ResourcesManager.GetLanguageForControlInForm(loginForm, "lblEmployeeNo");
            LblPassword = ResourcesManager.GetLanguageForControlInForm(loginForm, "lblPassword");
            BtnDBServer = ResourcesManager.GetLanguageForControlInForm(loginForm, "btnDBServerMante");
            BtnLogin = ResourcesManager.GetLanguageForControlInForm(loginForm, "btnLogin");
            BtnClose = ResourcesManager.GetLanguageForControlInForm(loginForm, "btnEnd");
            LblGrpInput = ResourcesManager.GetLanguageForControlInForm(loginForm, "grpInput");
            Message = null;
        }
        #region get/set
        public bool CanLogin
        {
            get
            {
                return canLogin;
            }
            set
            {
                if (canLogin != value)
                {
                    canLogin = value;
                    NotifyOfPropertyChange(() => CanLogin);
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

        public List<string> LanguageList
        {
            get { return languageList; }
        }
        public List<CompanyModel> CompanyList
        {
            get { return companyList; }
            set
            {
                companyList = value;
                NotifyOfPropertyChange(() => CompanyList);
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

        public LanguageModel LblLogin
        {
            get
            {
                return lblLogin;
            }
            set
            {
                if (lblLogin != value)
                {
                    lblLogin = value;
                    NotifyOfPropertyChange(() => LblLogin);
                }
            }
        }

        public LanguageModel LblCompany
        {
            get
            {
                return lblCompany;
            }
            set
            {
                if (lblCompany != value)
                {
                    lblCompany = value;
                    NotifyOfPropertyChange(() => LblCompany);
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
        public LanguageModel BtnDBServer
        {
            get
            {
                return btnDBServer;
            }
            set
            {
                if (btnDBServer != value)
                {
                    btnDBServer = value;
                    NotifyOfPropertyChange(() => BtnDBServer);
                }
            }
        }

        public LanguageModel BtnLogin
        {
            get
            {
                return btnLogin;
            }
            set
            {
                if (btnLogin != value)
                {
                    btnLogin = value;
                    NotifyOfPropertyChange(() => BtnLogin);
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


        #endregion

    }
}
