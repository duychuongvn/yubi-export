using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Caliburn.Micro;
using System.ComponentModel.Composition;
using YUBI_TOOL.Common;
using YUBI_TOOL.Model;
namespace YUBI_TOOL.ViewModel
{
    [Export(typeof(MainMenuViewModel))]
    public class MainMenuViewModel : ViewModelBase
    {
        private const string FORM_ID = "Menu";
        private IEventAggregator eventAggregator;
        private ViewModelBase displayScreen;

        private LanguageModel lblMain;
        private LanguageModel lblExit;
        private LanguageModel lblLogout;
        private LanguageModel lblEmployeeList;
        private LanguageModel lblMessageArea;
        private LanguageModel lblMessageText;

        [ImportingConstructor]
        public MainMenuViewModel(IEventAggregator eventAggregator)
        {
            this.eventAggregator = eventAggregator;
            this.eventAggregator.Subscribe(this);
        }

        public void SetDisplayScreen(ViewModelBase screen)
        {
            displayScreen = screen;
            ActivateItem(displayScreen);
        }
        public void ActivateEmployeeList()
        {
            displayScreen = new EmployeeListViewModel(eventAggregator);
            ActivateItem(displayScreen);
        }
        protected override void OnActivate()
        {
            base.OnActivate();
            ActivateEmployeeList();
        }

        protected override void SetMultiLanguage()
        {
            base.SetMultiLanguage();
            LanguageModel menuForm = ResourcesManager.GetLanguageForForm(FORM_ID);
            var display = ResourcesManager.GetLanguageForControlInForm(menuForm, "Text");
            if (display != null)
            {
                SetDisplayName(display.Text);
            }

            LblMain = ResourcesManager.GetLanguageForControlInForm(menuForm, "Title");
            LblMessageArea = ResourcesManager.GetLanguageForControlInForm(menuForm, "MessageArea");
            LblMessageText = ResourcesManager.GetLanguageForControlInForm(menuForm, "MessageText");
            LblEmployeeList = ResourcesManager.GetLanguageForControlInForm(menuForm, "btnEmployeeList");
            LblLogout = ResourcesManager.GetLanguageForControlInForm(menuForm, "btnLogout");
            LblExit = ResourcesManager.GetLanguageForControlInForm(menuForm, "btnEnd");
          
        }
        public void Logout()
        {
            eventAggregator.Publish(new LogoutEvent());
        }

        #region get/set
        public LanguageModel LblMain
        {
            get
            {
                return lblMain;
            }
            set
            {
                if (lblMain != value)
                {
                    lblMain = value;
                    NotifyOfPropertyChange(() => LblMain);
                }
            }
        }

        public LanguageModel LblExit
        {
            get
            {
                return lblExit;
            }
            set
            {
                if (lblExit != value)
                {
                    lblExit = value;
                    NotifyOfPropertyChange(() => LblExit);
                }
            }
        }

        public LanguageModel LblLogout
        {
            get
            {
                return lblLogout;
            }
            set
            {
                if (lblLogout != value)
                {
                    lblLogout = value;
                    NotifyOfPropertyChange(() => LblLogout);
                }
            }
        }

        public LanguageModel LblEmployeeList
        {
            get
            {
                return lblEmployeeList;
            }
            set
            {
                if (lblEmployeeList != value)
                {
                    lblEmployeeList = value;
                    NotifyOfPropertyChange(() => LblEmployeeList);
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


        #endregion
    }
}
