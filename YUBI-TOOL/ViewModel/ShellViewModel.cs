using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel.Composition;
using Caliburn.Micro;
using YUBI_TOOL.Common;

namespace YUBI_TOOL.ViewModel
{
    [Export(typeof(IShell))]
    public class ShellViewModel : ViewModelBase, IShell, IHandle<LogoutEvent>
    {
        private LoginViewModel loginViewModel;
        private ViewModelBase mainMenuViewModel;
        private IEventAggregator eventAggregator;
        [ImportingConstructor()]
        public ShellViewModel(IEventAggregator eventAggregator)
        {
            this.DisplayName = "YUBI-TOOL";
            this.eventAggregator = eventAggregator;
            this.loginViewModel = IoC.Get<LoginViewModel>();
            this.mainMenuViewModel = new MainMenuViewModel(eventAggregator);
            this.eventAggregator.Subscribe(this);
            ActivateItem(loginViewModel);
        }

        void ShellViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "ActiveItem" || e.PropertyName == "Items")
            {
                if (Items.Count == 0)
                {
                    ActivateItem(loginViewModel);
                }
            }
        }

        public void ActiveMainWindow()
        {
            ActivateItem(mainMenuViewModel);
        }
        public void ActiveDBServerMainte()
        {
            DBServerViewModel dbServer = new DBServerViewModel();
            ActivateItem(dbServer);
        }

        protected override void OnDeactivate(bool close)
        {
            base.OnDeactivate(close);
            if (close)
            {
                CommonUtil.ClearTemplate();
            }
        }
        void IHandle<LogoutEvent>.Handle(LogoutEvent message)
        {
            ActivateItem(loginViewModel);
        }
    }
}
