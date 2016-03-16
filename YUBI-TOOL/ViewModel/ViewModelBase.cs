using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Caliburn.Micro;
using YUBI_TOOL.Model;
using System.ComponentModel.DataAnnotations;
using System.Linq.Expressions;
using System.Data.SqlClient;
namespace YUBI_TOOL.ViewModel
{
    public class ViewModelBase : Caliburn.Micro.Conductor<IScreen>.Collection.OneActive
    {
        private System.Collections.ObjectModel.ObservableCollection<MessageModel> messageList;
        private MessageModel message;
        protected bool IsActivated { get; set; }
        public ViewModelBase()
        {
            message = new MessageModel();
        }
        public void Close()
        {
            ShellViewModel shellViewModel = (ShellViewModel)IoC.Get<IShell>();
            shellViewModel.TryClose();
        }

        protected void SetDisplayName(string displayName)
        {
            ShellViewModel shellViewModel = (ShellViewModel)IoC.Get<IShell>();
            shellViewModel.DisplayName = displayName;
        }
        protected void AddMessage(string messageCode)
        {
            if (MessageList == null)
            {
                MessageList = new System.Collections.ObjectModel.ObservableCollection<MessageModel>();
                MessageModel message = new MessageModel();
                messageList.Add(message);
            }
        }

        protected void ActiveScreen(ViewModelBase currentScreen, ViewModelBase newScreen)
        {

            MainMenuViewModel mainMenu = FindMainMenu(currentScreen) as MainMenuViewModel;
            if (mainMenu != null)
            {
                mainMenu.SetDisplayScreen(newScreen);
            }
        }

        private ViewModelBase FindMainMenu(ViewModelBase childScreen)
        {
            if (childScreen is MainMenuViewModel)
            {
                return childScreen;
            }
            else
            {
                ViewModelBase parrent = childScreen.Parent as ViewModelBase;
                if (parrent is MainMenuViewModel)
                {
                    return parrent;
                }
                return FindMainMenu(parrent);
            }
        }

        protected override void OnActivate()
        {
            base.OnActivate();
            SetMultiLanguage();
        }
        protected virtual void SetMultiLanguage()
        {
            Message = new MessageModel();
        }

        protected List<string> GetMessage(ModelBase model)
        {
            List<ValidationResult> validationResults = new List<ValidationResult>();
            List<string> messageCodes = new List<string>();

            if (!model.TryValidateNestedObject(validationResults))
            {
                foreach (var error in validationResults)
                {
                    messageCodes.Add(error.ErrorMessage);
                }
            }
            return messageCodes;
        }

        /// <summary>
        /// Get Name of property is Object using lambda expression
        /// </summary>
        protected string GetPropertyName(Expression<Func<object>> propertyExpression)
        {
            MemberExpression member = propertyExpression.Body as MemberExpression;
            if (member != null)
            {
                return member.Member.Name;
            }
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get Name of property with type in class T using lambda expression 
        /// </summary>
        protected string GetPropertyName<TSource, TType>(Expression<Func<TSource, TType>> expr)
        {
            var node = expr.Body as MemberExpression;
            if (!object.ReferenceEquals(null, node))
            {
                return node.Member.Name;
            }

            throw new NotImplementedException();

        }

        protected bool TestConnection(string connectString)
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
        #region get/set
        public System.Collections.ObjectModel.ObservableCollection<MessageModel> MessageList
        {
            get
            {
                return messageList;
            }
            set
            {
                if (messageList != value)
                {
                    messageList = value;
                    NotifyOfPropertyChange(() => MessageList);
                }
            }
        }
        public MessageModel Message
        {
            get
            {
                return message;
            }
            set
            {
                if (message != value)
                {
                    message = value;
                    NotifyOfPropertyChange(() => Message);
                }
            }
        }
        #endregion
    }
}
