using System;
using System.Text.RegularExpressions;

namespace YUBI_TOOL.Model
{
    public class MessageModel : ModelBase
    {
        private string messageCode;
        private string message;
        private string background = "Transparent";
        private string foreground = "White";
        private object[] parameters;

        public string GetMessage(string[] parameters = null)
        {
            this.parameters = parameters;
            if (parameters != null)
            {
                return string.Format(message, parameters);
            }
            return message;
        }

        public string MessageCode
        {
            get
            {
               
                return messageCode;
            }
            set
            {
                if (messageCode != value)
                {
                    messageCode = value;
                    NotifyOfPropertyChange(() => MessageCode);
                }
            }
        }

        public string Message
        {
            get
            {
                if (message != null)
                {
                    if (parameters != null)
                    {
                        string value = message;
                        var paramsInMessage = Regex.Match(message, @"({\d+})");
                        if (string.IsNullOrEmpty(paramsInMessage.Groups[0].Value))
                        {
                            foreach (var param in parameters)
                            {
                                value += Environment.NewLine + param;
                            }
                        }
                        return string.Format(value, parameters);
                    }
                   // return message.Replace("\n", Environment.NewLine);
                }
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

        public object[] Parameters
        {
            get { return parameters; }
            set { parameters = value; }
        }
        public string Background
        {
            get
            {
                return background;
            }
            set
            {
                if (background != value)
                {
                    background = value;
                    NotifyOfPropertyChange(() => Background);
                }
            }
        }

        public string Foreground
        {
            get
            {
                return foreground;
            }
            set
            {
                if (foreground != value)
                {
                    foreground = value;
                    NotifyOfPropertyChange(() => Foreground);
                }
            }
        }



        
    }
}
