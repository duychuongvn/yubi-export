using System.Collections.Generic;
using YUBI_TOOL.Model;

namespace YUBI_TOOL.Common
{
    public class LanguageMessageUtil
    {
        private static Dictionary<string, Dictionary<string, MessageModel>> languageMap;

        static LanguageMessageUtil()
        {
            languageMap = new Dictionary<string, Dictionary<string, MessageModel>>();
        }
        public static Dictionary<string, Dictionary<string, MessageModel>> GetInstance()
        {
            if (languageMap == null)
            {
                languageMap = new Dictionary<string, Dictionary<string, MessageModel>>();
            }
            return languageMap;
        }

        public static MessageModel GetMessageText(string messageCode)
        {
            if (GetInstance().ContainsKey(Properties.Settings.Default.SelectedLanguage))
            {
                Dictionary<string, MessageModel> messageByLanguage = GetInstance()[Properties.Settings.Default.SelectedLanguage];
                if (messageByLanguage.ContainsKey(messageCode))
                {
                    return messageByLanguage[messageCode];
                }
            }
            return null;
        }

        public static void PutMessage(MessageModel message)
        {
            Dictionary<string, MessageModel> messageByLanguage;
            if (GetInstance().ContainsKey(Properties.Settings.Default.SelectedLanguage))
            {
                messageByLanguage = GetInstance()[Properties.Settings.Default.SelectedLanguage];

            }
            else
            {
                messageByLanguage = new Dictionary<string, MessageModel>();
                GetInstance().Add(Properties.Settings.Default.SelectedLanguage, messageByLanguage);
            }

            if (messageByLanguage.ContainsKey(message.MessageCode))
            {
                messageByLanguage.Remove(message.MessageCode);
            }
            messageByLanguage.Add(message.MessageCode, message);
        }
    }
}
