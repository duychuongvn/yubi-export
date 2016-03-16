using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
using YUBI_TOOL.Model;

namespace YUBI_TOOL.Common
{
    public enum ResourceEnum
    {
        Caption = 0,
        Message = 1
    }
    public class ResourcesManager
    {
        private const string LANGUAGE_DIR = @"\Language\{0}.xml";
        private const string NODE_TEXT = "Text";
        private const string NODE_FONT = "Font";
        private const string NODE_TITLE = "Title";
        private const string DEFAULT_FONT = "Times New Roman";
        public const string KEY_COLOR_DEFAULT = "DEFAULT";
        public const string KEY_COLOR_MESSAGE_INFO = ",1,64";
        public const string KEY_COLOR_MESSAGE_ERROR = ",0,16";

        public static string LIST_DATE_FORMAT
        {
            get
            {
                CreateInstance();
                var constForm = GetLanguageForForm("const");
                if (constForm != null)
                {
                    var list_date_format = GetLanguageForControlInForm(constForm, "list_date_format");
                    if (list_date_format != null)
                    {
                        return list_date_format.Text;
                    }
                }
                return "YM";
            }
        }
        public static string[] WEEK_DAY_SHORT
        {
            get
            {
                CreateInstance();
                var constForm = GetLanguageForForm("const");
                if (constForm != null)
                {
                    var week_short = GetLanguageForControlInForm(constForm, "weekday_short");
                    if (week_short != null)
                    {
                        return week_short.Text.Split(',');
                    }
                }
                return new string[] { };
            }
        }
        public static string[] MONTH_SHORT
        {
            get
            {
                CreateInstance();
                var constForm = GetLanguageForForm("const");
                if (constForm != null)
                {
                    var week_short = GetLanguageForControlInForm(constForm, "month_short");
                    if (week_short != null)
                    {
                        return week_short.Text.Split(',');
                    }
                }
                return new string[] { };
            }
        }
        public static string[] WORK_TYPE
        {
            get
            {
                CreateInstance();
                var constForm = GetLanguageForForm("const");
                if (constForm != null)
                {
                    var week_short = GetLanguageForControlInForm(constForm, "work_type");
                    if (week_short != null)
                    {
                        return week_short.Text.Split(',');
                    }
                }
                return new string[] { };
            }
        }
        public static string[] MONTH_LONG
        {
            get
            {
                CreateInstance();
                var constForm = GetLanguageForForm("const");
                if (constForm != null)
                {
                    var week_short = GetLanguageForControlInForm(constForm, "month_long");
                    if (week_short != null)
                    {
                        return week_short.Text.Split(',');
                    }
                }
                return new string[] { };
            }
        }

        

        public string a { get; set; }
        private static XmlDocument docRoot = null;
        private static string selectedLanguage = Properties.Settings.Default.SelectedLanguage;

        private static Dictionary<string, int> MAP_FONT_SIZE = new Dictionary<string, int>
       {
           {"Japanese", 14},
           {"English", 13},
           {"Vietnamese", 13}
       };

        private static Dictionary<string, string[]> MAP_TEXT_COLOR = new Dictionary<string, string[]>
       {
           {KEY_COLOR_MESSAGE_ERROR, new string[] {"Yellow","Red"}},
            {KEY_COLOR_MESSAGE_INFO, new string[]{"Violet", "White"}},
            {KEY_COLOR_DEFAULT, new string[]{"Transparent", "Black"}},
       };

        public static string GetBackground(string color_key)
        {
            if (MAP_TEXT_COLOR.ContainsKey(color_key))
            {
                return MAP_TEXT_COLOR[color_key][0];
            }
            return MAP_TEXT_COLOR[KEY_COLOR_DEFAULT][0];
        }
        public static string GetForeground(string color_key)
        {
            if (MAP_TEXT_COLOR.ContainsKey(color_key))
            {
                return MAP_TEXT_COLOR[color_key][1];
            }
            return MAP_TEXT_COLOR[KEY_COLOR_DEFAULT][1];
        }


        private static void CreateInstance()
        {
            if (docRoot == null || selectedLanguage != Properties.Settings.Default.SelectedLanguage)
            {
                docRoot = new XmlDocument();
                selectedLanguage = Properties.Settings.Default.SelectedLanguage;
                docRoot.Load(Directory.GetCurrentDirectory() + string.Format(LANGUAGE_DIR, selectedLanguage));
            }
        }

        public static MessageModel GetMessage(string messageCode, object[] messageParams = null)
        {
            CreateInstance();
            MessageModel message = new MessageModel();
            LanguageModel language = GetLanguageForForm("message");
            message.MessageCode = messageCode;
            message.Parameters = messageParams;
            if (language != null)
            {
                foreach (var child in language.ChildControls)
                {
                    if (child.ControlId == messageCode)
                    {
                        string messageText = child.Text;
                        if (!string.IsNullOrEmpty(messageText))
                        {
                            var colorRegrex = Regex.Match(messageText, @"([,]\d+,\d+)$|([,][,]\d+)$");
                            string color = colorRegrex.Groups[0].Value;
                            if (!string.IsNullOrEmpty(color))
                            {
                                message.Message = messageText.Replace(color, "");
                                if (MAP_TEXT_COLOR.ContainsKey(color))
                                {
                                    message.Background = MAP_TEXT_COLOR[color][0];
                                    message.Foreground = MAP_TEXT_COLOR[color][1];
                                }
                                else
                                {
                                    message.Background = MAP_TEXT_COLOR[KEY_COLOR_MESSAGE_INFO][0];
                                    message.Foreground = MAP_TEXT_COLOR[KEY_COLOR_MESSAGE_INFO][1];
                                }
                            }
                            else
                            {
                                message.Message = messageText;
                                message.Background = MAP_TEXT_COLOR[KEY_COLOR_DEFAULT][0];
                                message.Foreground = MAP_TEXT_COLOR[KEY_COLOR_DEFAULT][1];
                            }
                        }

                        break;
                    }
                }
            }
            return message;
        }
        public static LanguageModel GetLanguageForControlInForm(LanguageModel parrent, string controlId)
        {

            if (parrent != null)
            {
                var child = parrent.ChildControls.Find(x => x.ControlId == controlId);
                if (child == null)
                {
                    child = FindControl(parrent, controlId);

                }
                return child;
            }
            return null;
        }


        public static LanguageModel GetLanguageForForm(string formId, string controlId)
        {
            LanguageModel language = GetLanguageForForm(formId);
            if (language != null)
            {
                var child = language.ChildControls.Find(x => x.ControlId == controlId);
                if (child == null)
                {
                    child = FindControl(language, controlId);

                }
                return child;
            }
            return language;
        }

        private static LanguageModel FindControl(LanguageModel parrent, string controlId)
        {
            LanguageModel child = parrent.ChildControls.Find(x => x.ControlId == controlId);
            if (child == null)
            {
                foreach (var language in parrent.ChildControls)
                {
                    child = FindControl(language, controlId);
                    if (child != null)
                    {
                        return child;
                    }
                }
            }
            return child;
        }

        public static LanguageModel GetLanguageForForm(string formId)
        {
            LanguageModel language = null;
            CreateInstance();
            XmlNodeList forms = docRoot.SelectNodes("//" + formId);
            if (forms.Count == 1)
            {
                language = new LanguageModel();
                language.FontSize = MAP_FONT_SIZE[selectedLanguage];
                language.FontFamily = DEFAULT_FONT;
                XmlNode form = forms[0];
                foreach (XmlNode child in form.ChildNodes)
                {

                    LanguageModel childControl = new LanguageModel();
                    childControl.FontSize = MAP_FONT_SIZE[selectedLanguage];
                    childControl.FontFamily = DEFAULT_FONT;
                    language.ChildControls.Add(childControl);
                    ReadNode(formId, child, childControl);
                }
            }
            return language;
        }

        private static void ReadNode(string formId, XmlNode node, LanguageModel parrent)
        {
            if (node.HasChildNodes)
            {
                parrent.FormId = formId;
                parrent.ControlId = node.Name;
                XmlNodeList childrenNode = node.ChildNodes;
                foreach (XmlNode child in childrenNode)
                {
                    if (child.Name == NODE_TEXT || child.Name == "#" + NODE_TEXT.ToLower())
                    {
                        parrent.Text = child.InnerText;
                    }
                    else if (child.Name == NODE_FONT)
                    {
                        string[] font = child.InnerText.Split(',');
                        if (font.Length == 2)
                        {
                            parrent.FontSize = decimal.Parse(font[1]);
                        }
                        parrent.FontFamily = font[0];
                    }
                    else
                    {
                        LanguageModel childControl = new LanguageModel();
                        childControl.FontSize = MAP_FONT_SIZE[selectedLanguage];
                        childControl.FontFamily = DEFAULT_FONT;
                        parrent.ChildControls.Add(childControl);
                        ReadNode(formId, child, childControl);
                    }

                }
            }
            else
            {


                parrent.FormId = formId;
                if (node.Name == NODE_TEXT || node.Name == "#" + NODE_TEXT)
                {
                    parrent.Text = node.InnerText;
                }
                else if (node.Name == NODE_FONT)
                {
                    string[] font = node.InnerText.Split(',');
                    if (font.Length == 2)
                    {
                        parrent.FontSize = int.Parse(font[1]);
                    }
                    parrent.FontFamily = font[0];
                }



            }
        }
    }
}
