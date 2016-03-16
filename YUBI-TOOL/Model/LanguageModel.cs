using System;
using System.Collections.Generic;

namespace YUBI_TOOL.Model
{
    public class LanguageModel : ModelBase
    {
        private string formId;
        private string controlId;
        private string text;
        private string fontFamily;
        private decimal fontSize;
        List<LanguageModel> childControls;
        public LanguageModel()
        {
            ChildControls = new List<LanguageModel>();
        }
             
       
        public string FormId
        {
            get
            {
                return formId;
            }
            set
            {
                if (formId != value)
                {
                    formId = value;
                    NotifyOfPropertyChange(() => FormId);
                }
            }
        }
        public string ControlId
        {
            get
            {
                return controlId;
            }
            set
            {
                if (controlId != value)
                {
                    controlId = value;
                    NotifyOfPropertyChange(() => ControlId);
                }
            }
        }

        public string Text
        {
            get
            {
                if (!string.IsNullOrEmpty(text))
                {
                    return text.Replace("\\n", Environment.NewLine);
                }
                return text;
            }
            set
            {
                if (text != value)
                {
                    text = value;
                    NotifyOfPropertyChange(() => Text);
                }
            }
        }

        public string FontFamily
        {
            get
            {
                return fontFamily;
            }
            set
            {
                if (fontFamily != value)
                {
                    fontFamily = value;
                    NotifyOfPropertyChange(() => FontFamily);
                }
            }
        }

        public decimal FontSize
        {
            get
            {
                return fontSize;
            }
            set
            {
                if (fontSize != value)
                {
                    fontSize = value;
                    NotifyOfPropertyChange(() => FontSize);
                }
            }
        }

        public List<LanguageModel> ChildControls
        {
            get { return childControls; }
            set { childControls = value; }
        }
    }
}
