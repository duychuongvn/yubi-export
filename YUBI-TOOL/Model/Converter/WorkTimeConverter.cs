using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace YUBI_TOOL.Model.Converter
{
    class WorkTimeConverter : IValueConverter
    {
        private const string TIME_FORMAT = "00:00";
        // Summary:
        //     Converts a value.
        //
        // Parameters:
        //   value:
        //     The value produced by the binding source.
        //
        //   targetType:
        //     The type of the binding target property.
        //
        //   parameter:
        //     The converter parameter to use.
        //
        //   culture:
        //     The culture to use in the converter.
        //
        // Returns:
        //     A converted value. If the method returns null, the valid null value is used.
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            decimal data;
            if (value == null || !decimal.TryParse(value.ToString(), out data))
            {
                return null;
            }

            data = Common.CommonUtil.MinuteToHr(data);
            string param = TIME_FORMAT;

            if (parameter != null)
            {
                param = parameter.ToString();
            }
            return data.ToString(param);
        }
        //
        // Summary:
        //     Converts a value.
        //
        // Parameters:
        //   value:
        //     The value that is produced by the binding target.
        //
        //   targetType:
        //     The type to convert to.
        //
        //   parameter:
        //     The converter parameter to use.
        //
        //   culture:
        //     The culture to use in the converter.
        //
        // Returns:
        //     A converted value. If the method returns null, the valid null value is used.
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            decimal data;
            if (value == null)
            {
                return null;
            }
            Match validTime = Regex.Match(value.ToString(), "^[0-9]{1,2}[:][0-5]{1}[0-9]{1}$|| (^[0-9]{1,2}[0-5]{1}[0-9]{1}$)");
            if (!validTime.Success)
            {
                return null;
            }
            string valuestr = Regex.Replace(value.ToString(), "[a-zA-Z./:-]", "");
            if (decimal.TryParse(valuestr, out data))
            {
                if (data > 4800)
                {
                    return null;
                }
                return Common.CommonUtil.HourToMinute(data);
            }
            return null;
        }

    }
}
