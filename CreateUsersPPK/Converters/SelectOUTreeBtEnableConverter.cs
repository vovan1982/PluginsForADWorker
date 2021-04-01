using CreateUsersPPK.Model;
using System;
using System.Windows.Data;

namespace CreateUsersPPK.Converters
{
    public class SelectOUTreeBtEnableConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
            {
                return false;
            }
            else
            {
                if (value is string)
                {
                    if (((string)value).StartsWith("DC"))
                        return false;
                    else
                        return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
