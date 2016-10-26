using System;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace CreateUsers.Converters
{
    public class PathToImageSourceConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            BitmapImage result = new BitmapImage();

            if (value is String)
            {
                result = new BitmapImage(new Uri((string)value, UriKind.RelativeOrAbsolute));
            }

            return result;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
