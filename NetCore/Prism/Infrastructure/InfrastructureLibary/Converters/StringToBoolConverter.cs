using System;
using System.Globalization;
using System.Windows.Data;

namespace InfrastructureLibary.Converters
{
    public class StringToBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string st = value as string;
            return !string.IsNullOrEmpty(st);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
