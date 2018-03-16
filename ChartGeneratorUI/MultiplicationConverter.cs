using System;
using System.Globalization;
using System.Windows.Data;

namespace ChartGeneratorUI
{
    class MultiplicationConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            Double.TryParse(value?.ToString(), out var x);
            Double.TryParse(parameter?.ToString(), out var y);
            return x * y;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
