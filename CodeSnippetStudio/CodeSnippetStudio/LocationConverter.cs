using Microsoft.CodeAnalysis;
using System;
using System.Globalization;
using System.Windows.Data;

namespace CodeSnippetStudio
{
    public class LocationConverter : IValueConverter
    {

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            Location loc = (Location)value;
            return (loc.GetLineSpan().StartLinePosition.Line + 1).ToString();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
