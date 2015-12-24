using System;
using System.Collections.Generic;
using System.Windows.Data;

namespace Converters
  {
  public class ValueConverterGroup : List<IValueConverter>, IValueConverter
    {
    #region IValueConverter Members

    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
      {
      //return this.Aggregate(value, (current, converter) => converter.Convert(current, targetType, parameter, culture));
      object returnValue = value;
      foreach (IValueConverter converter in this)
        returnValue = converter.Convert(returnValue, targetType, parameter, culture);
      return returnValue;
      }

    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
      {
      throw new NotImplementedException();
      }

    #endregion
    }

  [ValueConversion(typeof(bool), typeof(bool))]
  public class BooleanInverterConverter : IValueConverter
    {
    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
      {
      if (value != null && value is bool)
        {
        return !((bool)value);
        }

      return true;
      }

    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
      {
      return Convert(value, targetType, parameter, culture);
      }
    }
  }
