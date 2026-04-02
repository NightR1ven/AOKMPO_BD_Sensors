using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace AOKMPO_BD_Sensors.Converters
{
    /// <summary>
    /// Конвертер, сравнивающий значение (value) с параметром (parameter).
    /// Используется в XAML для активации/деактивации элементов в зависимости от выбранного периода.
    /// Например, DatePicker активен только когда SelectedPeriod == "custom".
    /// </summary>
    public class EqualsConverter : IValueConverter
    {
        /// <summary>
        /// Преобразует значение в булево: true, если value == parameter, иначе false.
        /// </summary>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value != null && value.Equals(parameter);
        }

        /// <summary>
        /// Обратное преобразование не поддерживается.
        /// </summary>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
