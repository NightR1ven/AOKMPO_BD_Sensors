﻿
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using AOKMPO_BD_Sensors;

namespace AOKMPO_BD_Sensors
{
     public class ApproachingSensorsConverter : IValueConverter
    {
        /// <summary>
        /// Конвертация - подсчет подходящим по срокам датчики
        /// </summary>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int count && App.Current.MainWindow?.DataContext is MainViewModel vm)
                return vm.Sensors.Count(t =>
                    t.ExpiryDate >= DateTime.Today &&
                    t.ExpiryDate <= DateTime.Today.AddDays(30));

            return 0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
