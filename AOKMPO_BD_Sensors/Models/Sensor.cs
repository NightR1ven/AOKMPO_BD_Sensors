using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace AOKMPO_BD_Sensors
{
    /// <summary>
    /// Класс, представляющий датчиков в системе учета
    /// Реализует INotifyPropertyChanged для уведомления об изменениях свойств
    /// </summary>
    public class Sensor : INotifyPropertyChanged
    {


            // Приватные поля для хранения значений свойств
            private string _name;
            private string _typeSensor;
            private string _serialNumber;
            private DateTime _placementDate;
            private DateTime _expiryDate;
            private string _location;
            private string _placeOfUse;
            private string _measurementLimits;
            private string _classForSure;

            /// <summary>
            /// Название датчика
            /// </summary>
            public string Name
            {
                get => _name;
                set { _name = value; OnPropertyChanged(nameof(Name)); }
            }

            /// <summary>
            /// Заводской номер датчика
            /// </summary>
            public string SerialNumber
            {
                get => _serialNumber;
                set { _serialNumber = value; OnPropertyChanged(nameof(SerialNumber)); }
            }

            /// <summary>
            /// Тип датчика
            /// </summary>
            public string TypeSensor
            {
                get => _typeSensor;
                set { _typeSensor = value; OnPropertyChanged(nameof(TypeSensor)); }
            }

        /// <summary>
        /// Дата обслуживание датчика
        /// </summary>
        public DateTime PlacementDate
            {
                get => _placementDate;
                set { _placementDate = value; OnPropertyChanged(nameof(PlacementDate)); }
            }

            /// <summary>
            /// Срок до проверки датчика
            /// При изменении также обновляет цвет и статус
            /// </summary>
            public DateTime ExpiryDate
            {
                get => _expiryDate;
                set
                {
                    _expiryDate = value;
                    OnPropertyChanged(nameof(ExpiryDate));
                    OnPropertyChanged(nameof(ExpiryColor));
                    OnPropertyChanged(nameof(ExpiryStatus));
                }
            }

        /// <summary>
        /// Класс точности
        /// </summary>
        public string ClassForSure
        {
            get => _classForSure;
            set { _classForSure = value; OnPropertyChanged(nameof(ClassForSure)); }
        }

        /// <summary>
        /// Место расположения датчика
        /// </summary>
        public string Location
            {
                get => _location;
                set { _location = value; OnPropertyChanged(nameof(Location)); }
            }

            /// <summary>
            /// Пределы измерения
            /// </summary>
            public string MeasurementLimits
            {
                get => _measurementLimits;
                set { _measurementLimits = value; OnPropertyChanged(nameof(MeasurementLimits)); }
            }

        /// <summary>
        /// Место эксплуатация
        /// </summary>
        public string PlaceOfUse
            {
                get => _placeOfUse;
                set { _placeOfUse = value; OnPropertyChanged(nameof(PlaceOfUse)); }
            }


            /// <summary>
            /// Цвет для отображения в зависимости от срока до проверки
            /// </summary>
            public Brush ExpiryColor
            {
                get
                {

                    TimeSpan remaining = ExpiryDate - DateTime.Today;

                    if (remaining.TotalDays <= 14)
                        return Brushes.Red; // Менее 2 недель
                    else if (remaining.TotalDays <= 30)
                        return Brushes.LightYellow; // Менее года - желтый
                    else
                        return Brushes.LightGreen; // Более года - зеленый
                }
            }

            /// <summary>
            /// Текстовый статус срока до проверки
            /// </summary>
            public string ExpiryStatus
            {
                get
                {
                    TimeSpan remaining = ExpiryDate - DateTime.Today;

                    if (remaining.TotalDays <= 0)
                        return "Годен"; // Убираем статус "Просрочено"
                    else if (remaining.TotalDays <= 30)
                        return $"Осталось {remaining.Days} дней";
                    else if (remaining.TotalDays <= 365)
                        return $"Осталось {remaining.Days / 30} месяцев";
                    else
                        return $"Осталось {remaining.Days / 365} лет";
                }
            }

        /// <summary>
        /// Событие, возникающее при изменении свойств
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

            /// <summary>
            /// Метод для вызова события PropertyChanged
            /// </summary>
            /// <param name="propertyName">Имя изменившегося свойства</param>
            protected virtual void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
    }
}
