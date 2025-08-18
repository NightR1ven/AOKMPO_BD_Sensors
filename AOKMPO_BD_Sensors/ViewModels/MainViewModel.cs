using AOKMPO_BD_Sensors.Service;
using AOKMPO_BD_Sensors.Service;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Xml.Serialization;


    namespace AOKMPO_BD_Sensors
    {
        /// <summary>
        /// Основная ViewModel приложения
        /// Содержит логику работы с данными и команды для UI
        /// </summary>
        public class MainViewModel : INotifyPropertyChanged
        {
            // Путь к файлу хранения данных
            private const string DataFilePath = "sensors.xml";

            // Выбранный датчик в списке
            private Sensor _selectedSensor;

            // Текст для поиска
            private string _searchText;

            //Стартовые данные поиска по дате
            private DateTime _startDate = DateTime.Today.AddMonths(-1);
            private DateTime _endDate = DateTime.Today.AddMonths(1);

            //Индикатор загрузки
            private bool _isLoading;
            private double _progressValue;
            private string _progressMessage;

            // Объект для потокобезопасной работы с коллекцией
            private readonly object _collectionLock = new object();

            /// <summary>
            /// Коллекция датчиков (привязана к ListView)
            /// </summary>
            public ObservableCollection<Sensor> Sensors { get; } = new ObservableCollection<Sensor>();

        // Заменена обычная коллекция на ICollectionView для эффективной фильтрации
        // ИЗМЕНЕНИЕ: Добавлена ленивая инициализация и фильтр по умолчанию
        private ICollectionView _sensorsView;
        public ICollectionView SensorsView
        {
            get
            {
                if (_sensorsView == null)
                {
                    _sensorsView = CollectionViewSource.GetDefaultView(Sensors);
                    _sensorsView.Filter = SensorFilter;
                }
                return _sensorsView;
            }
        }


        // Cвойства для отображения прогресса

        public bool IsLoading
            {
            get => _isLoading;
            set { _isLoading = value; OnPropertyChanged(nameof(IsLoading)); } // Уведомление UI об изменении состояния
            }

            public double ProgressValue
            {
                get => _progressValue;
                set { _progressValue = value; OnPropertyChanged(nameof(ProgressValue)); }
            }

            public string ProgressMessage
            {
            get => _progressMessage;
            set { _progressMessage = value; OnPropertyChanged(nameof(ProgressMessage)); }
            }

            /// <summary>
            /// Выбранный датчик
            /// </summary>
            public Sensor SelectedSensor
            {
                get => _selectedSensor;
                set { _selectedSensor = value; OnPropertyChanged(nameof(SelectedSensor)); }
            }

            /// <summary>
            /// Стартовая дата
            /// </summary>
            public DateTime StartDate
            {
                get => _startDate;
                set { _startDate = value; OnPropertyChanged(nameof(StartDate)); }
            }

            /// <summary>
            /// Последняя дата
            /// </summary>
            public DateTime EndDate
            {
                get => _endDate;
                set { _endDate = value; OnPropertyChanged(nameof(EndDate)); }
            }

        /// <summary>
        /// Текст для поиска датчиков
        /// При изменении автоматически фильтрует список
        /// </summary>
        public string SearchText
        {
            get => _searchText;
            set
            {
                _searchText = value; OnPropertyChanged(nameof(SearchText)); ApplyFilters();
                // ИЗМЕНЕНИЕ: Заменен FilterSensors на ApplyFilters
            }
        }

            // Команды для привязки к кнопкам в UI
            public ICommand AddCommand { get; }       // Добавить датчик
            public ICommand EditCommand { get; }     // Редактировать датчиков
            public ICommand DeleteCommand { get; }    // Удалить датчиков
            public ICommand ShowAllCommand { get; }   // Показать все датчики
            public ICommand ShowExpiredCommand { get; } // Показать просроченные
            public ICommand ShowExpiringCommand { get; } // Показать истекающие
            public ICommand ExportCommand { get; }    // Экспорт в CSV
            public ICommand ExportToExcelCommand { get; } // Экспорт в Excel
            public ICommand ImportFromExcelCommand { get; } // Импорт в Excel
             public ICommand ExportToExcelReportCommand { get; } // Экспорт отчета в Excel
            public ICommand FilterByDateCommand { get; } // Фильтр по дате

            /// <summary>
            /// Конструктор ViewModel
            /// Инициализирует команды и загружает данные
            /// </summary>
            public MainViewModel()
            {
                // Инициализация команд
                AddCommand = new RelayCommand(AddSensor);
                EditCommand = new RelayCommand(EditSensor, CanEditOrDelete);
                DeleteCommand = new RelayCommand(DeleteSensor, CanEditOrDelete);
                ShowAllCommand = new RelayCommand(ShowAllSensors);
                ShowExpiringCommand = new RelayCommand((ShowApproachingSensors));
                ShowExpiredCommand = new RelayCommand(ShowExpiredSensors);
                ExportCommand = new RelayCommand(ExportSensors);
                ExportToExcelCommand = new RelayCommand(_ => ExportToExcel());
                ExportToExcelReportCommand = new RelayCommand(_ => ExportReportToExcel());
                ImportFromExcelCommand = new RelayCommand(_ => ImportFromExcel());
                FilterByDateCommand = new RelayCommand(FilterByDateRange);

                _ = LoadDataAndCheckExpiredAsync();

            this.PropertyChanged += (s, e) =>
            {
                if ((e.PropertyName == nameof(StartDate) ||
                    (e.PropertyName == nameof(EndDate))))
                {
                    // При ручном изменении дат переключаем на custom период
                    if (_selectedPeriod != "custom")
                    {
                        _selectedPeriod = "custom";
                        OnPropertyChanged(nameof(SelectedPeriod));
                    }
                    ApplyFilters();
                }
            };
        }

        /// <summary>
        /// Универсальный фильтр для датчиков, объединяющий все условия
        /// С учетом возможных null-значений в полях датчика
        /// </summary>
        private bool SensorFilter(object item)
        {
            if (item is not Sensor sensor) return false;

            // Фильтр по дате
            bool dateMatch = FilterByDate(sensor);

            // Если нет текста для поиска - возвращаем только результат фильтрации по дате
            if (string.IsNullOrEmpty(SearchText))
                return dateMatch;

            // Специальный фильтр для просроченных датчиков
            if (SearchText == "status:expired")
                return dateMatch && sensor.ExpiryDate < DateTime.Today;

            // Специальный фильтр для истекающих датчиков
            if (SearchText == "status:approaching")
                return dateMatch && sensor.ExpiryDate <= DateTime.Today.AddDays(30);

            // Стандартный текстовый поиск
            var searchTerm = SearchText.ToLower();

            bool textMatch =
                (sensor.Name?.ToLower().Contains(searchTerm) ?? false) ||
                (sensor.SerialNumber?.ToLower().Contains(searchTerm) ?? false) ||
                (sensor.TypeSensor?.ToLower().Contains(searchTerm) ?? false) ||
                (sensor.MeasurementLimits?.ToLower().Contains(searchTerm) ?? false) ||
                (sensor.ClassForSure?.ToLower().Contains(searchTerm) ?? false) ||
                (sensor.Location?.ToLower().Contains(searchTerm) ?? false) ||
                (sensor.PlaceOfUse?.ToLower().Contains(searchTerm) ?? false) ||
                (sensor.ExpiryStatus?.ToLower().Contains(searchTerm) ?? false) ||
                (sensor.SerialNumber != null && Regex.IsMatch(sensor.SerialNumber, searchTerm)) ||
                sensor.PlacementDate.ToString("dd.MM.yyyy").Contains(searchTerm) ||
                sensor.ExpiryDate.ToString("dd.MM.yyyy").Contains(searchTerm);

            return dateMatch && textMatch;
        }

        /// <summary>
        /// Фильтрация датчика по диапазону дат с проверкой на значения по умолчанию
        /// </summary>
        private bool FilterByDate(Sensor sensor)
        {
            // Для периода "Все даты" не фильтруем
            if (SelectedPeriod == "all")
            {
                return true;
            }
                

            return sensor.ExpiryDate >= StartDate && sensor.ExpiryDate <= EndDate;
        }

        /// <summary>
        /// Применить текущие фильтры к коллекции
        /// НОВЫЙ МЕТОД: Заменяет прямое изменение коллекции
        /// </summary>
        
        /// <summary>
/// Флаг, указывающий применять ли фильтр по дате
/// </summary>
private bool _useDateFilter;
public bool UseDateFilter
{
    get => _useDateFilter;
    set
    {
        _useDateFilter = value;
        OnPropertyChanged(nameof(UseDateFilter));
        ApplyFilters();
    }
}

/// <summary>
/// Доступные предустановленные периоды для фильтрации
/// </summary>
    public Dictionary<string, DatePeriod> DatePeriods { get; } = new()
    {
        ["all"] = new DatePeriod
        {
            Name = "Все даты",
            Action = (vm) => vm.SetFullDateRange()
        },
        ["month"] = new DatePeriod
        {
            Name = "Текущий месяц",
            Action = (vm) => vm.SetCurrentMonthRange()
        },
        ["year"] = new DatePeriod
        {
            Name = "Текущий год",
            Action = (vm) => vm.SetCurrentYearRange()
        },
        ["custom"] = new DatePeriod
        {
            Name = "Произвольный диапазон",
            Action = (vm) => { }
        }
    };

        private string _selectedPeriod = "all";
        public string SelectedPeriod
        {
            get => _selectedPeriod;
            set
            {
                _selectedPeriod = value;
                OnPropertyChanged(nameof(SelectedPeriod));

                // Очищаем поисковую строку при изменении периода
                SearchText = string.Empty;

                if (DatePeriods.TryGetValue(value, out var period))
                {
                    period.Action(this);
                    ApplyFilters();
                }
            }
        }
        /// <summary>
        /// Класс для описания периода фильтрации
        /// </summary>
        public class DatePeriod
        {
            public string Name { get; set; }
            public Action<MainViewModel> Action { get; set; }
        }
        /// <summary>
        /// Обновляет диапазон дат, чтобы включить все имеющиеся датчики
        /// </summary>
        public void SetFullDateRange()
        {
            if (Sensors.Count == 0)
            {
                StartDate = DateTime.Today.AddMonths(-1);
                EndDate = DateTime.Today.AddMonths(1);
                return;
            }

            StartDate = Sensors.Min(s => s.ExpiryDate);
            EndDate = Sensors.Max(s => s.ExpiryDate);

            // Добавляем небольшой запас по краям
            StartDate = StartDate.AddDays(-7);
            EndDate = EndDate.AddDays(7);

            OnPropertyChanged(nameof(StartDate));
            OnPropertyChanged(nameof(EndDate));
        }

        // <summary>
        /// Устанавливает диапазон текущего месяца
        /// </summary>
        public void SetCurrentMonthRange()
        {
            var today = DateTime.Today;
            StartDate = new DateTime(today.Year, today.Month, 1);
            EndDate = StartDate.AddMonths(1).AddDays(-1);

            OnPropertyChanged(nameof(StartDate));
            OnPropertyChanged(nameof(EndDate));
        }

        /// <summary>
        /// Устанавливает диапазон текущего года
        /// </summary>
        public void SetCurrentYearRange()
        {
            var today = DateTime.Today;
            StartDate = new DateTime(today.Year, 1, 1);
            EndDate = new DateTime(today.Year, 12, 31);

            OnPropertyChanged(nameof(StartDate));
            OnPropertyChanged(nameof(EndDate));
        }

        private void ApplyFilters()
        {
            if (_sensorsView != null)
            {
                _sensorsView.Refresh();
            }
        }
        private async Task LoadDataAndCheckExpiredAsync()
            {
            await LoadDataAsync();
            CheckExpiredSensors(); // Показываем уведомление о просроченных
            CheckApproachingSensors(); // Показываем уведомление о подходящие к сроку
        }

        private void CheckExpiredSensors()
        {
            var expired = Sensors.Where(s => s.ExpiryDate < DateTime.Today).ToList();

            if (expired.Any())
            {
                var message = new StringBuilder();
                message.AppendLine("Обнаружены просроченные датчики:");
                foreach (var sensor in expired.Take(5)) // Показываем первые 5
                {
                    message.AppendLine($"- {sensor.Name} (№{sensor.SerialNumber}) просрочен {sensor.ExpiryDate:dd.MM.yyyy}");
                }
                if (expired.Count > 5) message.AppendLine($"... и еще {expired.Count - 5}");

                MessageBox.Show(message.ToString(), "Внимание!",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// Проверка истекающих датчиков
        /// Показывает предупреждение, если такие есть
        /// </summary>

        private void CheckApproachingSensors()
        {
            var approaching = Sensors.Where(t =>
        t.ExpiryDate >= DateTime.Today &&
        t.ExpiryDate <= DateTime.Today.AddDays(30)).ToList();

            if (approaching.Any())
            {
                var message = new StringBuilder();
                message.AppendLine("Обнаружены датчики на проверку в следующем месяце:");
                foreach (var sensor in approaching.Take(5)) // Показываем первые 5
                {
                    message.AppendLine($"- {sensor.Name} (№{sensor.SerialNumber}) проверка {sensor.ExpiryDate:dd.MM.yyyy}");
                }
                if (approaching.Count > 5) message.AppendLine($"... и еще {approaching.Count - 5}");

                MessageBox.Show(message.ToString(), "Внимание!",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// Проверка, можно ли редактировать или удалять датчик
        /// (должен быть выбран хотя бы один датчик)
        /// </summary>
        private bool CanEditOrDelete(object obj) => SelectedSensor != null;

            /// <summary>
            /// Фильтр по дате
            /// </summary>
            private void FilterByDateRange(object obj)
            {
            // Применяем фильтр через механизм ICollectionView
            ApplyFilters();

            // Можно добавить дополнительную логику, например:
            if (obj is string period)
            {
                switch (period)
                {
                    case "month":
                        StartDate = DateTime.Today;
                        EndDate = DateTime.Today.AddMonths(1);
                        break;
                    case "year":
                        StartDate = DateTime.Today;
                        EndDate = DateTime.Today.AddYears(1);
                        break;
                }
                ApplyFilters();
            }
        }

            /// <summary>
            /// Добавление нового датчика
            /// </summary>
            private void AddSensor(object obj)
            {
                // Создаем диалоговое окно с новым датчиком
                var dialog = new SensorEditDialog(new Sensor
                {
                    PlacementDate = DateTime.Today,
                    ExpiryDate = DateTime.Today.AddYears(1) // По умолчанию срок - 1 год
                });

            // Если пользователь нажал "Сохранить"
            if (dialog.ShowDialog() == true)
            {
                // Добавляем датчик в коллекцию
                Sensors.Add(dialog.Sensor);
                // Сохраняем изменения
                SaveData();
                // После добавления нового элемента обновляем фильтр
                SelectedSensor = dialog.Sensor; // Выделяем новый элемент
                ApplyFilters(); // Принудительно обновляем фильтр

                // Если всё ещё не отображается, временно сбросим фильтры
                if (!SensorsView.Contains(dialog.Sensor))
                {
                    var currentSearch = SearchText;
                    var currentStart = StartDate;
                    var currentEnd = EndDate;

                    SearchText = string.Empty;
                    StartDate = DateTime.MinValue;
                    EndDate = DateTime.MaxValue;

                    ApplyFilters();

                    // Возвращаем фильтры после отображения
                    SearchText = currentSearch;
                    StartDate = currentStart;
                    EndDate = currentEnd;
                    ApplyFilters();
                }
            }
        }

        /// <summary>
        /// Редактирование выбранного датчика
        /// Теперь не вызывает проблем с фильтрацией, так как работаем с исходной коллекцией
        /// </summary>
        private void EditSensor(object obj)
        {
            if (SelectedSensor == null) return;

            // Создаем копию датчика для редактирования
            var sensorCopy = new Sensor
            {
                Name = SelectedSensor.Name,
                SerialNumber = SelectedSensor.SerialNumber,
                TypeSensor = SelectedSensor.TypeSensor,
                MeasurementLimits = SelectedSensor.MeasurementLimits,
                PlacementDate = SelectedSensor.PlacementDate,
                ClassForSure = SelectedSensor.ClassForSure,
                ExpiryDate = SelectedSensor.ExpiryDate,
                Location = SelectedSensor.Location,
                PlaceOfUse = SelectedSensor.PlaceOfUse
            };

            var dialog = new SensorEditDialog(sensorCopy);
            if (dialog.ShowDialog() == true)
            {
                // Обновляем свойства исходного датчика
                SelectedSensor.Name = sensorCopy.Name;
                SelectedSensor.SerialNumber = sensorCopy.SerialNumber;
                SelectedSensor.TypeSensor = sensorCopy.TypeSensor;
                SelectedSensor.MeasurementLimits = sensorCopy.MeasurementLimits;
                SelectedSensor.PlacementDate = sensorCopy.PlacementDate;
                SelectedSensor.Location = sensorCopy.Location;
                SelectedSensor.PlaceOfUse = sensorCopy.PlaceOfUse;
                SelectedSensor.ClassForSure = sensorCopy.ClassForSure;
                SelectedSensor.ExpiryDate = sensorCopy.ExpiryDate;

                SaveData();
                // Обновляем отображение после редактирования
                ApplyFilters();
            }
        }

        /// <summary>
        /// Удаление выбранного датчика
        /// </summary>
        private void DeleteSensor(object obj)
            {
                if (obj is KeyEventArgs keyArgs && keyArgs.Key != Key.Delete)
                return;

                if (SelectedSensor == null) return;

                // Запрос подтверждения
                if (MessageBox.Show("Удалить выбранный датчик?", "Подтверждение",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    // Удаляем датчик из коллекции
                    Sensors.Remove(SelectedSensor);
                    // Сохраняем изменения
                    SaveData();
                    // После удаления обновляем фильтр
                    ApplyFilters();
                }
                }

        /// <summary>
        /// Показать все датчики (сброс фильтров)
        /// </summary>
        private void ShowAllSensors(object obj)
        {
            SearchText = string.Empty;
            SetFullDateRange(); // Устанавливаем полный диапазон
            ApplyFilters();
        }

        /// <summary>
        /// Показать только просроченные датчики
        /// </summary>
        private void ShowApproachingSensors(object obj)
            {
                SearchText = "status:approaching";
                StartDate = DateTime.Today;
                EndDate = DateTime.Today.AddDays(30);
                ApplyFilters();
            }

        /// <summary>
        /// Показать просроченные датчики
        /// </summary>
        private void ShowExpiredSensors(object obj)
        {
            // Очищаем текстовый поиск
            SearchText = string.Empty;

            // Устанавливаем диапазон дат "все что раньше сегодня"
            StartDate = DateTime.MinValue;
            EndDate = DateTime.Today;

            // Применяем фильтр по статусу "Просрочено"
            SearchText = "status:expired";

            OnPropertyChanged(nameof(StartDate));
            OnPropertyChanged(nameof(EndDate));
            ApplyFilters();
        }


        // Метод асинхронной загрузки с прогрессом
        private async Task LoadDataAsync()
        {
            try
            {
                IsLoading = true; // Активируем индикатор загрузки
                ProgressMessage = "Загрузка данных...";

                await Task.Run(() => // Запуск в фоновом потоке
                {
                    if (File.Exists(DataFilePath))
                    {
                        var serializer = new XmlSerializer(typeof(List<Sensor>));
                        using (var stream = File.OpenRead(DataFilePath))
                        {
                            var data = (List<Sensor>)serializer.Deserialize(stream);

                            // Пакетное добавление для минимизации обращений к UI-потоку
                            int batchSize = 500;
                            for (int i = 0; i < data.Count; i += batchSize)
                            {
                                var batch = data.Skip(i).Take(batchSize).ToList();

                                // Обновление UI через Dispatcher
                                Application.Current.Dispatcher.Invoke(() =>
                                {
                                    foreach (var sensor in batch)
                                    {
                                        Sensors.Add(sensor);
                                    }

                                    // Автоматически устанавливаем полный диапазон дат после загрузки
                                    SetFullDateRange();
                                });

                                ProgressValue = (double)i / data.Count * 100; // Обновление прогресса
                                Thread.Sleep(50); // Искусственная задержка для плавности UI
                            }
                        }
                    }
                });
            }
            finally
            {
                IsLoading = false; // Выключаем индикатор в любом случае
                ApplyFilters(); // Применяем фильтры после загрузки
            }
        }



            /// <summary>
            /// Экспорт данных в CSV файл
            /// </summary>
            private void ExportSensors(object obj)
            {
                // Диалог сохранения файла
                var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "CSV файл (*.csv)|*.csv",
                    Title = "Экспорт данных в CSV"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    try
                    {// Получаем отфильтрованные данные
                    var filteredSensors = SensorsView.Cast<Sensor>().ToList();
                    // Запись данных в файл
                    using (var writer = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                    {
                            // Заголовки столбцов
                            writer.WriteLine("Название;Заводской номер;Дата размещения;Срок хранения;Стенд;Место хранения;Количество;Назначение;Статус");

                            // Данные по каждому датчику
                            foreach (var sensor in filteredSensors)
                            {
                                writer.WriteLine(
                                    $"\"{sensor.Name}\";\"{sensor.TypeSensor}\";\"{sensor.SerialNumber}\";" +
                                    $"\"{sensor.MeasurementLimits}\";\"{sensor.PlacementDate:dd.MM.yyyy}\";\"{sensor.ClassForSure}\";"+
                                    $"\"{sensor.ExpiryDate:dd.MM.yyyy}\";\"{sensor.Location}\";" +
                                    $"\"{sensor.PlaceOfUse}\";\"{sensor.ExpiryStatus}\"");
                            }
                        }

                        MessageBox.Show("Данные успешно экспортированы", "Экспорт завершен",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }

            /// <summary>
            /// Экспорт данных Xlslx файла
            /// </summary>
            private void ExportToExcel()
            {
                var dialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = $"ПереченьСИ_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
                    Title = "Экспорт в Excel"
                };

                if (dialog.ShowDialog() == true)
                {
                    try
                    {
                        // Получаем текущие отфильтрованные данные
                        var filteredSensors = SensorsView.Cast<Sensor>().ToList();

                        ExcelExportService.ExportSensorsToExcel(filteredSensors, dialog.FileName);
                    
                        if (MessageBox.Show("Экспорт завершен успешно! Открыть файл?", 
                            "Готово", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                        {
                            Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка", 
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }

        /// <summary>
        /// Экспорт данных Xlslx файла
        /// </summary>
        private void ExportReportToExcel()
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                FileName = $"Отчет_ПеречениСИ_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
                Title = "Экспорт в Excel"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    // Получаем текущие отфильтрованные данные
                    var filteredSensors = SensorsView.Cast<Sensor>().ToList();

                    ExcelExportServiceReport.ExportReportToExcel(filteredSensors, dialog.FileName);

                    if (MessageBox.Show("Экспорт завершен успешно! Открыть файл?",
                        "Готово", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        /// <summary>
        /// Импорт данных Xlslx файла
        /// </summary>
        private void ImportFromExcel()
            {
                var dialog = new OpenFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    Title = "Импорт из Excel"
                };

                if (dialog.ShowDialog() == true)
                {
                    try
                    {
                        var importedSensors = ExcelImportService.ImportSensorFromExcel(dialog.FileName);
                        Sensors.Clear();
                        foreach (var sensor in importedSensors)
                        {
                            Sensors.Add(sensor);
                        }
                        MessageBox.Show($"Успешно импортировано {importedSensors.Count} средств измерения",
                            "Импорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
                        SaveData();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка импорта: {ex.Message}", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }


            /// <summary>
            /// Сохранение данных в XML файл
            /// </summary>
            private void SaveData()
            {
                try
                {
                    var serializer = new XmlSerializer(typeof(List<Sensor>));
                    using (var stream = File.Create(DataFilePath))
                    {
                        serializer.Serialize(stream, Sensors.ToList());
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка сохранения данных: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        /// <summary>
        /// Событие изменения свойства (для INotifyPropertyChanged)
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Метод вызова события PropertyChanged
        /// </summary>
        /// <param name="propertyName">Имя изменившегося свойства</param>
        protected virtual void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        /// <summary>
        /// Реализация команды ICommand для MVVM
        /// </summary>
        public class RelayCommand : ICommand
        {
            private readonly Action<object> _execute;
            private readonly Predicate<object> _canExecute;
            private Action exportToExcel;

            /// <summary>
            /// Конструктор команды
            /// </summary>
            /// <param name="execute">Метод выполнения</param>
            /// <param name="canExecute">Метод проверки возможности выполнения (может быть null)</param>
            public RelayCommand(Action<object> execute, Predicate<object> canExecute = null)
            {
                _execute = execute ?? throw new ArgumentNullException(nameof(execute));
                _canExecute = canExecute;
            }

            /// <summary>
            /// Проверка возможности выполнения команды
            /// </summary>
            public bool CanExecute(object parameter) => _canExecute == null || _canExecute(parameter);

            /// <summary>
            /// Выполнение команды
            /// </summary>
            public void Execute(object parameter) => _execute(parameter);

            /// <summary>
            /// Событие изменения возможности выполнения команды
            /// </summary>
            public event EventHandler CanExecuteChanged
            {
                add => CommandManager.RequerySuggested += value;
                remove => CommandManager.RequerySuggested -= value;
            }

        
        }
    }