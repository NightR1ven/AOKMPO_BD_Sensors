using AOKMPO_BD_Sensors.Service;
using AOKMPO_BD_Sensors.ViewModels;
using AOKMPO_BD_Sensors.Views;
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
    /// ViewModel главного окна. Содержит всю бизнес-логику приложения:
    /// загрузка/сохранение данных, фильтрация, команды для кнопок, статус-бар.
    /// </summary>
    public class MainViewModel : INotifyPropertyChanged
    {
        #region Поля и константы
        private const string DataFilePath = "sensors.xml";   // Файл для хранения данных
        private Sensor _selectedSensor;                      // Выбранный датчик в списке
        private string _searchText;                          // Текст поиска
        private DateTime _startDate = DateTime.Today.AddMonths(-1); // Начальная дата фильтра
        private DateTime _endDate = DateTime.Today.AddMonths(1);    // Конечная дата фильтра
        private bool _isLoading;                             // Флаг загрузки данных
        private double _progressValue;                       // Значение прогресс-бара
        private string _progressMessage;                     // Сообщение прогресса
        private string _selectedPeriod = "all";              // Выбранный период фильтрации
        private readonly object _collectionLock = new object(); // Для потокобезопасности (если потребуется)

        // Свойства для статус-бара
        private string _statusText = "Готов";
        private int _recordCount;
        private string _lastUpdate;
        #endregion

        #region Коллекции и представление
        /// <summary>
        /// Исходная коллекция датчиков (все записи).
        /// </summary>
        public ObservableCollection<Sensor> Sensors { get; } = new ObservableCollection<Sensor>();

        /// <summary>
        /// Представление коллекции с фильтрацией. Ленивая инициализация.
        /// </summary>
        private ICollectionView _sensorsView;
        public ICollectionView SensorsView
        {
            get
            {
                if (_sensorsView == null)
                {
                    _sensorsView = CollectionViewSource.GetDefaultView(Sensors);
                    _sensorsView.Filter = SensorFilter;   // Устанавливаем фильтр
                }
                return _sensorsView;
            }
        }
        #endregion

        #region Свойства для привязки к UI
        public bool IsLoading
        {
            get => _isLoading;
            set { _isLoading = value; OnPropertyChanged(nameof(IsLoading)); }
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

        public Sensor SelectedSensor
        {
            get => _selectedSensor;
            set { _selectedSensor = value; OnPropertyChanged(nameof(SelectedSensor)); }
        }

        public DateTime StartDate
        {
            get => _startDate;
            set { _startDate = value; OnPropertyChanged(nameof(StartDate)); ApplyFilters(); }
        }

        public DateTime EndDate
        {
            get => _endDate;
            set { _endDate = value; OnPropertyChanged(nameof(EndDate)); ApplyFilters(); }
        }

        /// <summary>
        /// Текст поиска. При изменении автоматически обновляется фильтр.
        /// </summary>
        public string SearchText
        {
            get => _searchText;
            set { _searchText = value; OnPropertyChanged(nameof(SearchText)); ApplyFilters(); }
        }

        // Свойства для статус-бара (добавлены для отображения внизу окна)
        public string StatusText
        {
            get => _statusText;
            set { _statusText = value; OnPropertyChanged(nameof(StatusText)); }
        }

        public int RecordCount
        {
            get => _recordCount;
            set { _recordCount = value; OnPropertyChanged(nameof(RecordCount)); }
        }

        public string LastUpdate
        {
            get => _lastUpdate;
            set { _lastUpdate = value; OnPropertyChanged(nameof(LastUpdate)); }
        }
        #endregion

        #region Команды (привязка к кнопкам)
        public ICommand AddCommand { get; }
        public ICommand EditCommand { get; }
        public ICommand DeleteCommand { get; }
        public ICommand ShowAllCommand { get; }
        public ICommand ShowExpiringCommand { get; }
        public ICommand ExportCommand { get; }
        public ICommand ExportToExcelCommand { get; }
        public ICommand ImportFromExcelCommand { get; }
        public ICommand ExportToExcelReportCommand { get; }
        public ICommand FilterByDateCommand { get; }
        public ICommand SearchCommand { get; }          // Команда для кнопки поиска или нажатия Enter
        #endregion

        #region Периоды фильтрации (для выпадающего списка)
        /// <summary>
        /// Доступные предустановленные периоды. Словарь, где ключ – идентификатор,
        /// значение – объект с именем и действием для установки диапазона.
        /// </summary>
        public Dictionary<string, DatePeriod> DatePeriods { get; } = new()
        {
            ["all"] = new DatePeriod { Name = "Все даты", Action = vm => vm.SetFullDateRange() },
            ["month"] = new DatePeriod { Name = "Текущий месяц", Action = vm => vm.SetCurrentMonthRange() },
            ["year"] = new DatePeriod { Name = "Текущий год", Action = vm => vm.SetCurrentYearRange() },
            ["custom"] = new DatePeriod { Name = "Произвольный диапазон", Action = vm => { } }
        };

        public string SelectedPeriod
        {
            get => _selectedPeriod;
            set
            {
                _selectedPeriod = value;
                OnPropertyChanged(nameof(SelectedPeriod));

                if (DatePeriods.TryGetValue(value, out var period))
                {
                    period.Action(this);   // Устанавливаем даты в зависимости от выбранного периода
                    ApplyFilters();        // Обновляем фильтр
                }
            }
        }
        #endregion

        #region Конструктор и инициализация
        public MainViewModel()
        {
            // Инициализация всех команд
            AddCommand = new RelayCommand(_ => AddSensor());
            EditCommand = new RelayCommand(_ => EditSensor(), _ => SelectedSensor != null);
            DeleteCommand = new RelayCommand(_ => DeleteSensor(), _ => SelectedSensor != null);
            ShowAllCommand = new RelayCommand(_ => ShowAllSensors());
            ShowExpiringCommand = new RelayCommand(_ => ShowApproachingSensors());
            ExportToExcelCommand = new RelayCommand(_ => ExportToExcel());
            ExportToExcelReportCommand = new RelayCommand(_ => ExportReportToExcel());
            ImportFromExcelCommand = new RelayCommand(_ => ImportFromExcel());
            FilterByDateCommand = new RelayCommand(FilterByDateRange);
            SearchCommand = new RelayCommand(_ => ApplyFilters());  // При нажатии Enter – обновляем фильтр

            // Запускаем асинхронную загрузку данных и проверку просроченных датчиков
            _ = LoadDataAndCheckExpiredAsync();
        }
        #endregion

        #region Фильтрация
        /// <summary>
        /// Универсальный фильтр для датчиков. Вызывается для каждого элемента коллекции.
        /// </summary>
        private bool SensorFilter(object item)
        {
            if (item is not Sensor sensor) return false;

            // Проверка по дате
            bool dateMatch = FilterByDate(sensor);

            // Если поисковый запрос пуст – возвращаем только результат фильтрации по дате
            if (string.IsNullOrEmpty(SearchText))
                return dateMatch;

            // Специальный запрос для истекающих датчиков (используется кнопкой "Истекающие")
            if (SearchText == "status:approaching")
                return dateMatch && sensor.ExpiryDate <= DateTime.Today.AddDays(30);

            // Обычный текстовый поиск (регистронезависимый)
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
                (sensor.SerialNumber != null && Regex.IsMatch(sensor.SerialNumber, searchTerm, RegexOptions.IgnoreCase)) ||
                sensor.PlacementDate.ToString("dd.MM.yyyy").Contains(searchTerm) ||
                sensor.ExpiryDate.ToString("dd.MM.yyyy").Contains(searchTerm);

            return dateMatch && textMatch;
        }

        /// <summary>
        /// Фильтрация по диапазону дат.
        /// </summary>
        private bool FilterByDate(Sensor sensor)
        {
            if (SelectedPeriod == "all") return true;
            return sensor.ExpiryDate >= StartDate && sensor.ExpiryDate <= EndDate;
        }

        /// <summary>
        /// Принудительное обновление фильтра и статус-бара.
        /// </summary>
        private void ApplyFilters()
        {
            if (_sensorsView != null)
            {
                _sensorsView.Refresh();
                UpdateStatusBar();
            }
        }

        /// <summary>
        /// Обновляет информацию в статус-баре (количество записей, время последнего обновления).
        /// </summary>
        private void UpdateStatusBar()
        {
            RecordCount = _sensorsView?.Cast<object>().Count() ?? 0;
            StatusText = "Готов";
            LastUpdate = DateTime.Now.ToString("HH:mm:ss");
        }
        #endregion

        #region Установка диапазонов дат
        public void SetFullDateRange()
        {
            if (Sensors.Count == 0)
            {
                StartDate = DateTime.Today.AddMonths(-1);
                EndDate = DateTime.Today.AddMonths(1);
                return;
            }
            StartDate = Sensors.Min(s => s.ExpiryDate).AddDays(-7);
            EndDate = Sensors.Max(s => s.ExpiryDate).AddDays(7);
        }

        public void SetCurrentMonthRange()
        {
            var today = DateTime.Today;
            StartDate = new DateTime(today.Year, today.Month, 1);
            EndDate = StartDate.AddMonths(1).AddDays(-1);
        }

        public void SetCurrentYearRange()
        {
            var today = DateTime.Today;
            StartDate = new DateTime(today.Year, 1, 1);
            EndDate = new DateTime(today.Year, 12, 31);
        }
        #endregion

        #region Команды: добавление, редактирование, удаление
        private void AddSensor()
        {
            var newSensor = new Sensor
            {
                PlacementDate = DateTime.Today,
                ExpiryDate = DateTime.Today.AddYears(1),
                PlaceOfDoc = "не известно"
            };
            var dialog = new SensorEditDialog(newSensor);
            if (dialog.ShowDialog() == true)
            {
                Sensors.Add(dialog.Sensor);
                SaveData();
                SelectedSensor = dialog.Sensor;
                ApplyFilters();
            }
        }

        private void EditSensor()
        {
            if (SelectedSensor == null) return;
            // Создаём копию для редактирования, чтобы изменения можно было отменить
            var copy = new Sensor
            {
                Name = SelectedSensor.Name,
                SerialNumber = SelectedSensor.SerialNumber,
                TypeSensor = SelectedSensor.TypeSensor,
                MeasurementLimits = SelectedSensor.MeasurementLimits,
                PlacementDate = SelectedSensor.PlacementDate,
                ClassForSure = SelectedSensor.ClassForSure,
                ExpiryDate = SelectedSensor.ExpiryDate,
                Location = SelectedSensor.Location,
                PlaceOfUse = SelectedSensor.PlaceOfUse,
                PlaceOfDoc = SelectedSensor.PlaceOfDoc
            };
            var dialog = new SensorEditDialog(copy);
            if (dialog.ShowDialog() == true)
            {
                // Копируем изменённые данные обратно в исходный объект
                SelectedSensor.Name = copy.Name;
                SelectedSensor.SerialNumber = copy.SerialNumber;
                SelectedSensor.TypeSensor = copy.TypeSensor;
                SelectedSensor.MeasurementLimits = copy.MeasurementLimits;
                SelectedSensor.PlacementDate = copy.PlacementDate;
                SelectedSensor.Location = copy.Location;
                SelectedSensor.PlaceOfUse = copy.PlaceOfUse;
                SelectedSensor.ClassForSure = copy.ClassForSure;
                SelectedSensor.ExpiryDate = copy.ExpiryDate;
                SelectedSensor.PlaceOfDoc = copy.PlaceOfDoc;

                SaveData();
                ApplyFilters();
            }
        }

        private void DeleteSensor()
        {
            if (SelectedSensor == null) return;
            if (MessageBox.Show("Удалить выбранный датчик?", "Подтверждение",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                Sensors.Remove(SelectedSensor);
                SaveData();
                ApplyFilters();
            }
        }
        #endregion

        #region Команды: показать всё / истекающие
        private void ShowAllSensors()
        {
            SearchText = string.Empty;
            SetFullDateRange();
            ApplyFilters();
        }

        private void ShowApproachingSensors()
        {
            SearchText = "status:approaching";
            StartDate = DateTime.Today;
            EndDate = DateTime.Today.AddDays(30);
            ApplyFilters();
        }
        #endregion

        #region Команды: фильтр по дате (для совместимости)
        private void FilterByDateRange(object obj)
        {
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
        #endregion

        #region Загрузка и сохранение данных (XML)
        private async Task LoadDataAndCheckExpiredAsync()
        {
            await LoadDataAsync();
            CheckExpiredSensors();
        }

        private async Task LoadDataAsync()
        {
            try
            {
                IsLoading = true;
                ProgressMessage = "Загрузка данных...";
                await Task.Run(() =>
                {
                    if (File.Exists(DataFilePath))
                    {
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(List<Sensor>));
                        using (var stream = File.OpenRead(DataFilePath))
                        {
                            var data = (List<Sensor>)serializer.Deserialize(stream);
                            int batchSize = 500;
                            for (int i = 0; i < data.Count; i += batchSize)
                            {
                                var batch = data.Skip(i).Take(batchSize).ToList();
                                Application.Current.Dispatcher.Invoke(() =>
                                {
                                    foreach (var sensor in batch)
                                        Sensors.Add(sensor);
                                    SetFullDateRange();
                                });
                                ProgressValue = (double)i / data.Count * 100;
                                Thread.Sleep(50); // Имитация прогресса (можно убрать)
                            }
                        }
                    }
                });
            }
            finally
            {
                IsLoading = false;
                ApplyFilters();
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
        #endregion

        #region Проверка истекающих датчиков при старте
        private void CheckExpiredSensors()
        {
            var approaching = Sensors.Where(t =>
                t.ExpiryDate >= DateTime.Today &&
                t.ExpiryDate <= DateTime.Today.AddDays(30)).ToList();
            if (approaching.Any())
            {
                var dialog = new ExpiredSensorsDialog(approaching);
                dialog.ShowDialog();
            }
        }
        #endregion


        /// <summary>
        /// Экспорт данных в Excel файл (XLSX)
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
        /// Экспорт отчёта в Excel (расширенная форма)
        /// </summary>
        private void ExportReportToExcel()
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                FileName = $"Отчет_ПереченьСИ_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
                Title = "Экспорт отчёта в Excel"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    // Получаем текущие отфильтрованные данные
                    var filteredSensors = SensorsView.Cast<Sensor>().ToList();

                    ExcelExportServiceReport.ExportReportToExcel(filteredSensors, dialog.FileName);

                    if (MessageBox.Show("Экспорт отчёта завершен успешно! Открыть файл?",
                        "Готово", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка экспорта отчёта: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        /// <summary>
        /// Импорт данных из Excel файла (XLSX)
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
                    ApplyFilters(); // Обновить отображение после импорта
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка импорта: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }


        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        #endregion
    }

    /// <summary>
    /// Класс для описания предустановленного периода фильтрации.
    /// </summary>
    public class DatePeriod
    {
        public string Name { get; set; }
        public Action<MainViewModel> Action { get; set; }
    }
}