using AOKMPO_BD_Sensors.Serviec;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
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

        /// <summary>
        /// Коллекция датчиков (привязана к ListView)
        /// </summary>
        public ObservableCollection<Sensor> Sensors { get; } = new ObservableCollection<Sensor>();

        /// <summary>
        /// Выбранный датчик
        /// </summary>
        public Sensor SelectedSensor
        {
            get => _selectedSensor;
            set { _selectedSensor = value; OnPropertyChanged(nameof(SelectedSensor)); }
        }

        /// <summary>
        /// Текст для поиска датчиков
        /// При изменении автоматически фильтрует список
        /// </summary>
        public string SearchText
        {
            get => _searchText;
            set { _searchText = value; OnPropertyChanged(nameof(SearchText)); FilterSensors(); }
        }

        // Команды для привязки к кнопкам в UI
        public ICommand AddCommand { get; }       // Добавить датчик
        public ICommand EditCommand { get; }     // Редактировать датчиков
        public ICommand DeleteCommand { get; }    // Удалить датчиков
        public ICommand ShowAllCommand { get; }   // Показать все датчики
        public ICommand ShowExpiredCommand { get; } // Показать просроченные
        public ICommand ShowExpiringCommand { get; } // Показать истекающие
        public ICommand ExportCommand { get; }    // Экспорт в CSV
        public ICommand ExportToExcelCommand { get; }

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
            ExportCommand = new RelayCommand(ExportSensors);
            ExportToExcelCommand = new RelayCommand(_ => ExportToExcel());


            // Загрузка данных и проверка сроков
            LoadData();
            CheckExpiredSensors();
        }

        /// <summary>
        /// Проверка, можно ли редактировать или удалять датчик
        /// (должен быть выбран хотя бы один датчик)
        /// </summary>
        private bool CanEditOrDelete(object obj) => SelectedSensor != null;

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
            }
        }

        /// <summary>
        /// Редактирование выбранного датчика
        /// </summary>
        private void EditSensor(object obj)
        {
            // Создаем диалоговое окно с выбранным датчикок
            var dialog = new SensorEditDialog(SelectedSensor);
            if (dialog.ShowDialog() == true)
            {
                // Сохраняем изменения (данные обновляются через привязку)
                SaveData();
            }
        }

        /// <summary>
        /// Удаление выбранного датчика
        /// </summary>
        private void DeleteSensor(object obj)
        {
            // Запрос подтверждения
            if (MessageBox.Show("Удалить выбранный датчик?", "Подтверждение",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                // Удаляем датчик из коллекции
                Sensors.Remove(SelectedSensor);
                // Сохраняем изменения
                SaveData();
            }
        }

        /// <summary>
        /// Показать все датчики (сброс фильтров)
        /// </summary>
        private void ShowAllSensors(object obj)
        {
            SearchText = string.Empty; // Очистка текста поиска загрузит все данные
        }

        /// <summary>
        /// Показать только просроченные датчики
        /// </summary>
        private void ShowApproachingSensors(object obj)
        {
            SearchText = "status:approaching"; // Специальное значение для фильтрации
        }

        /// <summary>
        /// Показать датчики с истекающим сроком
        /// </summary>
        private void ShowExpiringSensors(object obj)
        {
            SearchText = "status:expiring"; // Специальное значение для фильтрации
        }

        /// <summary>
        /// Фильтрация датчиков по тексту поиска
        /// </summary>
        private void FilterSensors()
        {
            // Если строка поиска пустая - загружаем все данные
            if (string.IsNullOrWhiteSpace(SearchText))
            {
                LoadData();
                return;
            }

            // Фильтр для истекающих датчиков
            if (SearchText == "status:approaching")
            {
                var filtered = Sensors.Where(t =>
                    t.ExpiryDate >= DateTime.Today &&
                    t.ExpiryDate <= DateTime.Today.AddDays(30)).ToList();
                Sensors.Clear();
                foreach (var sensor in filtered)
                    Sensors.Add(sensor);
            }
            // Обычный поиск по всем полям
            else
            {
                var searchTerm = SearchText.ToLower();
                var filtered = Sensors.Where(t =>
                    t.Name.ToLower().Contains(searchTerm) ||
                    t.SerialNumber.ToLower().Contains(searchTerm) ||
                    t.PlaceOfUse.ToLower().Contains(searchTerm) ||
                    t.Location.ToLower().Contains(searchTerm) ||
                    t.TypeSensor.ToLower().Contains(searchTerm) ||
                    t.ClassForSure.ToLower().Contains(searchTerm) ||
                    t.MeasurementLimits.ToLower().Contains(searchTerm));

                // Обновляем коллекцию с учетом фильтра
                Sensors.Clear();
                foreach (var sensor in filtered)
                    Sensors.Add(sensor);
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
                {
                    // Запись данных в файл
                    using (var writer = new StreamWriter(saveFileDialog.FileName))
                    {
                        // Заголовки столбцов
                        writer.WriteLine("Название;Заводской номер;Дата размещения;Срок хранения;Стенд;Место хранения;Количество;Назначение;Статус");

                        // Данные по каждому датчику
                        foreach (var sensor in Sensors)
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
        /// Загрузка данных из XML файла
        /// </summary>
        private void LoadData()
        {
            if (File.Exists(DataFilePath))
            {
                try
                {
                    var serializer = new XmlSerializer(typeof(List<Sensor>));
                    using (var stream = File.OpenRead(DataFilePath))
                    {
                        var data = (List<Sensor>)serializer.Deserialize(stream);
                        Sensors.Clear();
                        foreach (var sensor in data)
                            Sensors.Add(sensor);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ExportToExcel()
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                FileName = $"Инструменты_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
                Title = "Экспорт в Excel"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    ExcelExportService.ExportSensorsToExcel(Sensors, dialog.FileName);
                    
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
        /// Проверка просроченных и истекающих датчиков
        /// Показывает предупреждение, если такие есть
        /// </summary>
        private void CheckExpiredSensors()
        {
            var approaching = Sensors.Where(t =>
        t.ExpiryDate >= DateTime.Today &&
        t.ExpiryDate <= DateTime.Today.AddDays(30)).ToList();

            if (approaching.Any())
            {
                var message = new System.Text.StringBuilder();
                message.AppendLine("Скоро истекает срок у следующих датчиков:");
                foreach (var sensor in approaching)
                    message.AppendLine($"- {sensor.Name} (№{sensor.SerialNumber}) - до {sensor.ExpiryDate:dd.MM.yyyy}");

                MessageBox.Show(message.ToString(), "Внимание! Проверьте сроки",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
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