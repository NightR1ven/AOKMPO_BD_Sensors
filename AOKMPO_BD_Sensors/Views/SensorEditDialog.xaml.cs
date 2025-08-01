using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace AOKMPO_BD_Sensors
{
    /// <summary>
    /// Диалоговое окно для редактирования датчика
    /// </summary>
    public partial class SensorEditDialog : Window
    {
        /// <summary>
        /// Редактируемый датчик
        /// </summary>
        public Sensor Sensor { get; }

        /// <summary>
        /// Конструктор окна редактирования
        /// </summary>
        /// <param name="sensor">Датчик для редактирования</param>
        public SensorEditDialog(Sensor sensor)
        {
            InitializeComponent();
            Sensor = sensor; // Сохраняем ссылку на датчик
            DataContext = this; // Устанавливаем контекст данных для привязок
        }

        /// <summary>
        /// Обработчик нажатия кнопки "Сохранить"
        /// </summary>
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка обязательных полей
            if (string.IsNullOrWhiteSpace(Sensor.Name))
            {
                MessageBox.Show("Введите название датчика", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(Sensor.SerialNumber))
            {
                MessageBox.Show("Введите заводской номер", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Устанавливаем результат диалога и закрываем окно
            DialogResult = true;
            Close();
        }
    }
}
