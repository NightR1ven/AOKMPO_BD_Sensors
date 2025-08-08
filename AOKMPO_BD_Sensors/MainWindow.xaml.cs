using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.ComponentModel;
using System.Collections.ObjectModel;

namespace AOKMPO_BD_Sensors
{

    /// <summary>
    /// Главное окно приложения
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Конструктор главного окна
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = new MainViewModel();
        }

        /// <summary>
        /// Обработчик двойного клика по списку датчиков
        /// </summary>
        private void ListView_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // Игнорируем клики по заголовкам колонок
            if (e.OriginalSource is GridViewColumnHeader) return;

            // Получаем ViewModel
            var viewModel = (MainViewModel)DataContext;
            // Если команда редактирования доступна - выполняем ее
            if (viewModel.EditCommand.CanExecute(null))
                viewModel.EditCommand.Execute(null);
        }
    }
}