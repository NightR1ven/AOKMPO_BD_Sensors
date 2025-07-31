using System.Configuration;
using System.Data;
using System.Windows;
using System;
using System.Globalization;
using System.Windows.Data;

namespace AOKMPO_BD_Sensors
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Инициализация ресурсов или сервисов
            InitializeServices();

            // Обработка необработанных исключений
            DispatcherUnhandledException += App_DispatcherUnhandledException;
        }

        private void InitializeServices()
        {
            // Здесь можно инициализировать сервисы, БД и т.д.
        }

        private void App_DispatcherUnhandledException(object sender,
            System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            // Логирование и обработка необработанных исключений
            MessageBox.Show($"Произошла ошибка: {e.Exception.Message}",
                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // Очистка ресурсов при выходе
            base.OnExit(e);
        }
    }

}
