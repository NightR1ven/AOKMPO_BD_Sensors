using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace AOKMPO_BD_Sensors.ViewModels
{
    /// <summary>
    /// Базовая реализация команды ICommand для MVVM.
    /// Позволяет связывать методы ViewModel с действиями в XAML.
    /// </summary>
    public class RelayCommand : ICommand
    {
        private readonly Action<object> _execute;      // Метод, выполняемый командой
        private readonly Predicate<object> _canExecute; // Метод, определяющий доступность команды

        /// <summary>
        /// Конструктор команды.
        /// </summary>
        /// <param name="execute">Метод, вызываемый при выполнении команды.</param>
        /// <param name="canExecute">Метод, проверяющий возможность выполнения (опционально).</param>
        public RelayCommand(Action<object> execute, Predicate<object> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        /// <summary>
        /// Определяет, может ли команда быть выполнена в данный момент.
        /// </summary>
        public bool CanExecute(object parameter) => _canExecute == null || _canExecute(parameter);

        /// <summary>
        /// Выполняет команду.
        /// </summary>
        public void Execute(object parameter) => _execute(parameter);

        /// <summary>
        /// Событие, возникающее при изменении возможности выполнения команды.
        /// Подписывается на глобальное событие CommandManager.RequerySuggested.
        /// </summary>
        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
    }
}

