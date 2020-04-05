using System;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Input;

namespace TFSTasksInOutlook.Common
{
    internal class DelegateCommand : ICommand
    {
        readonly Predicate<object> _canExecute;
        readonly Action<object> _execute;
        public DelegateCommand(Predicate<object> canexecute, Action<object> execute)
          : this()
        {
            _canExecute = canexecute;
            _execute = execute;
        }
        public DelegateCommand() { }
        bool ICommand.CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute(parameter);
        }
        [SuppressMessage("Build","CS0067")]
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
        void ICommand.Execute(object parameter)
        {
            _execute(parameter);
        }
    }
}
