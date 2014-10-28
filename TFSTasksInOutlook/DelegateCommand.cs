using System;
using System.Windows.Input;

namespace TFSTasksInOutlook
  {
  class DelegateCommand : ICommand
    {
    Predicate<object> canExecute;
    Action<object> execute;
    public DelegateCommand(Predicate<object> _canexecute, Action<object> _execute)
      : this()
      {
      canExecute = _canexecute;
      execute = _execute;
      }
    public DelegateCommand() { }
    public bool CanExecute(object parameter)
      {
      return canExecute == null ? true : canExecute(parameter);
      }
    public event EventHandler CanExecuteChanged;
    public void Execute(object parameter)
      {
      execute(parameter);
      }
    }
  }
