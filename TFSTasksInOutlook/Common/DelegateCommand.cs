using System;
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
    public bool CanExecute(object parameter)
      {
      return _canExecute == null || _canExecute(parameter);
      }
    public event EventHandler CanExecuteChanged;
    public void Execute(object parameter)
      {
      _execute(parameter);
      }
    }
  }
