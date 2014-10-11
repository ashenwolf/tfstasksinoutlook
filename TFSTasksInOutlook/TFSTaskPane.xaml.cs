using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reactive;
using System.Reactive.Concurrency;
using System.Reactive.Linq;
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

namespace TFSTasksInOutlook
{
  /// <summary>
  /// Interaction logic for TFSTaskPane.xaml
  /// </summary>
  public partial class TFSTaskPane : UserControl, ITFSTaskPaneView
  {
    private IObservable<Unit> onConnectToTfs;
    private IObservable<string> onProjectSelected;
    private IObservable<WorkItemInfo> onTaskDoubleClicked;
    private Point startPoint;

    public TFSTaskPane()
    {
      InitializeComponent();

      onConnectToTfs = Observable.FromEventPattern(ConnectToTFS, "Click").Select(_ => Unit.Default);

      onProjectSelected = Observable.FromEventPattern<SelectionChangedEventArgs>(TFSProjects, "SelectionChanged")
        .Throttle(TimeSpan.FromSeconds(1), Scheduler.CurrentThread)
        .Select(e => e.EventArgs.AddedItems.Cast<string>().FirstOrDefault());

      onTaskDoubleClicked = Observable.FromEventPattern<MouseButtonEventArgs>(TFSTasks, "MouseDoubleClick")
        .Select(e => ItemsControl.ContainerFromElement(e.Sender as ListBox, e.EventArgs.OriginalSource as DependencyObject) as ListBoxItem)
        .Where(item => item != null)
        .Select(item => (WorkItemInfo) TFSTasks.ItemContainerGenerator.ItemFromContainer(item));
    }

    public IObservable<Unit> OnConnectToTfs()
    {
      return onConnectToTfs;
    }

    public IObservable<string> OnProjectSelected()
    {
      return onProjectSelected;
    }

    public IObservable<WorkItemInfo> OnTaskDoubleClicked()
      {
      return onTaskDoubleClicked;
      }

    public void SetTasksList(IEnumerable<WorkItemInfo> tasks)
    {
      TFSTasks.ItemsSource = tasks;
    }
  }
}
