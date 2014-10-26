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
    private IObservable<Unit> onGoToReportClicked;
    private IObservable<WorkItemFilter> onTaskFilterChanged;
    private IObservable<WorkItemInfo> onTaskDoubleClicked;

    public static readonly DependencyProperty BusyProperty =
      DependencyProperty.Register("Busy", typeof(bool), typeof(TFSTaskPane));

    public bool Busy
      {
      get { return (bool)this.GetValue(BusyProperty); }
      set { this.SetValue(BusyProperty, value); }
      }

    public TFSTaskPane()
      {
      InitializeComponent();

      onConnectToTfs = Observable.FromEventPattern(ConnectToTFS, "Click").Select(_ => Unit.Default);

      onTaskFilterChanged = Observable.Merge(
          Observable.FromEventPattern<SelectionChangedEventArgs>(TFSProjects, "SelectionChanged").Select(e => _GetCurrentFilter()),
          Observable.FromEventPattern(RefreshButton, "Click").Select(e => _GetCurrentFilter()))
        .Where(f => f.Project != null)
        .Do(e => SetBusy(true));

      onTaskDoubleClicked = Observable.FromEventPattern<MouseButtonEventArgs>(TFSTasks, "MouseDoubleClick")
        .Select(e => ItemsControl.ContainerFromElement(e.Sender as ListBox, e.EventArgs.OriginalSource as DependencyObject) as ListBoxItem)
        .Where(item => item != null)
        .Select(item => (WorkItemInfo)TFSTasks.ItemContainerGenerator.ItemFromContainer(item));

      onGoToReportClicked = Observable.FromEventPattern(GoToReportWebsite, "Click").Select(_ => Unit.Default);
      }

    public IObservable<Unit> OnConnectToTfs()
      {
      return onConnectToTfs;
      }

    public IObservable<Unit> OnGoToReport()
      {
      return onGoToReportClicked;
      }

    public IObservable<WorkItemFilter> OnTaskFilterChanged()
      {
      return onTaskFilterChanged;
      }

    public IObservable<WorkItemInfo> OnTaskDoubleClicked()
      {
      return onTaskDoubleClicked;
      }

    public void SetTasksList(IEnumerable<WorkItemInfo> tasks)
      {
      TFSTasks.ItemsSource = tasks;
      }

    public void SetBusy(bool busy)
      {
      Busy = busy;
      }

    private WorkItemFilter _GetCurrentFilter()
      {
      return new WorkItemFilter()
      {
        Project = TFSProjects.SelectedItem != null ? TFSProjects.SelectedItem.ToString() : null,
        ShowTasks = ShowTasks.IsChecked.GetValueOrDefault(false),
        ShowBugs = ShowBugs.IsChecked.GetValueOrDefault(false),
        ShowProposed = ShowProposed.IsChecked.GetValueOrDefault(false),
        ShowActive = ShowActive.IsChecked.GetValueOrDefault(false),
        ShowResolved = ShowResolved.IsChecked.GetValueOrDefault(false),
        ShowClosed = ShowClosed.IsChecked.GetValueOrDefault(false),
      };
      }
    }
  }
