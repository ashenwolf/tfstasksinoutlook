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
    private IObservable<long> onAddFavTask;
    private IObservable<WorkItemInfo> onRemoveFavorite;

    private ICommand removeFavorite;

    public class RemoveFavoriteEventArgs: EventArgs {
      public RemoveFavoriteEventArgs(WorkItemInfo item) { Item = item; }
      public WorkItemInfo Item { get; private set; }
      }
    public event EventHandler<RemoveFavoriteEventArgs> OnRemoveFavoriteEvent;

    public static readonly DependencyProperty BusyProperty =
      DependencyProperty.Register("Busy", typeof(bool), typeof(TFSTaskPane));

    public static readonly DependencyProperty BusyAddFavProperty =
      DependencyProperty.Register("BusyAddFav", typeof(bool), typeof(TFSTaskPane));

    public bool BusyGetTasks
      {
      get { return (bool)this.GetValue(BusyProperty); }
      set { this.SetValue(BusyProperty, value); }
      }

    public bool BusyAddFav
      {
      get { return (bool)this.GetValue(BusyAddFavProperty); }
      set { this.SetValue(BusyAddFavProperty, value); }
      }

    public ICommand RemoveFavorite
      {
      get
        {
        return removeFavorite ?? (removeFavorite = new DelegateCommand(null, p =>
          {
            OnRemoveFavoriteEvent(this, new RemoveFavoriteEventArgs(p as WorkItemInfo));
          }));
        }
      }

    public TFSTaskPane()
      {
      InitializeComponent();
      BusyGetTasks = false;
      BusyAddFav = false;

      onConnectToTfs = Observable.FromEventPattern(ConnectToTFS, "Click").Select(_ => Unit.Default);

      onTaskFilterChanged = Observable.Merge(
          Observable.FromEventPattern<SelectionChangedEventArgs>(TFSProjects, "SelectionChanged").Select(e => _GetCurrentFilter()),
          Observable.FromEventPattern(RefreshButton, "Click").Select(e => _GetCurrentFilter()))
        .Where(f => f.Project != null)
        .Do(_ => SetBusyGetTasks(true));

      onTaskDoubleClicked = Observable.Merge(
          Observable.FromEventPattern<MouseButtonEventArgs>(TFSTasks, "MouseDoubleClick")
            .Select(e => ItemsControl.ContainerFromElement(e.Sender as ListBox, e.EventArgs.OriginalSource as DependencyObject) as ListBoxItem)
            .Where(item => item != null)
            .Select(item => (WorkItemInfo)TFSTasks.ItemContainerGenerator.ItemFromContainer(item)),
          Observable.FromEventPattern<MouseButtonEventArgs>(FavouriteTasks, "MouseDoubleClick")
            .Select(e => ItemsControl.ContainerFromElement(e.Sender as ListBox, e.EventArgs.OriginalSource as DependencyObject) as ListBoxItem)
            .Where(item => item != null)
            .Select(item => (WorkItemInfo)FavouriteTasks.ItemContainerGenerator.ItemFromContainer(item))
        );

      onAddFavTask = Observable.Merge(
        Observable.FromEventPattern(AddFavTask, "Click").Select(_ => Unit.Default),
        Observable.FromEventPattern<KeyEventArgs>(NewTaskID, "PreviewKeyDown").Where(e => e.EventArgs.Key == Key.Enter).Select(_ => Unit.Default))
        .Where(e => NewTaskID.Text.Trim() != "" && NewTaskID.Text.Count() <= 12)
        .Select(e => Convert.ToInt64(NewTaskID.Text))
        .Do(_ => NewTaskID.Text = "");

      onGoToReportClicked = Observable.FromEventPattern(GoToReportWebsite, "Click").Select(_ => Unit.Default);

      onRemoveFavorite = Observable.FromEventPattern<RemoveFavoriteEventArgs>(this, "OnRemoveFavoriteEvent")
        .Where(e => e.EventArgs.Item != null)
        .Select(e => e.EventArgs.Item);
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

    public IObservable<long> OnAddFavTask()
      {
      return onAddFavTask;
      }

    public IObservable<WorkItemInfo> OnRemoveFavorite()
      {
      return onRemoveFavorite;
      }

    public void SetTasksList(IEnumerable<WorkItemInfo> tasks)
      {
      TFSTasks.ItemsSource = tasks;
      }

    public void SetFavTaskList(IEnumerable<WorkItemInfo> tasks)
      {
      FavouriteTasks.ItemsSource = tasks;
      }
    
    public void SetBusyGetTasks(bool busy)
      {
      BusyGetTasks = busy;
      }

    public void SetBusyAddFav(bool busy)
      {
      BusyAddFav = busy;
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
