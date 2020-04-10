using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive;
using System.Reactive.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using TFSTasksInOutlook.Common;

namespace TFSTasksInOutlook
{
    /// <summary>
    /// Interaction logic for TFSTaskPane.xaml
    /// </summary>
    public partial class TfsTaskPane : UserControl, ITfsTaskPaneView
    {
        #region Properties and Commands
        private readonly IObservable<Unit> _onConnectToTfs;
        private readonly IObservable<Unit> _onGoToReportClicked;
        private readonly IObservable<WorkItemFilter> _onTaskFilterChanged;
        private readonly IObservable<IEnumerable<WorkItemInfo>> _onAddTasksToCalendar;
        private readonly IObservable<long> _onAddFavTask;
        private readonly IObservable<WorkItemInfo> _onRemoveFavorite;
        private readonly IObservable<WorkItemInfo> _onCopyToClipboard;
        private readonly IObservable<IDataObject> _onDropItem;

        private ICommand _removeFavorite;
        private ICommand _copyToClipboard;

        public event EventHandler<WorkItemActionEventArgs> OnRemoveFavoriteEvent;
        public event EventHandler<WorkItemActionEventArgs> OnCopyToClipboardEvent;

        public static readonly DependencyProperty BusyProperty =
          DependencyProperty.Register("Busy", typeof(bool), typeof(TfsTaskPane));

        public static readonly DependencyProperty BusyAddFavProperty =
          DependencyProperty.Register("BusyAddFav", typeof(bool), typeof(TfsTaskPane));

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
                return _removeFavorite ?? (_removeFavorite = new DelegateCommand(null, p =>
                  {
                      OnRemoveFavoriteEvent(this, new WorkItemActionEventArgs(p as WorkItemInfo));
                  }));
            }
        }

        public ICommand CopyToClipboard
        {
            get
            {
                return _copyToClipboard ?? (_copyToClipboard = new DelegateCommand(null, p =>
                {
                    OnCopyToClipboardEvent(this, new WorkItemActionEventArgs(p as WorkItemInfo));
                }));
            }
        }
        #endregion

        #region constructor
        public TfsTaskPane()
        {
            InitializeComponent();
            BusyGetTasks = false;
            BusyAddFav = false;

            _onConnectToTfs = Observable.FromEventPattern(ConnectToTfs, "Click").Select(_ => Unit.Default);

            _onTaskFilterChanged = Observable.Merge(
                Observable.FromEventPattern<SelectionChangedEventArgs>(TfsProjects, "SelectionChanged").Select(e => _GetCurrentFilter()),
                Observable.FromEventPattern(RefreshButton, "Click").Select(e => _GetCurrentFilter()),
                Observable.FromEventPattern(UpdateTodaysTaskButton, "Click").Select(e => _GetCurrentFilterForToday()))
              .Where(f => f.Project != null)
              .Do(_ => SetBusyGetTasks(true));

            _onAddTasksToCalendar = Observable.Merge(
                Observable.FromEventPattern<MouseButtonEventArgs>(FavoriteTasks, "MouseDoubleClick")
                    .Select(e => ItemsControl.ContainerFromElement(e.Sender as ListBox, e.EventArgs.OriginalSource as DependencyObject) as ListBoxItem)
                    .Where(item => item != null)
                    .Select(item => Enumerable.Repeat(FavoriteTasks.ItemContainerGenerator.ItemFromContainer(item) as WorkItemInfo, 1)),
                Observable.FromEventPattern<MouseButtonEventArgs>(TfsTasks, "MouseDoubleClick")
                    .Select(e => ItemsControl.ContainerFromElement(e.Sender as ListBox, e.EventArgs.OriginalSource as DependencyObject) as ListBoxItem)
                    .Where(item => item != null)
                    .Select(item => Enumerable.Repeat(TfsTasks.ItemContainerGenerator.ItemFromContainer(item) as WorkItemInfo, 1)),
                Observable.FromEventPattern(AddAllWorkItemsToCalendar, "Click")
                    .Select(e => TfsTasks.ItemsSource as IEnumerable<WorkItemInfo>)
                );
            _onAddFavTask = Observable.Merge(
              Observable.FromEventPattern(AddFavTask, "Click").Select(_ => Unit.Default),
              Observable.FromEventPattern<KeyEventArgs>(NewTaskId, "PreviewKeyDown").Where(e => e.EventArgs.Key == Key.Enter).Select(_ => Unit.Default))
              .Where(e => NewTaskId.Text.Trim() != "" && NewTaskId.Text.Count() <= 12)
              .Select(e => Convert.ToInt64(NewTaskId.Text))
              .Do(_ => NewTaskId.Text = "");

            _onGoToReportClicked = Observable.FromEventPattern(GoToReportWebsite, "Click").Select(_ => Unit.Default);

            _onRemoveFavorite = Observable.FromEventPattern<WorkItemActionEventArgs>(this, "OnRemoveFavoriteEvent")
              .Where(e => e.EventArgs.Item != null)
              .Select(e => e.EventArgs.Item);

            _onCopyToClipboard = Observable.FromEventPattern<WorkItemActionEventArgs>(this, "OnCopyToClipboardEvent")
              .Where(e => e.EventArgs.Item != null)
              .Select(e => e.EventArgs.Item);

            Observable.FromEventPattern<DragEventArgs>(this, "DragOver")
              .Select(e =>
                {
                    if (e.EventArgs.Data.GetDataPresent(DataFormats.StringFormat))
                    {
                        string dataString = (string)e.EventArgs.Data.GetData(DataFormats.StringFormat);
                        e.EventArgs.Effects = DragDropEffects.Copy;
                    }
                    e.EventArgs.Handled = true;
                    return e;
                });

            _onDropItem = Observable.FromEventPattern<DragEventArgs>(this, "Drop")
              .Do(e =>
                {
                    Console.WriteLine(e.EventArgs.Data.ToString());
                })
              .Select(e => e.EventArgs.Data);
            #region Settings tab
            // Cannot capture outlook shutdown event. So save when things change
            // https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee720183(v=office.14)?redirectedfrom=MSDN#add-in-shutdown-changes-in-outlook-2010
            Observable.Merge(
            Observable.FromEventPattern<EventArgs>(ShowStartAndFinishDatesCheckBox, "Click"),
            Observable.FromEventPattern<EventArgs>(ShowOnlyWorkItemIdInCalendarCheckBox, "Click"),
            Observable.FromEventPattern<EventArgs>(UseStartAndFinishDatesofWorkItemToCreateCalendarEntriesCheckBox, "Click"))
                .Subscribe(e =>
            {
                Properties.Settings.Default.Save();
            });
            #endregion
        }
        #endregion

        #region ITfsTaskPaneView implementation
        IObservable<Unit> ITfsTaskPaneView.OnConnectToTfs()
        {
            return _onConnectToTfs;
        }
        IObservable<Unit> ITfsTaskPaneView.OnGoToReport()
        {
            return _onGoToReportClicked;
        }
        IObservable<WorkItemFilter> ITfsTaskPaneView.OnTaskFilterChanged()
        {
            return _onTaskFilterChanged;
        }
        IObservable<IEnumerable<WorkItemInfo>> ITfsTaskPaneView.OnAddTasksToCalendar()
        {
            return _onAddTasksToCalendar;
        }
        IObservable<long> ITfsTaskPaneView.OnAddFavTask()
        {
            return _onAddFavTask;
        }
        IObservable<IDataObject> ITfsTaskPaneView.OnDropItem()
        {
            return _onDropItem;
        }
        IObservable<WorkItemInfo> ITfsTaskPaneView.OnRemoveFavorite()
        {
            return _onRemoveFavorite;
        }
        IObservable<WorkItemInfo> ITfsTaskPaneView.OnCopyToClipboard()
        {
            return _onCopyToClipboard;
        }
        void ITfsTaskPaneView.SetProjectsList(IEnumerable<string> projects)
        {
            TfsProjects.ItemsSource = projects;
        }
        void ITfsTaskPaneView.SetTasksList(IEnumerable<WorkItemInfo> tasks)
        {
            TfsTasks.ItemsSource = tasks;
        }
        void ITfsTaskPaneView.SetFavTaskList(IEnumerable<WorkItemInfo> tasks)
        {
            FavoriteTasks.ItemsSource = tasks;
        }
        public void SetBusyGetTasks(bool busy)
        {
            BusyGetTasks = busy;
        }
        void ITfsTaskPaneView.SetBusyAddFav(bool busy)
        {
            BusyAddFav = busy;
            if (!busy) NewTaskId.Focus();
        }
        #endregion

        #region Private methods
        private WorkItemFilter _GetCurrentFilter()
        {
            return new WorkItemFilter()
            {
                Project = TfsProjects.SelectedItem != null ? TfsProjects.SelectedItem.ToString() : null,
                ShowTasks = ShowTasks.IsChecked.GetValueOrDefault(false),
                ShowBugs = ShowBugs.IsChecked.GetValueOrDefault(false),
                ShowProposed = ShowProposed.IsChecked.GetValueOrDefault(false),
                ShowActive = ShowActive.IsChecked.GetValueOrDefault(false),
                ShowResolved = ShowResolved.IsChecked.GetValueOrDefault(false),
                ShowClosed = ShowClosed.IsChecked.GetValueOrDefault(false),
                ShowStartAndFinishDates = ShowStartAndFinishDatesCheckBox.IsChecked.Value
            };
        }
        private WorkItemFilter _GetCurrentFilterForToday()
        {
            WorkItemFilter filter = _GetCurrentFilter();
            filter.StartDate = DateTime.Today;
            filter.FinishDate = DateTime.Today.AddDays(1).AddSeconds(-1);
            return filter;
        }
        #endregion
    }
}
