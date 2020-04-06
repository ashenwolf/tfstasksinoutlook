using System;
using System.Linq;
using System.Diagnostics;
using System.Reactive.Concurrency;
using System.Reactive.Linq;
using System.Collections.ObjectModel;
using System.Windows;
using TFSTasksInOutlook.Calendar;

namespace TFSTasksInOutlook
{
    class TfsTaskPaneController
    {
        private readonly TfsProxy _tfsProxy = new TfsProxy();
        private readonly Dataset.TfsTasksStorage _dataset = Dataset.TfsTasksStorage.Load();

        private readonly ITfsTaskPaneView _paneView;
        private readonly ObservableCollection<WorkItemInfo> _favoriteWorkItems;

        public TfsTaskPaneController(ITfsTaskPaneView pane)
        {
            _paneView = pane;
            _favoriteWorkItems = new ObservableCollection<WorkItemInfo>();
            _paneView.SetProjectsList(_dataset.TfsProjects);
            _paneView.SetFavTaskList(_favoriteWorkItems);

            _SubscribeObservables();
            _LoadFavorites();
        }

        private void _LoadFavorites()
        {
            _favoriteWorkItems.Clear();
            _dataset.FavoriteWorkItems
              .ToObservable()
              .Subscribe(wi =>
                {
                    _favoriteWorkItems.Add(wi);
                    _paneView.SetBusyAddFav(false);
                });
        }

        private void _SubscribeObservables()
        {
            _paneView.OnConnectToTfs().Subscribe(_ => _SelectNewTfsServer());

            _paneView.OnGoToReport().Subscribe(_ => _OpenReportsPageInBrowser());

            _paneView.OnTaskFilterChanged()
              .ObserveOn(Scheduler.Default)
              .Select(s => _tfsProxy.GetTasks(_dataset.TfsUri, s))
              .ObserveOn(DispatcherScheduler.Current)
              .Subscribe(r =>
                {
                    _paneView.SetTasksList(r);
                    _paneView.SetBusyGetTasks(false);
                });

            _paneView.OnTaskDoubleClicked().Subscribe(CalendarManager.CreateItemInCalendar);

            _paneView.OnAddFavTask()
              .Where(id => _favoriteWorkItems.All(wi => wi.Id != Convert.ToInt64(id)))
              .Do(_ => _paneView.SetBusyAddFav(true))
              .ObserveOn(Scheduler.Default)
              .Select(_GetTaskInfo)
              .ObserveOn(DispatcherScheduler.Current)
              .Do(_ => _paneView.SetBusyAddFav(false))
              .Where(r => r != null)
              .Subscribe(r =>
                {
                    _favoriteWorkItems.Add(r);
                    _SaveFavoriteItems();
                });

            _paneView.OnRemoveFavorite()
              .Subscribe(item =>
                {
                    _favoriteWorkItems.Remove(item);
                    _SaveFavoriteItems();
                });

            _paneView.OnCopyToClipboard()
              .Subscribe(_CopyToClipboard);
        }

        private void _CopyToClipboard(WorkItemInfo item)
        {
            Clipboard.SetText(item.ItemType + " #" + item.Id + ": " + item.Title);
        }

        private WorkItemInfo _GetTaskInfo(long id)
        {
            return _tfsProxy.GetTaskInfo(_dataset.TfsUri, id);
        }

        private void _OpenReportsPageInBrowser()
        {
            Process.Start(Properties.Settings.Default.TimesheetReportUrl);
        }

        private void _SelectNewTfsServer()
        {
            var tfsServer = _tfsProxy.GetNewTfsServer();
            if (tfsServer != null)
            {
                _dataset.TfsUri = tfsServer.TfsUri;
                _SaveProjectsList(tfsServer.TfsProjects);
                _LoadFavorites();
            }
        }
        private void _SaveFavoriteItems()
        {
            _dataset.FavoriteWorkItems = _favoriteWorkItems.ToList();
            _dataset.Save();
        }

        private void _SaveProjectsList(string[] projects)
        {
            _dataset.TfsProjects = projects.OrderBy(s => s).ToList();
            _paneView.SetProjectsList(_dataset.TfsProjects);
            _dataset.Save();
        }
    }
}
