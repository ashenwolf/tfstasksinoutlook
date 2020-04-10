using System;
using System.Collections.Generic;
using System.Reactive;
using System.Windows;

namespace TFSTasksInOutlook
{
    internal interface ITfsTaskPaneView
    {
        IObservable<Unit> OnConnectToTfs();
        IObservable<Unit> OnGoToReport();
        IObservable<WorkItemFilter> OnTaskFilterChanged();
        IObservable<IEnumerable<WorkItemInfo>> OnAddTasksToCalendar();
        IObservable<long> OnAddFavTask();
        IObservable<WorkItemInfo> OnRemoveFavorite();
        IObservable<WorkItemInfo> OnCopyToClipboard();
        IObservable<IDataObject> OnDropItem();

        void SetProjectsList(IEnumerable<string> projects);
        void SetTasksList(IEnumerable<WorkItemInfo> tasks);
        void SetFavTaskList(IEnumerable<WorkItemInfo> tasks);
        void SetBusyGetTasks(bool busy);
        void SetBusyAddFav(bool busy);
    }
}
