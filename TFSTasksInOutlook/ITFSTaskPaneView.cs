using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive;
using System.Text;

namespace TFSTasksInOutlook
  {
  interface ITFSTaskPaneView
    {
    IObservable<Unit> OnConnectToTfs();
    IObservable<Unit> OnGoToReport();
    IObservable<WorkItemFilter> OnTaskFilterChanged();
    IObservable<WorkItemInfo> OnTaskDoubleClicked();

    void SetTasksList(IEnumerable<WorkItemInfo> tasks);
    void SetBusy(bool busy);
    }
  }
