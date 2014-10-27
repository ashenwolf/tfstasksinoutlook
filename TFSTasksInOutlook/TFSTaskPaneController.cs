using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reactive;
using System.Reactive.Concurrency;
using System.Reactive.Linq;
using System.Text;

namespace TFSTasksInOutlook
  {
  public class WorkItemInfo
    {
    public long   Id {get; set;}
    public string Title { get; set; }
    public double CompletedWork { get; set; }
    public string ItemType { get; set; }
    public string Project { get; set; }
    }

  public class WorkItemFilter
    {
    public string Project { get; set; }
    public bool ShowTasks { get; set; }
    public bool ShowBugs { get; set; }
    public bool ShowProposed { get; set; }
    public bool ShowActive { get; set; }
    public bool ShowResolved { get; set; }
    public bool ShowClosed { get; set; }
    }

  class TFSTaskPaneController
    {
    private TFSProxy tfsProxy = new TFSProxy();
    private ITFSTaskPaneView paneView;
    private List<WorkItemInfo> favouriteWorkItems;

    public TFSTaskPaneController(ITFSTaskPaneView pane)
      {
      paneView = pane;
      favouriteWorkItems = new List<WorkItemInfo>();
      paneView.SetFavTaskList(favouriteWorkItems);

      _SubscribeObservables();
      _LoadFavourites();
      }

    private void _LoadFavourites()
      {
      var ids = new List<string>[] { new List<string>() };
      if (Properties.Settings.Default.FavouriteWorkItems != null)
        foreach (var id in Properties.Settings.Default.FavouriteWorkItems)
          ids[0].Add(id);

      ids.ToObservable()
        .Do(_ => paneView.SetBusyAddFav(true))
        .ObserveOn(Scheduler.Default)
        .Select(x => tfsProxy.GetTasksByIds(x))
        .ObserveOn(DispatcherScheduler.Current)
        .Subscribe(wi =>
          {
            favouriteWorkItems.AddRange(wi);
            paneView.SetBusyAddFav(false);
          });
      }

    private void _SubscribeObservables()
      {
      paneView.OnConnectToTfs().Subscribe(_ => _SelectNewTfsServer());

      paneView.OnGoToReport().Subscribe(_ => _OpenReportsPageInBrowser());

      paneView.OnTaskFilterChanged()
        .ObserveOn(Scheduler.Default)
        .Select(s => tfsProxy.GetTasks(s))
        .ObserveOn(DispatcherScheduler.Current)
        .Subscribe(r =>
          {
          paneView.SetTasksList(r);
          paneView.SetBusyGetTasks(false);
          });

      paneView.OnTaskDoubleClicked().Subscribe(t => _CreateItemInCalendar(t));

      paneView.OnAddFavTask()
        .Where(id => !favouriteWorkItems.Any(wi => wi.Id == Convert.ToInt64(id)))
        .Do(_ => paneView.SetBusyAddFav(true))
        .ObserveOn(Scheduler.Default)
        .Select(id => _GetTaskInfo(id))
        .ObserveOn(DispatcherScheduler.Current)
        .Do(_ => paneView.SetBusyAddFav(false))
        .Where(r => r != null)
        .Subscribe(r =>
          {
          favouriteWorkItems.Add(r);
          var coll = new System.Collections.Specialized.StringCollection();
          coll.AddRange(favouriteWorkItems.Select(wi => wi.Id.ToString()).ToArray());
          Properties.Settings.Default.FavouriteWorkItems = coll;
          Properties.Settings.Default.Save();
          });
      }

    private WorkItemInfo _GetTaskInfo(long id)
      {
      return tfsProxy.GetTaskInfo(id);
      }

    private void _OpenReportsPageInBrowser()
      {
      Process.Start(Properties.Settings.Default.TimesheetReportUrl);
      }

    private void _SelectNewTfsServer()
      {
      var tfsServer = tfsProxy.GetNewTfsServer();
      if (tfsServer != null)
        {
        Properties.Settings.Default.TfsUri = tfsServer.TfsUri;
        var coll = new System.Collections.Specialized.StringCollection();
        coll.AddRange(tfsServer.TfsProjects);
        Properties.Settings.Default.TfsProjects = coll;
        Properties.Settings.Default.Save();
        }
      }

    private void _CreateItemInCalendar(WorkItemInfo item)
      {
      var minDate = DateTime.Today;
      minDate -= TimeSpan.FromDays(365);

      var expl = Globals.ThisAddIn.Application.ActiveExplorer();
      var view = expl.CurrentView as Microsoft.Office.Interop.Outlook.View;
      if (view.ViewType == Microsoft.Office.Interop.Outlook.OlViewType.olCalendarView)
        {
        var calView = view as Microsoft.Office.Interop.Outlook.CalendarView;
        var dstart = calView.SelectedStartTime;
        var dend = calView.SelectedEndTime;

        if (dstart > minDate && dend > minDate)
          {
          var folder = expl.CurrentFolder as Microsoft.Office.Interop.Outlook.Folder;
          var appointment = folder.Items.Add("IPM.Appointment") as Microsoft.Office.Interop.Outlook.AppointmentItem;
          appointment.Subject = "#" + item.Id + " " + item.Title;
          appointment.Start = dstart;
          appointment.End = dend;
          appointment.Save();
          }
        }
      }
    }
  }
