using System;
using System.Collections.Generic;
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
    private int backgroundOperations;

    public TFSTaskPaneController(ITFSTaskPaneView pane)
      {
      paneView = pane;
      backgroundOperations = 0;

      _SubscribeObservables();
      }

    private void _SubscribeObservables()
      {
      paneView.OnConnectToTfs().Subscribe(_ => SelectNewTfsServer());

      paneView.OnProjectSelected()
        .ObserveOn(Scheduler.Default)
        .Select(s => tfsProxy.GetTasks(s))
        .ObserveOn(DispatcherScheduler.Current)
        .Subscribe(r =>
          {
            paneView.SetTasksList(r);
            paneView.SetBusy(false);
          });

      paneView.OnTaskDoubleClicked().Subscribe(t => _CreateItemInCalendar(t));
      }

    public void SelectNewTfsServer()
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
