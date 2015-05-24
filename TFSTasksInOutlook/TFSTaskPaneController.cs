using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reactive;
using System.Reactive.Concurrency;
using System.Reactive.Linq;
using System.Text;
using System.Collections.ObjectModel;
using System.Windows;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;

namespace TFSTasksInOutlook
  {
  public class WorkItemInfo
    {
    public long Id { get; set; }
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
    private Dataset.TFSTasksStorage dataset = new Dataset.TFSTasksStorage();

    private ITFSTaskPaneView paneView;
    private ObservableCollection<WorkItemInfo> favoriteWorkItems;

    public TFSTaskPaneController(ITFSTaskPaneView pane)
      {
      dataset.Load();
      paneView = pane;
      favoriteWorkItems = new ObservableCollection<WorkItemInfo>();
      paneView.SetProjectsList(dataset.TfsProjects);
      paneView.SetFavTaskList(favoriteWorkItems);

      _SubscribeObservables();
      _LoadFavorites();
      }

    private void _LoadFavorites()
      {
      var ids = new List<string>[] { new List<string>() };
      if (dataset.FavoriteWorkItems != null && dataset.FavoriteWorkItems.Count > 0)
        {
        foreach (var id in dataset.FavoriteWorkItems)
          ids[0].Add(id);

        favoriteWorkItems.Clear();
        ids.ToObservable()
          .Do(_ => paneView.SetBusyAddFav(true))
          .ObserveOn(Scheduler.Default)
          .Select(x => tfsProxy.GetTasksByIds(dataset.TfsUri, x))
          .ObserveOn(DispatcherScheduler.Current)
          .Subscribe(wi =>
            {
              wi.ToList().ForEach(item => favoriteWorkItems.Add(item));
              paneView.SetBusyAddFav(false);
            });
        }
      }

    private void _SubscribeObservables()
      {
      paneView.OnConnectToTfs().Subscribe(_ => _SelectNewTfsServer());

      paneView.OnGoToReport().Subscribe(_ => _OpenReportsPageInBrowser());

      paneView.OnTaskFilterChanged()
        .ObserveOn(Scheduler.Default)
        .Select(s => tfsProxy.GetTasks(dataset.TfsUri, s))
        .ObserveOn(DispatcherScheduler.Current)
        .Subscribe(r =>
          {
            paneView.SetTasksList(r);
            paneView.SetBusyGetTasks(false);
          });

      paneView.OnTaskDoubleClicked().Subscribe(t => _CreateItemInCalendar(t));

      paneView.OnAddFavTask()
        .Where(id => !favoriteWorkItems.Any(wi => wi.Id == Convert.ToInt64(id)))
        .Do(_ => paneView.SetBusyAddFav(true))
        .ObserveOn(Scheduler.Default)
        .Select(id => _GetTaskInfo(id))
        .ObserveOn(DispatcherScheduler.Current)
        .Do(_ => paneView.SetBusyAddFav(false))
        .Where(r => r != null)
        .Subscribe(r =>
          {
            favoriteWorkItems.Add(r);
            _SaveFavoriteItems();
          });

      paneView.OnRemoveFavorite()
        .Subscribe(item =>
          {
            favoriteWorkItems.Remove(item);
            _SaveFavoriteItems();
          });
      }

    private WorkItemInfo _GetTaskInfo(long id)
      {
      return tfsProxy.GetTaskInfo(dataset.TfsUri, id);
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
        dataset.TfsUri = tfsServer.TfsUri;
        _SaveProjectsList(tfsServer.TfsProjects);
        _LoadFavorites();
        }
      }

    private void _CreateNewItemsInCalendar(WorkItemInfo item, View view, Explorer explorer)
      {
      var minDate = DateTime.Today - TimeSpan.FromDays(365);
      var calView = view as Microsoft.Office.Interop.Outlook.CalendarView;
      var dstart = calView.SelectedStartTime;
      var dend = calView.SelectedEndTime;

      if (dstart <= minDate || dend <= minDate)
        return;
      
      var folder = explorer.CurrentFolder as Microsoft.Office.Interop.Outlook.Folder;

      // Check for multiday event
      // Add 8-hour event for each day 9:00 - 17:00
      if (dend - dstart >= TimeSpan.FromDays(1))
        {
        Observable.Generate(
          dstart + TimeSpan.FromHours(9),
          date => date <= dend,
          date => date + TimeSpan.FromDays(1),
          date => new { Start = date, End = date + TimeSpan.FromHours(8) })
          .Subscribe(i =>
            _AddAppointment(folder.Items,
              "#" + item.Id + " " + item.Title,
              i.Start, i.End));

        return;
        }

      // Handle single day event differently - check for existing appointments
      // Criteria for any appointment, overlapping with [Start, End] interval
      string restrictCriteria = "[Start] <= '" + dend.ToString("g") + "'" +
                      " AND [End] >= '" + dstart.ToString("g") + "'";


      // Filter items according to the criteria and add fake item to the end 
      // as an artifitial upper boundary
      var items = folder.Items;
      items.IncludeRecurrences = true;
      items.Sort("[Start]", Type.Missing);
      items = items.Restrict(restrictCriteria);

      var appointments = items
        .Cast<Microsoft.Office.Interop.Outlook.AppointmentItem>()
        .Where(i => !i.AllDayEvent)
        .Select(i => new { Start = i.Start, End = i.End })
        .ToList();
      appointments.Add(new { Start = dend, End = dend.AddHours(1) });

      Observable.ToObservable(appointments)
        .Where(i => i.End > dstart)
        .Subscribe(i =>
        {
          if (dstart < i.Start)
            _AddAppointment(folder.Items,
              "#" + item.Id + " " + item.Title,
              dstart, i.Start);

          dstart = i.End;
        });
      }

    private void _UpdateItemsInCalendar(WorkItemInfo item, View view, Explorer explorer)
      {
      foreach (var selected in explorer.Selection)
        {
        if (selected is AppointmentItem)
          {
          var appointment = selected as AppointmentItem;
          Regex rgx = new Regex(@"^(#\d+)");
          if (rgx.IsMatch(appointment.Subject.Trim()))
            appointment.Subject = rgx.Replace(appointment.Subject, "#" + item.Id, 1);
          else
            appointment.Subject = "#" + item.Id + " " + appointment.Subject;
          appointment.Save();
          }
        }
      }

    private void _CreateItemInCalendar(WorkItemInfo item)
      {
      var expl = Globals.ThisAddIn.Application.ActiveExplorer();
      var view = expl.CurrentView as View;
      if (view.ViewType == OlViewType.olCalendarView)
        {
        if (expl.Selection.Count > 0)
          {
          _UpdateItemsInCalendar(item, view, expl);
          }
        else
          {
          _CreateNewItemsInCalendar(item, view, expl);
          }
        }
      }

    private void _SaveFavoriteItems()
      {
      dataset.FavoriteWorkItems = favoriteWorkItems.Select(wi => wi.Id.ToString()).ToList();
      dataset.Save();
      }

    private void _SaveProjectsList(string[] projects)
      {
      dataset.TfsProjects = projects.OrderBy(s => s).ToList();
      paneView.SetProjectsList(dataset.TfsProjects);
      dataset.Save();
      }

    private void _AddAppointment(Items items, string subject, DateTime start, DateTime end)
      {
      var appointment = items.Add("IPM.Appointment") as Microsoft.Office.Interop.Outlook.AppointmentItem;
      appointment.Subject = subject;
      appointment.Start = start;
      appointment.End = end;
      appointment.Save();
      }
    }
  }
