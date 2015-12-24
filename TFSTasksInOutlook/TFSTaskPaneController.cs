using System;
using System.Linq;
using System.Diagnostics;
using System.Reactive.Concurrency;
using System.Reactive.Linq;
using System.Collections.ObjectModel;
using System.Windows;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;

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

      _paneView.OnTaskDoubleClicked().Subscribe(_CreateItemInCalendar);

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

    private void _CreateNewItemsInCalendar(WorkItemInfo item, View view, Explorer explorer)
      {
      var minDate = DateTime.Today - TimeSpan.FromDays(365);
      var calView = view as CalendarView;
      if (calView == null) return;
      var dstart = calView.SelectedStartTime;
      var dend = calView.SelectedEndTime;

      if (dstart <= minDate || dend <= minDate)
        return;
      
      var folder = explorer.CurrentFolder as Folder;
      if (folder == null) return;

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
      var restrictCriteria = "[Start] <= '" + dend.ToString("g") + "'" +
                             " AND [End] >= '" + dstart.ToString("g") + "'";


      // Filter items according to the criteria and add fake item to the end 
      // as an artifitial upper boundary
      var items = folder.Items;
      items.IncludeRecurrences = true;
      items.Sort("[Start]", Type.Missing);
      items = items.Restrict(restrictCriteria);

      var appointments = items
        .Cast<AppointmentItem>()
        .Where(i => !i.AllDayEvent)
        .Select(i => new { i.Start, i.End })
        .ToList();
      appointments.Add(new { Start = dend, End = dend.AddHours(1) });

      appointments.ToObservable()
        .Where(i => i.End > dstart)
        .Subscribe(i =>
          {
          if (dstart < i.Start)
            _AddAppointment(folder.Items,
              _WorkItemInfoToText(item),
              dstart, i.Start);

          dstart = i.End;
          });
      }

    private void _UpdateItemsInCalendar(WorkItemInfo item, Explorer explorer)
      {
      foreach (var selected in explorer.Selection)
        {
        if (!(selected is AppointmentItem)) continue;
        var appointment = selected as AppointmentItem;
        var rgx = new Regex(@"^(#\d+)");
        appointment.Subject =
          rgx.IsMatch(appointment.Subject.Trim())
            ? rgx.Replace(appointment.Subject, "#" + item.Id, 1)
            : appointment.Subject.Trim().Length > 0
              ? string.Format("#{0} {1}", item.Id, appointment.Subject)
              : _WorkItemInfoToText(item);
        appointment.Save();
        }
      }

    private void _CreateItemInCalendar(WorkItemInfo item)
      {
      var expl = Globals.ThisAddIn.Application.ActiveExplorer();
      var view = expl.CurrentView as View;
      if (view == null || view.ViewType != OlViewType.olCalendarView) return;
      if (expl.Selection.Count > 0)
        {
        _UpdateItemsInCalendar(item, expl);
        }
      else
        {
        _CreateNewItemsInCalendar(item, view, expl);
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

    private void _AddAppointment(Items items, string subject, DateTime start, DateTime end)
      {
      var appointment = items.Add("IPM.Appointment") as AppointmentItem;
      if (appointment == null) return;
      appointment.Subject = subject;
      appointment.Start = start;
      appointment.End = end;
      appointment.Save();
      }

    private string _WorkItemInfoToText(WorkItemInfo wi)
      {
      return "#" + wi.Id + " " + wi.Title;
      }
    }
  }
