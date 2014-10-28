﻿using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reactive;
using System.Reactive.Concurrency;
using System.Reactive.Linq;
using System.Text;
using System.Collections.ObjectModel;
using System.Windows;

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
    private ObservableCollection<WorkItemInfo> favoriteWorkItems;

    public TFSTaskPaneController(ITFSTaskPaneView pane)
      {
      paneView = pane;
      favoriteWorkItems = new ObservableCollection<WorkItemInfo>();
      paneView.SetFavTaskList(favoriteWorkItems);

      _SubscribeObservables();
      _LoadFavorites();
      }

    private void _LoadFavorites()
      {
      var ids = new List<string>[] { new List<string>() };
      if (Properties.Settings.Default.FavoriteWorkItems != null && Properties.Settings.Default.FavoriteWorkItems.Count > 0)
        {
        foreach (var id in Properties.Settings.Default.FavoriteWorkItems)
          ids[0].Add(id);

        favoriteWorkItems.Clear();
        ids.ToObservable()
          .Do(_ => paneView.SetBusyAddFav(true))
          .ObserveOn(Scheduler.Default)
          .Select(x => tfsProxy.GetTasksByIds(x))
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
        .Select(s => tfsProxy.GetTasks(s))
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
        _SaveProjectsList(tfsServer.TfsProjects);
        _LoadFavorites();
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

    private void _SaveFavoriteItems()
      {
      var coll = new System.Collections.Specialized.StringCollection();
      coll.AddRange(favoriteWorkItems.Select(wi => wi.Id.ToString()).ToArray());
      Properties.Settings.Default.FavoriteWorkItems = coll;
      Properties.Settings.Default.Save();
      }

    private void _SaveProjectsList(string[] projects)
      {
      var coll = new System.Collections.Specialized.StringCollection();
      coll.AddRange(projects.OrderBy(s => s).ToArray());
      Properties.Settings.Default.TfsProjects = coll;
      Properties.Settings.Default.Save();
      }
    }
  }
