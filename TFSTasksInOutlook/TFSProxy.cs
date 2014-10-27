using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace TFSTasksInOutlook
  {
  class TFSProxy
    {
    public class TFSServerData
      {
      public TFSServerData(string tfsUri, string[] tfsProjects)
        {
        TfsUri = tfsUri;
        TfsProjects = tfsProjects;
        }

      public string TfsUri { get; private set; }
      public string[] TfsProjects { get; private set; }
      }

    public TFSServerData GetNewTfsServer()
      {
      using (TeamProjectPicker tpp = new TeamProjectPicker(TeamProjectPickerMode.MultiProject, false))
        {
        DialogResult result = tpp.ShowDialog();
        if (result == DialogResult.OK)
          {
          System.Console.WriteLine("Selected Team Project Collection Uri: " + tpp.SelectedTeamProjectCollection.Uri);
          return new TFSServerData(
            tpp.SelectedTeamProjectCollection.Uri.ToString(),
            tpp.SelectedProjects.Select(p => p.Name).ToArray());
          }
        }
      return null;
      }

    public IEnumerable<WorkItemInfo> GetTasks(WorkItemFilter s)
      {
      var q =
        @"Select [Id], [Title], [Completed Work], [System.TeamProject] From WorkItems " +
        @"Where " +
        @"[System.TeamProject] = '" + s.Project + "' And [System.AssignedTo] = @Me " +
        _GetItemTypeFilter(s) +
        _GetStateFilter(s) +
        @"Order By [Work Item Type] ";
      return _QueryAll(q);
      }

    public WorkItemInfo GetTaskInfo(long id)
      {
      var q = @"Select [Id], [Title], [Completed Work], [System.TeamProject] From WorkItems Where [Id] = " + id.ToString();
      return _QueryOne(q);
      }

    public IEnumerable<WorkItemInfo> GetTasksByIds(IEnumerable<string> ids)
      {
      var q = @"Select [Id], [Title], [Completed Work], [System.TeamProject] From WorkItems Where [Id] IN (" + String.Join(", ", ids) +") ";
      return _QueryAll(q);
      }

    private string _GetItemTypeFilter(WorkItemFilter s)
      {
      var filters = new List<string>();
      if (s.ShowTasks) filters.Add("[Work Item Type] = 'Task'");
      if (s.ShowBugs) filters.Add("[Work Item Type] = 'Bug'");

      return filters.Count > 0
        ? @"And (" + String.Join(" Or ", filters) + ") "
        : "And ([Work Item Type] = 'Task' Or [Work Item Type] = 'Bug') ";
      }

    private string _GetStateFilter(WorkItemFilter s)
      {
      var filters = new List<string>();
      if (s.ShowProposed) filters.Add("[State] = 'Proposed'");
      if (s.ShowActive)   filters.Add("[State] = 'Active'");
      if (s.ShowResolved) filters.Add("[State] = 'Resolved'");
      if (s.ShowClosed)   filters.Add("[State] = 'Closed'");

      return  filters.Count > 0
        ? @"And (" + String.Join(" Or ", filters) + ") "
        : "And [State] = 'Active' ";
      }

    private WorkItemInfo _WorkItemToWorkItemInfo(WorkItem wi)
      {
      return new WorkItemInfo()
            {
              Id = wi.Id,
              Title = wi.Title,
              CompletedWork = wi.Fields["Completed Work"].Value != null ? Convert.ToDouble(wi.Fields["Completed Work"].Value) : 0.0,
              ItemType = wi.Type.Name,
              Project = wi.Project.Name
            };
      }

    private IEnumerable<WorkItemInfo> _QueryAll(string wiql)
      {
      var tasks = new List<WorkItemInfo>();
      try
        {
        TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(new Uri(Properties.Settings.Default.TfsUri));
        WorkItemStore workItemStore = (WorkItemStore)tpc.GetService(typeof(WorkItemStore));
        WorkItemCollection queryResults = workItemStore.Query(wiql);
        foreach (WorkItem wi in queryResults)
          {
          if (wi.Type.Name.Equals("Bug", StringComparison.InvariantCultureIgnoreCase) || wi.Type.Name.Equals("Task", StringComparison.InvariantCultureIgnoreCase))
            tasks.Add(_WorkItemToWorkItemInfo(wi));
          }
        }
      catch (Microsoft.TeamFoundation.TeamFoundationServiceUnavailableException) { }
      catch (Microsoft.TeamFoundation.WorkItemTracking.Client.UnexpectedErrorException) { }
      return tasks;
      }

    private WorkItemInfo _QueryOne(string wiql)
      {
      return _QueryAll(wiql).FirstOrDefault();
      }
    }
  }
