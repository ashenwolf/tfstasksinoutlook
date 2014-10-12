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
      TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(new Uri(Properties.Settings.Default.TfsUri));
      WorkItemStore workItemStore = (WorkItemStore)tpc.GetService(typeof(WorkItemStore));
      var q =
        @"Select [Id], [Title], [Completed Work] From WorkItems " +
        @"Where " +
        @"[System.TeamProject] = '" + s.Project + "' And [System.AssignedTo] = @Me " +
        GetItemTypeFilter(s) +
        _GetStateFilter(s) +
        @"Order By [Work Item Type] ";
      WorkItemCollection queryResults = workItemStore.Query(q);

      var tasks = new List<WorkItemInfo>();
      foreach (WorkItem wi in queryResults)
        {
        tasks.Add(new WorkItemInfo()
        {
          Id = wi.Id,
          Title = wi.Title,
          CompletedWork = wi.Fields["Completed Work"].Value != null ? Convert.ToDouble(wi.Fields["Completed Work"].Value) : 0.0,
          ItemType = wi.Type.Name
        });
        }
      return tasks;
      }

    private string GetItemTypeFilter(WorkItemFilter s)
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
    }
  }
