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

    public IEnumerable<WorkItemInfo> GetTasks(string s)
      {
      TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(new Uri(Properties.Settings.Default.TfsUri));
      WorkItemStore workItemStore = (WorkItemStore)tpc.GetService(typeof(WorkItemStore));
      WorkItemCollection queryResults = workItemStore.Query(
        @"Select [Id], [Title], [Completed Work] From WorkItems " +
        @"Where ([Work Item Type] = 'Task' or [Work Item Type] = 'Bug') " +
        @"And [System.TeamProject] = '" + s + "' And [System.AssignedTo] = @Me " +
        @"And ([State] = 'Active' Or [State] = 'Resolved') " +
        @"Order By [Work Item Type]");

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
    }
  }
