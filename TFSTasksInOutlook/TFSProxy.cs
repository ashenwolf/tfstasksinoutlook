using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using TFSTasksInOutlook.TFS;

namespace TFSTasksInOutlook
{
    class TfsProxy
    {
        public class TfsServerData
        {
            public TfsServerData(string tfsUri, string[] tfsProjects)
            {
                TfsUri = tfsUri;
                TfsProjects = tfsProjects;
            }

            public string TfsUri { get; private set; }
            public string[] TfsProjects { get; private set; }
        }

        public TfsServerData GetNewTfsServer()
        {
            using (TeamProjectPicker tpp = new TeamProjectPicker(TeamProjectPickerMode.MultiProject, false))
            {
                DialogResult result = tpp.ShowDialog();
                if (result == DialogResult.OK)
                {
                    System.Console.WriteLine("Selected Team Project Collection Uri: " + tpp.SelectedTeamProjectCollection.Uri);
                    return new TfsServerData(
                      tpp.SelectedTeamProjectCollection.Uri.ToString(),
                      tpp.SelectedProjects.Select(p => p.Name).ToArray());
                }
            }
            return null;
        }

        public IEnumerable<WorkItemInfo> GetTasks(string tfsUri, WorkItemFilter s)
        {
            var q = WorkItemFilterQueryBuilder.Build(s);

            return _QueryAll(tfsUri, q,s.ShowStartAndFinishDates);
        }

        public WorkItemInfo GetTaskInfo(string tfsUri, long id)
        {
            var q = @"Select [Id], [Title], [Completed Work], [System.TeamProject] From WorkItems Where [Id] = " + id.ToString();
            return _QueryOne(tfsUri, q);
        }

        public IEnumerable<WorkItemInfo> GetTasksByIds(string tfsUri, IEnumerable<string> ids)
        {
            var q = @"Select [Id], [Title], [Completed Work], [System.TeamProject] From WorkItems Where [Id] IN (" + String.Join(", ", ids) + ") ";
            return _QueryAll(tfsUri, q,false);
        }

        private WorkItemInfo _WorkItemToWorkItemInfo(WorkItem wi,bool includeStartAndFinishDates)
        {
            // Even if the Start and Finish dates are not in the SELECT part of query, its fetching those values. So control based on parameter.
            return new WorkItemInfo()
            {
                Id = wi.Id,
                Title = wi.Title,
                CompletedWork = wi.Fields["Completed Work"].Value != null ? Convert.ToDouble(wi.Fields["Completed Work"].Value) : 0.0,
                ItemType = wi.Type.Name,
                Project = wi.Project.Name,
                StartDate = includeStartAndFinishDates? (wi.Fields["Start Date"].Value == null ? null : Convert.ToDateTime(wi.Fields["Start Date"].Value) as DateTime? ):null,
                FinishDate = includeStartAndFinishDates ? (wi.Fields["Finish Date"].Value == null ? null : Convert.ToDateTime(wi.Fields["Finish Date"].Value) as DateTime?):null
            };
        }

        private IEnumerable<WorkItemInfo> _QueryAll(string tfsUri, string wiql, bool includeStartAndFinishDates)
        {
            var tasks = new List<WorkItemInfo>();
            try
            {
                TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(new Uri(tfsUri));
                WorkItemStore workItemStore = (WorkItemStore)tpc.GetService(typeof(WorkItemStore));
                WorkItemCollection queryResults = _ExecuteQuery(workItemStore, wiql);
                foreach (WorkItem wi in queryResults)
                {
                    if (wi.Type.Name.Equals("Bug", StringComparison.InvariantCultureIgnoreCase) || wi.Type.Name.Equals("Task", StringComparison.InvariantCultureIgnoreCase))
                        tasks.Add(_WorkItemToWorkItemInfo(wi, includeStartAndFinishDates));
                }
            }
            catch (Microsoft.TeamFoundation.TeamFoundationServiceUnavailableException) { }
            catch (Microsoft.TeamFoundation.WorkItemTracking.Client.UnexpectedErrorException) { }
            catch (System.UriFormatException) { }
            return tasks;
        }

        private static WorkItemCollection _ExecuteQuery(WorkItemStore workItemStore, string wiql)
        {
            WorkItemCollection queryResults = null;
            if (Regex.IsMatch(wiql, "date", RegexOptions.IgnoreCase))
            {
                // To include date and time precision - http://teamfoundation.blogspot.com/2008/01/specifying-date-and-time-in-wiql.html
                Query query = new Query(workItemStore, wiql, null, false);
                ICancelableAsyncResult cancelableAsyncResult = query.BeginQuery();
                queryResults = query.EndQuery(cancelableAsyncResult);
            }
            else
            {
                // Fallback to existing execution.
                queryResults = workItemStore.Query(wiql);
            }
            return queryResults;
        }

        private WorkItemInfo _QueryOne(string tfsUri, string wiql)
        {
            return _QueryAll(tfsUri, wiql,false).FirstOrDefault();
        }
    }
}
