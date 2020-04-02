using System;
using System.Collections.Generic;

namespace TFSTasksInOutlook.TFS
{
    public class WorkItemFilterQueryBuilder
    {
        public static string Build(WorkItemFilter filter)
        {
            return @"Select [Id], [Title], [Completed Work], [System.TeamProject] From WorkItems " +
         @"Where " +
         @"[System.TeamProject] = '" + filter.Project + "' And [System.AssignedTo] = @Me " +
         WorkItemFilterQueryBuilder._GetItemTypeFilter(filter) +
         WorkItemFilterQueryBuilder._GetStateFilter(filter) +
         WorkItemFilterQueryBuilder._GetStartAndFinishDateFilter(filter) +
         @"Order By [Work Item Type] ";
        }

        private static string _GetStartAndFinishDateFilter(WorkItemFilter filter)
        {
            if (filter.FinishDate == null || filter.StartDate == null) return string.Empty;
            else return $"And [Start Date] >= '{filter.StartDate:MM/dd/yyyy HH:mm:ssZ}' And [Finish Date] <= '{filter.FinishDate:MM/dd/yyyy HH:mm:ssZ}' ";
        }

        private static string _GetItemTypeFilter(WorkItemFilter s)
        {
            var filters = new List<string>();
            if (s.ShowTasks) filters.Add("[Work Item Type] = 'Task'");
            if (s.ShowBugs) filters.Add("[Work Item Type] = 'Bug'");

            return filters.Count > 0
              ? @"And (" + String.Join(" Or ", filters) + ") "
              : "And ([Work Item Type] = 'Task' Or [Work Item Type] = 'Bug') ";
        }
        private static string _GetStateFilter(WorkItemFilter s)
        {
            var filters = new List<string>();
            if (s.ShowProposed) filters.Add("[State] = 'Proposed'");
            if (s.ShowActive) filters.Add("[State] = 'Active'");
            if (s.ShowResolved) filters.Add("[State] = 'Resolved'");
            if (s.ShowClosed) filters.Add("[State] = 'Closed'");

            return filters.Count > 0
              ? @"And (" + String.Join(" Or ", filters) + ") "
              : "And [State] = 'Active' ";
        }
    }
}
