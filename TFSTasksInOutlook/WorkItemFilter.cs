using System;

namespace TFSTasksInOutlook
{
    public class WorkItemFilter
    {
        public string Project { get; set; }
        public bool ShowTasks { get; set; }
        public bool ShowBugs { get; set; }
        public bool ShowProposed { get; set; }
        public bool ShowActive { get; set; }
        public bool ShowResolved { get; set; }
        public bool ShowClosed { get; set; }

        public DateTime? FinishDate { get; set; }
        public DateTime? StartDate { get; set; }
    }
}