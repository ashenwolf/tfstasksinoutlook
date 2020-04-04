using System;

namespace TFSTasksInOutlook
{
    public partial class TfsTaskPane
    {
        public class WorkItemActionEventArgs : EventArgs
        {
            public WorkItemActionEventArgs(WorkItemInfo item) { Item = item; }
            public WorkItemInfo Item { get; private set; }
        }
    }
}
