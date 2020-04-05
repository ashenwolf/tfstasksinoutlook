using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TFSTasksInOutlook.TFS;

namespace TFSTasksInOutlook.Tests
{
    [TestClass]
    public class WorkItemFilterQueryBuilder_Build
    {
        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void WhenFilterIsNull_ThrowException()
        {
            WorkItemFilter filter = null;
            WorkItemFilterQueryBuilder.Build(filter);
        }
        [TestMethod]
        public void WhenFilterHasAllValuesButStartAndFinishDates_Success()
        {
            WorkItemFilter filter = new WorkItemFilter()
            {
                Project = "TestTfsProj",
                ShowTasks = true,
                ShowBugs = false,
                ShowProposed = false,
                ShowActive = true,
                ShowResolved = true,
                ShowClosed = false,
            };
            string actual = WorkItemFilterQueryBuilder.Build(filter);
            string expected = "Select [Id], [Title], [Completed Work], [System.TeamProject] From WorkItems Where [System.TeamProject] = 'TestTfsProj' And [System.AssignedTo] = @Me And ([Work Item Type] = 'Task') And ([State] = 'Active' Or [State] = 'Resolved') Order By [Work Item Type] ";
            Assert.AreEqual(expected, actual);
        }
        [TestMethod]
        public void WhenFilterHasAllValuesButStartAndFinishDatesAndShowsStartAndFinishDatesSetToTrue_Success()
        {
            WorkItemFilter filter = new WorkItemFilter()
            {
                Project = "TestTfsProj",
                ShowTasks = true,
                ShowBugs = false,
                ShowProposed = false,
                ShowActive = true,
                ShowResolved = true,
                ShowClosed = false,
                ShowStartAndFinishDates = true
            };
            string actual = WorkItemFilterQueryBuilder.Build(filter);
            string expected = "Select [Id], [Title], [Completed Work], [System.TeamProject], [Start Date], [Finish Date] From WorkItems Where [System.TeamProject] = 'TestTfsProj' And [System.AssignedTo] = @Me And ([Work Item Type] = 'Task') And ([State] = 'Active' Or [State] = 'Resolved') Order By [Work Item Type] ";
            Assert.AreEqual(expected, actual);
        }
        [TestMethod]
        public void WhenFilterHasAllValues_Success()
        {
            WorkItemFilter filter = new WorkItemFilter()
            {
                Project = "TestTfsProj",
                ShowTasks = true,
                ShowBugs = false,
                ShowProposed = false,
                ShowActive = true,
                ShowResolved = true,
                ShowClosed = false,
                StartDate = new DateTime(2020,4,1),
                FinishDate = new DateTime(2020, 4, 1).AddDays(1)
            };
            string actual = WorkItemFilterQueryBuilder.Build(filter);
            string expected = "Select [Id], [Title], [Completed Work], [System.TeamProject] " +
                              "From WorkItems " +
                              "Where [System.TeamProject] = 'TestTfsProj' " +
                                    "And [System.AssignedTo] = @Me " +
                                    "And ([Work Item Type] = 'Task') " +
                                    "And ([State] = 'Active' Or [State] = 'Resolved') " +
                                    "And [Start Date] >= '04/01/2020 00:00:00Z' And [Finish Date] <= '04/02/2020 00:00:00Z' " +
                              "Order By [Work Item Type] ";
            Assert.AreEqual(expected, actual,$"Actual message:{actual}");
        }
    }
}
