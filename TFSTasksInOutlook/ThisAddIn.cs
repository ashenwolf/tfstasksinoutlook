using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace TFSTasksInOutlook
{
    public partial class ThisAddIn
    {
      private Microsoft.Office.Tools.CustomTaskPane tfsPane;
      private Outlook.NavigationPane navigationPane;
      private TFSTaskPane tfsPaneView = new TFSTaskPane();
      private TFSTaskPaneHost tfsPaneHost;
      private TFSTaskPaneController tfsCtrl;
      private bool isPaneShown;

      private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
          tfsPaneHost = new TFSTaskPaneHost(tfsPaneView);
          tfsPane = this.CustomTaskPanes.Add(tfsPaneHost, "TFS Tasks");
          tfsPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
          tfsPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
          tfsPane.Width = 300;
          tfsCtrl = new TFSTaskPaneController(tfsPaneView);

          navigationPane_ModuleSwitch(this.Application.ActiveExplorer().NavigationPane.CurrentModule);
          navigationPane = this.Application.ActiveExplorer().NavigationPane;
          navigationPane.ModuleSwitch += new Outlook.NavigationPaneEvents_12_ModuleSwitchEventHandler(navigationPane_ModuleSwitch);
        }

      private void navigationPane_ModuleSwitch(Outlook.NavigationModule currentModule)
        {
          tfsPane.Visible = currentModule.NavigationModuleType == Outlook.OlNavigationModuleType.olModuleCalendar && isPaneShown;
        }

      private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

      protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
          return new TFSTasks();
        }

      #region External interfaces
      public void SetPaneShown(bool isShown)
        {
        isPaneShown = isShown;
        navigationPane_ModuleSwitch(this.Application.ActiveExplorer().NavigationPane.CurrentModule);
        }
      #endregion

      #region VSTO generated code

      /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
