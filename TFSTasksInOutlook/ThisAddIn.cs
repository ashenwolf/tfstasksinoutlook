using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace TFSTasksInOutlook
  {
  public partial class ThisAddIn
    {
    private Microsoft.Office.Tools.CustomTaskPane _tfsPane;
    private Outlook.NavigationPane _navigationPane;
    private readonly TfsTaskPane _tfsPaneView = new TfsTaskPane();
    private TfsTaskPaneHost _tfsPaneHost;
    private TfsTaskPaneController _tfsCtrl;
    private bool _isPaneShown;

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
      {
      _tfsPaneHost = new TfsTaskPaneHost(_tfsPaneView);
      _tfsPane = this.CustomTaskPanes.Add(_tfsPaneHost, "TFS Tasks");
      _tfsPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
      _tfsPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
      _tfsPane.Width = 300;
      _tfsCtrl = new TfsTaskPaneController(_tfsPaneView);

      navigationPane_ModuleSwitch(this.Application.ActiveExplorer().NavigationPane.CurrentModule);
      _navigationPane = this.Application.ActiveExplorer().NavigationPane;
      _navigationPane.ModuleSwitch += new Outlook.NavigationPaneEvents_12_ModuleSwitchEventHandler(navigationPane_ModuleSwitch);
      }

    private void navigationPane_ModuleSwitch(Outlook.NavigationModule currentModule)
      {
      _tfsPane.Visible = currentModule.NavigationModuleType == Outlook.OlNavigationModuleType.olModuleCalendar && _isPaneShown;
      }

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
      {
      }

    protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
      {
      return new TfsTasks();
      }

    #region External interfaces
    public void SetPaneShown(bool isShown)
      {
      _isPaneShown = isShown;
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
