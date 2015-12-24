using System.Windows.Forms;

namespace TFSTasksInOutlook
{
  public partial class TfsTaskPaneHost : UserControl
  {
    public TfsTaskPaneHost(TfsTaskPane tfsPane)
    {
      InitializeComponent();

      this.elementHost.Dock = System.Windows.Forms.DockStyle.Fill;
      this.elementHost.Child = tfsPane;
    }
  }
}
