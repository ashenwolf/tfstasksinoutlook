using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TFSTasksInOutlook
{
  public partial class TFSTaskPaneHost : UserControl
  {
    public TFSTaskPaneHost(TFSTaskPane tfsPane)
    {
      InitializeComponent();

      this.elementHost.Dock = System.Windows.Forms.DockStyle.Fill;
      this.elementHost.Child = tfsPane;
    }
  }
}
