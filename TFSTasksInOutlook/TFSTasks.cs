using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Core;
using TFSTasksInOutlook.Properties;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  


namespace TFSTasksInOutlook
{
  [ComVisible(true)]
  public class TFSTasks : Office.IRibbonExtensibility
  {
    private Office.IRibbonUI ribbon;

    public TFSTasks()
    {
    }

    #region IRibbonExtensibility Members

    public string GetCustomUI(string ribbonID)
    {
      return GetResourceText("TFSTasksInOutlook.TFSTasks.xml");
    }

    #endregion

    #region Ribbon Callbacks
    //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

    public void Ribbon_Load(Office.IRibbonUI ribbonUI)
    {
      this.ribbon = ribbonUI;
    }

    public Bitmap Ribbon_GetImage(IRibbonControl control)
    {
      switch (control.Id)
        {
        case "DentsplyTools_TFSTasks": return Resources.tfs_task;
        }
      return null;
    }

    public void Ribbon_TFSTaskActivate(IRibbonControl control, bool pressed)
      {
      switch (control.Id)
        {
        case "DentsplyTools_TFSTasks":
          Globals.ThisAddIn.SetPaneShown(pressed);
          return;
        }
      }

    #endregion

    #region Helpers

    private static string GetResourceText(string resourceName)
    {
      Assembly asm = Assembly.GetExecutingAssembly();
      string[] resourceNames = asm.GetManifestResourceNames();
      for (int i = 0; i < resourceNames.Length; ++i)
      {
        if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
        {
          using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
          {
            if (resourceReader != null)
            {
              return resourceReader.ReadToEnd();
            }
          }
        }
      }
      return null;
    }

    #endregion
  }
}
