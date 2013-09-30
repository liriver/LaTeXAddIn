using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace latex
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if (Application.ActiveSheet != null)
            {
                Globals.Ribbons.Ribbon.tableButton.Enabled = true;
                Globals.Ribbons.Ribbon.clipboardBox.Enabled = true;
                Globals.Ribbons.Ribbon.fileBox.Enabled = true;
            }
            Application.WindowActivate += new Excel.AppEvents_WindowActivateEventHandler(Application_WindowActivate);
            Application.WindowDeactivate += new Excel.AppEvents_WindowDeactivateEventHandler(Application_WindowDeactivate);
        }

        void Application_WindowDeactivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            Globals.Ribbons.Ribbon.tableButton.Enabled = false;
            Globals.Ribbons.Ribbon.clipboardBox.Enabled = false;
            Globals.Ribbons.Ribbon.fileBox.Enabled = false;
        }

        void Application_WindowActivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            Globals.Ribbons.Ribbon.tableButton.Enabled = true;
            Globals.Ribbons.Ribbon.clipboardBox.Enabled = true;
            Globals.Ribbons.Ribbon.fileBox.Enabled = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

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
