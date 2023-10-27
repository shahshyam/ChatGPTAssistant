using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using AssistantForWord.UI;
using Microsoft.Office.Core;

namespace AssistantForWord
{
    public partial class ThisAddIn
    {
        private SettingUserControl _settingControl;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        public delegate void RefreshSidebarPanel();
        public event RefreshSidebarPanel OnRefreshSidebarPanel;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           
        }
        
        public void ProcessSideBarPanel(bool isVisible)
        {
            if (_settingControl == null)
            {
                _settingControl = new SettingUserControl();
            }
            if (CustomTaskPanes.Count == 0)
            {
                myCustomTaskPane = this.CustomTaskPanes.Add(_settingControl, "AI Assistant", this.Application.ActiveWindow);
                myCustomTaskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;              
                myCustomTaskPane.Width = 400;
            }
            if (isVisible)
            {
                OnRefreshSidebarPanel?.Invoke();
            }
            myCustomTaskPane.Visible = isVisible;
        }
       
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new CustomRibbonExplorer();            
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
