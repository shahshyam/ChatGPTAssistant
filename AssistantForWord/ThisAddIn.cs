using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
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
        public Dictionary<Word.Document, bool> Dictdocument = new Dictionary<Word.Document, bool>();
        Word.Application application;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            application = this.Application;
            if (application != null)
            {
                ((Word.ApplicationEvents4_Event)application).NewDocument += ThisAddIn_NewDocument;
                ((Word.ApplicationEvents4_Event)application).DocumentOpen += ThisAddIn_NewDocument;
            }
        }

        private void ThisAddIn_NewDocument(Word.Document Doc)
        {
            Word.Document document = Application.ActiveDocument;
            ProcessSideBarPanel(document, false);
        }

        public void ProcessSideBarPanel(Word.Document document, bool isVisible)
        {            
            if (!Dictdocument.ContainsKey(document))
            {
                _settingControl = new SettingUserControl();                          
                myCustomTaskPane = this.CustomTaskPanes.Add(_settingControl, "Writing Assistant", document?.ActiveWindow);
                myCustomTaskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                myCustomTaskPane.Width = 400;
            }            
            if (isVisible)
            {
                OnRefreshSidebarPanel?.Invoke();
            }
            myCustomTaskPane.Visible = isVisible;
            Dictdocument[document] = isVisible;
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
