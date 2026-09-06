using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using AssistantForWord.UI;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using AssistantForWord.UI.Helpers;

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
                application.DocumentOpen += ThisAddIn_NewDocument;
                application.DocumentBeforeClose += OnCloseDocument;
                ((Word.ApplicationEvents4_Event)application).Quit += ThisAddIn_Quit; ;
            }
        }

        private void ThisAddIn_Quit()
        {
            ((Word.ApplicationEvents4_Event)application).NewDocument -= ThisAddIn_NewDocument;
            application.DocumentOpen -= ThisAddIn_NewDocument;
            application.DocumentBeforeClose -= OnCloseDocument;
            ((Word.ApplicationEvents4_Event)application).Quit -= ThisAddIn_Quit;
        }

        private void OnCloseDocument(Word.Document document, ref bool cancel)
        {
            if (document != null)
            {
                Dictdocument.Remove(document);
                Marshal.ReleaseComObject(document);
            }
        }

        private void ThisAddIn_NewDocument(Word.Document Doc)
        {
            Word.Document document = null;
            if (Application.Documents.Count > 0)
            {
                document = Application.ActiveDocument;
            }

            if (document == null)
            {
                document = Doc; // fallback to event document
            }

            if (document != null)
            {
                ProcessSideBarPanel(document, false);
            }
            // else: no document available — skip or log
        }

        public void ProcessSideBarPanel(Word.Document document, bool isVisible)
        {            
            if (!Dictdocument.ContainsKey(document))
            {
                _settingControl = new SettingUserControl();                          
                myCustomTaskPane = this.CustomTaskPanes.Add(_settingControl, AppConstant.AppTitle, document?.ActiveWindow);
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
