using AssistantForWord.SaveOption;
using AssistantForWord.UI.Models;
using System;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using System.Drawing;

namespace AssistantForWord
{
    [ComVisible(true)]
    public class CustomRibbonExplorer : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;       
        private bool isFirstTimeDisplay;
        public  const string defaultInsertion = "btnToggleInsertAfter";
        private string selectedButtonId = defaultInsertion;
        private List<string> toggleButtonIds;        
        private readonly string btnTemplate = @"<button id=""{0}"" tag=""{1}"" label=""{2}"" onAction=""GetSelectedText"" imageMso=""AutoSummarize""/>";
        public CustomRibbonExplorer()
        {
           toggleButtonIds= new List<string>() { "btnToggleInsertAfter", "btnToggleReplace", "btnToggleInsertOnLocation" };
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string ribbonUIContent = GetResourceText("AssistantForWord.CustomRibbonExplorer.xml");            
            return ribbonUIContent;
        }
        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
        public async void GetSelectedText(Office.IRibbonControl control)
        {
            if (control != null)
            {
                var application = Globals.ThisAddIn.Application;
                var oldStatus = application.DisplayStatusBar;
                application.DisplayStatusBar = true;
                application.StatusBar = "Assistant is processing request";
                string promptId = control.Tag;
                string selectedText = WordDocumentHelper.GetSelectedText();
                var promptDetail = ProcessData.GetPromptDetailById(promptId);
                if (promptDetail != null)
                {
                    var result = await OpenAIClient.GetResponse($"{promptDetail.Name} {selectedText}");
                    WordDocumentHelper.InsertTextAfterSelection(result, selectedButtonId);
                }
                application.Selection.Collapse();
                application.DisplayStatusBar = oldStatus;
                application.DisplayStatusBar = false;
                application.DisplayAutoCompleteTips = true;
                
            }
           
        }
        public bool IsVisibleAssistant(Office.IRibbonControl control)
        {
            string selectedText = WordDocumentHelper.GetSelectedText();
            return !string.IsNullOrWhiteSpace(selectedText);
        }
        public bool GetPressed(Office.IRibbonControl control)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Globals.ThisAddIn.Dictdocument.TryGetValue(document, out bool result);
            return result;
        }
        public Bitmap GetImageLocally(Office.IRibbonControl control)
        {
            if (control.Id == "btnPanel")
                return Properties.Resources.window_gear;
            else if (control.Id == "btnToggleReplace")
                return Properties.Resources.replace;
            else if (control.Id == "btnToggleInsertOnLocation")
                return Properties.Resources.pointer;
            else if (control.Id == "btnToggleInsertAfter")
                return Properties.Resources.button_angle_right;
            return null;
        }
        public void InvalidateControl(string controlId)
        {
            ribbon.InvalidateControl(controlId);           
        }
        public void ShowSetting(Office.IRibbonControl control, bool isPressed)
        {            
            InvalidateControl(control.Id);
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Globals.ThisAddIn.ProcessSideBarPanel(document, isPressed);
        }
       
        public bool GroupGetPressed(Office.IRibbonControl control)
        {
            if (!isFirstTimeDisplay)
                return false;
            return selectedButtonId == control.Id;
        }

        public void ToggleButton_Click(Office.IRibbonControl control, bool isPressed)
        {
            // Handle the click event
            if (isPressed)
            {
                selectedButtonId = control.Id;
                // Deselect all other buttons
                DeselectAllButtonsExcept(control.Id);
            }
            else
            {
                //default value restore
                selectedButtonId = defaultInsertion;
            }
        }      

        private void DeselectAllButtonsExcept(string buttonId)
        {
            isFirstTimeDisplay = true;
            //Deselect all buttons except the selected one
            foreach (var id in toggleButtonIds)
            {
                if (selectedButtonId == id)
                    continue;
                ribbon.InvalidateControl(id);
            }
            selectedButtonId = buttonId; // Update the selected button
        }

        public string GetMenuContent(Office.IRibbonControl control)
        {
            var config = ProcessData.GetData();
            bool hasContentAdded = false;

            StringBuilder sb = new StringBuilder(@"<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" >");

            if (!string.IsNullOrEmpty(config.APIKEY) && config.PromptDetailList.Count > 0)
            {
                hasContentAdded = true;
                foreach (var promot in config.PromptDetailList)
                {
                    string Id = $"btn{GetIdString()}";
                    sb.Append(string.Format(btnTemplate, Id, promot.Id, promot.Title));
                }
            }
            if (hasContentAdded)
                sb.Append(@"</menu>");
            return sb.ToString();
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers
        private static Random random = new Random();
        public static string GetIdString()
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            return new string(Enumerable.Repeat(chars, 10)
                .Select(s => s[random.Next(s.Length)]).ToArray());
        }
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
