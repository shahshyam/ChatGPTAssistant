using AssistantForWord.SaveOption;
using AssistantForWord.UI.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;


namespace AssistantForWord
{
    [ComVisible(true)]
    public class CustomRibbonExplorer : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private bool isPressed;
        private string selectedButtonId = "toggleButton1";
        private List<string> toggleButtonIds;
        private List<PromptDetail> promptDetails;
        private readonly string btnTemplate = @"<button id=""{0}"" tag=""{1}"" label=""{1}"" onAction=""GetSelectedText"" getVisible=""GetSubMenuVisible""/>";
        public CustomRibbonExplorer()
        {
           toggleButtonIds= new List<string>() { "toggleButton1" , "toggleButton2" , "toggleButton3" };
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string ribbonUIContent = GetResourceText("AssistantForWord.CustomRibbonExplorer.xml");
            var config = ProcessData.GetData();
            var sb = new StringBuilder();
            if (!string.IsNullOrEmpty(config.APIKEY))
            {
                promptDetails = config.PromptDetailList;
                foreach (var promot in config.PromptDetailList)
                {
                    string Id = GetIdString();
                    sb.Append(string.Format(btnTemplate, Id, promot.Title));
                }
                if (sb.Length > 0)
                {
                    int index = ribbonUIContent.LastIndexOf("</menu>");
                    ribbonUIContent = ribbonUIContent.Insert(index, sb.ToString());
                }
            }
            return ribbonUIContent;
        }
        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
        public async void GetSelectedText(Office.IRibbonControl control)
        {
            if (control != null)
            {
                string prompt = control.Tag as string;
                string selectedText = WordDocumentHelper.GetSelectedText();
                var result = await OpenAIClient.GetResponse($"{prompt}:{selectedText}");
                WordDocumentHelper.InsertTextAfterSelection(result);
            }
           
        }
        public bool IsVisibleAssistant(Office.IRibbonControl control)
        {
            string selectedText = WordDocumentHelper.GetSelectedText();
            return !string.IsNullOrWhiteSpace(selectedText);
        }
        public bool GetPressed(Office.IRibbonControl control)
        {
            return isPressed;
        } 
        
        public void InvalidateControl(string controlId)
        {
            ribbon.InvalidateControl(controlId);
            promptDetails = null;
        }
        public void ShowSetting(Office.IRibbonControl control, bool isPressed)
        {
            this.isPressed = isPressed;
            InvalidateControl(control.Id);
            Globals.ThisAddIn.ProcessSideBarPanel(this.isPressed);
        }
        public bool GetSubMenuVisible(Office.IRibbonControl control)
        {
            if (control != null)
            {
                string title = control.Tag;
                if (promptDetails == null)
                {
                    var config = ProcessData.GetData();
                    promptDetails = config.PromptDetailList;
                }
                return promptDetails.FindIndex(x => x.Title == title) > -1;
            }
            return false;
        }
        public bool GroupGetPressed(Office.IRibbonControl control)
        {
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
        }        

        private void DeselectAllButtonsExcept(string buttonId)
        {
            //Deselect all buttons except the selected one
            foreach (var id in toggleButtonIds)
            {
                if (selectedButtonId == id)
                    continue;
                ribbon.InvalidateControl(id);
            }
            selectedButtonId = buttonId; // Update the selected button
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
