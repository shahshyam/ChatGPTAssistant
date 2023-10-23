using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace AssistantForWord
{
    internal class WordDocumentHelper
    {
        internal static string GetSelectedText()
        {
            string selectText = string.Empty;
            Word.Selection wordSelection = Globals.ThisAddIn.Application.Selection;
            if (wordSelection != null && wordSelection.Range != null)
            {
                selectText = wordSelection.Text;
            }
            return selectText;
        }
        internal static void InsertTextAfterSelection(string message, string selectedPosition)
        {
            Word.Selection wordSelection = Globals.ThisAddIn.Application.Selection;
            switch(selectedPosition)
            {
                case "toggleButton1":
                    wordSelection.InsertAfter(Environment.NewLine + message);
                    break;
                case "toggleButton2":
                    ReplaceSelectedText(message);
                    break;
                case "toggleButton3":
                    InsertOnSelectedLocation(Environment.NewLine + message);
                    break;
            }
            
        }
        private static void ReplaceSelectedText(string updatedText)
        {
            Word.Application application = Globals.ThisAddIn.Application;
            if (application.Selection != null && application.Selection.Type == Word.WdSelectionType.wdSelectionNormal)
            {
                string selectedText = application.Selection.Text;
                if (!string.IsNullOrEmpty(selectedText))
                    application.Selection.Text = updatedText;
            }
        }

        private static void InsertOnSelectedLocation(string updatedText)
        {
            Word.Application application = Globals.ThisAddIn.Application;
            Word.Document document = application.ActiveDocument;          
         
        }
    }
}
