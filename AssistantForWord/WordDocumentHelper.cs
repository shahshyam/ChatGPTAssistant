using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
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
                case "btnToggleInsertAfter":
                    wordSelection.InsertAfter(Environment.NewLine + message);
                    break;
                case "btnToggleReplace":
                    ReplaceSelectedText(message);
                    break;
                case "btnToggleInsertOnLocation":
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
            Thread thread = new Thread(() => Clipboard.SetText(updatedText));
            thread.SetApartmentState(ApartmentState.STA); //Set the thread to STA
            thread.Start();
            thread.Join();
            MessageBox.Show("Please paste result on selected location", "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
