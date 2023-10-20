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
        internal static void InsertTextAfterSelection(string message)
        {
            Word.Selection wordSelection = Globals.ThisAddIn.Application.Selection;
            wordSelection.InsertAfter(Environment.NewLine+message);
        }
    }
}
