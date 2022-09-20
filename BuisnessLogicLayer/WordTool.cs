using Microsoft.Office.Interop.Word;
using Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BuisnessLogicLayer
{
    public static class WordTool
    {
        public static void GenerateLetter(DocTemplate input)
        {
            object docxPathInput = input.FileTempPath;
            object newDocPathOutput = input.OutputPath;
            object UnknownType = Type.Missing;
            object FileFormat = (input.Pdf == true) ? WdSaveFormat.wdFormatPDF : WdSaveFormat.wdFormatDocumentDefault;

            Application MSWordDoc = (Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
            MSWordDoc.Visible = false;
            Documents documents = MSWordDoc.Documents;

            documents.Open(ref docxPathInput, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType);
            MSWordDoc.Application.Visible = false;
            MSWordDoc.WindowState = WdWindowState.wdWindowStateMinimize;

            Document activeDocument = MSWordDoc.ActiveDocument;
            activeDocument.Activate();
            Range range = activeDocument.Content;

            range.Find.Execute(FindText: "[word1]", ReplaceWith: input.Word1, Replace: WdReplace.wdReplaceAll);
            range.Find.Execute(FindText: "[word2]", ReplaceWith: input.Word2, Replace: WdReplace.wdReplaceAll);
            range.Find.Execute(FindText: "[word3]", ReplaceWith: input.Word3, Replace: WdReplace.wdReplaceAll);
            range.Find.Execute(FindText: "[word4]", ReplaceWith: input.Word4, Replace: WdReplace.wdReplaceAll);
            range.Find.Execute(FindText: "[word5]", ReplaceWith: input.Word5, Replace: WdReplace.wdReplaceAll);

            activeDocument.SaveAs(ref newDocPathOutput, ref FileFormat, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType, ref UnknownType);
            activeDocument.Close();
        }
    }
}
