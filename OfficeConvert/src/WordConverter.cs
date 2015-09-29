using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;

namespace OfficeConvert
{
    
    public class WordConverter : Converter
    {
        private Word.Application app;
        private Word.Documents docs;
        private Word.Document doc;

        public void Convert(String inputFile, String outputFile)
        {
            Object nothing = System.Reflection.Missing.Value;
            try
            {
                app = new Word.Application();
                docs = app.Documents;
                doc = docs.Open(inputFile, false, true, false, nothing, nothing, true, nothing, nothing, nothing, nothing, false, false, nothing, true, nothing);
                //doc.SaveAs2(outputFile, Word.WdSaveFormat.wdFormatPDF, nothing, nothing, nothing, nothing, nothing, nothing, nothing, nothing, nothing, nothing, nothing, nothing, nothing, nothing, nothing);
                doc.ExportAsFixedFormat(outputFile, Word.WdExportFormat.wdExportFormatPDF, false, Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument, 1, 1, Word.WdExportItem.wdExportDocumentContent, false, false, Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, false, false, false, nothing);
            }
            catch (Exception e)
            {
                throw new ConvertException(e.Message);
            }
            
            if (doc != null)
            {
                try
                {
                    doc.Close(false);
                }
                catch (Exception e)
                {

                }
            }

            if (docs != null)
            {
                try
                {
                    docs.Close(false);
                }
                catch (Exception e)
                {

                }
            }

            if (app != null)
            {
                try
                {
                    app.Quit();
                }
                catch (Exception e)
                {

                }
            }

        }

    }
}
