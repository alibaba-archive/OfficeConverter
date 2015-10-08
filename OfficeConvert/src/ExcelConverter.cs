using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace OfficeConvert
{
    public class ExcelConverter : Converter
    {
        private Excel.Application app;
        private Excel.Workbooks books;
        private Excel.Workbook book;
        // private Excel.Worksheet sheet;

        public override void Convert(String inputFile, String outputFile)
        {
            Object nothing = Type.Missing;
            try
            {
                app = new Excel.Application();
                books = app.Workbooks;
                book = books.Open(inputFile, false, true, nothing, nothing, nothing, true, nothing, nothing, false, false, nothing, false, nothing, false);

                bool hasContent = false;
                foreach (Excel.Worksheet sheet in book.Worksheets)
                {
                    Excel.Range range = sheet.UsedRange;
                    if (range != null) {
                        Excel.Range found = range.Cells.Find("*", nothing, nothing, nothing, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, nothing, nothing, nothing);
                        if (found != null) hasContent = true;
                        releaseCOMObject(found);
                        releaseCOMObject(range);
                    }
                }

                if (!hasContent) throw new ConvertException("No Content");
                book.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputFile, Excel.XlFixedFormatQuality.xlQualityMinimum, false, false, nothing, nothing, false, nothing);
            }
            catch (Exception e)
            {
                release();
                Console.WriteLine(e.Message);
                throw new ConvertException(e.Message);
            }

            release();
        }

        private void release()
        {
            if (book != null)
            {
                try
                {
                    book.Close(false);
                    releaseCOMObject(book);
                }
                catch (Exception e)
                {

                }
            }

            if (books != null)
            {
                try
                {
                    books.Close();
                    releaseCOMObject(books);
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
                    releaseCOMObject(app);
                }
                catch (Exception e)
                {

                }
            }
        }
    }
}
