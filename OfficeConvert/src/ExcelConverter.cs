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
        private Excel.Worksheet sheet;

        public void Convert(String inputFile, String outputFile)
        {
            try
            {
                Object nothing = System.Reflection.Missing.Value;
                app = new Excel.Application();
                books = app.Workbooks;
                book = books.Open(inputFile, false, true, nothing, nothing, nothing, true, nothing, nothing, false, false, nothing, false, nothing, false);
                book.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputFile, Excel.XlFixedFormatQuality.xlQualityMinimum, false, false, 1, 1, false, nothing);
            }
            catch (Exception e)
            {
                throw new ConvertException(e.Message);
            }

            if (book != null)
            {
                book.Close();
            }

            if (books != null)
            {
                books.Close();
            }

            if (app != null)
            {
                app.Quit();
            }
        }
    }
}
