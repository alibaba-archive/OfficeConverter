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
            Object nothing = Type.Missing;
            try
            {
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
                try
                {
                    book.Close(false);
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
