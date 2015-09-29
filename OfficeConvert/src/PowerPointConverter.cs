using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace OfficeConvert
{
    public class PowerPointConverter : Converter
    {

        private PowerPoint.Application app;
        private PowerPoint.Presentations presentations;
        private PowerPoint.Presentation presentation;

        public void Convert(String inputFile, String outputFile)
        {
            try
            {
                app = new PowerPoint.Application();
                presentations = app.Presentations;
                presentation = presentations.Open(inputFile, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                presentation.ExportAsFixedFormat(outputFile, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
                // presentation.ExportAsFixedFormat(outputFile, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF, PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen, MsoTriState.msoFalse, PowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst, PowerPoint.PpPrintOutputType.ppPrintOutputSlides, MsoTriState.msoFalse, null, PowerPoint.PpPrintRangeType.ppPrintAll, "", false, false, false, false, false, null);
            }
            catch (Exception e)
            {
                throw new ConvertException(e.Message);
            }

            if (presentation != null)
            {
                presentation.Close();
            }

            if (app != null)
            {
                app.Quit();
            }
        }
    }
}
