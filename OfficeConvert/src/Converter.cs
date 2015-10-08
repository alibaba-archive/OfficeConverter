using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeConvert
{
    public abstract class Converter
    {
        public abstract void Convert(String inputFile, String outputFile);

        protected void releaseCOMObject(object obj)
        {
            try
            {
                if (System.Runtime.InteropServices.Marshal.IsComObject(obj))
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
                }
            }
            catch
            {
            }
            obj = null;
        }
    }
}
