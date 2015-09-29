using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeConvert
{
    public interface Converter
    {
        void Convert(String inputFile, String outputFile);
    }
}
