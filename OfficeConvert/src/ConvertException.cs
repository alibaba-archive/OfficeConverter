using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeConvert
{
    public class ConvertException: Exception
    {
        public ConvertException(String message) : base(message)
        {
        }
    }
}
