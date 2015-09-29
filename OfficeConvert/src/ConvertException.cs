using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeConvert
{
    public class ConvertException: Exception
    {
        private String message;

        public String getMessage()
        {
            return message;
        }

        public ConvertException(String message)
        {
            this.message = message;
        }
    }
}
