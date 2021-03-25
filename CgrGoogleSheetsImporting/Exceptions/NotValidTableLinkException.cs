using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CgrGoogleSheetsImporting.Exceptions
{
    public class NotValidTableLinkException : Exception
    {
        public NotValidTableLinkException(string message) :
                base(message)
        {
        }
    }
}
