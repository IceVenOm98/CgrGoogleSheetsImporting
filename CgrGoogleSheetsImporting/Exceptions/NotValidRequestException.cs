using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace CgrGoogleSheetsImporting
{
    public class NotValidRequestException : Exception
    {
        public NotValidRequestException(Exception e, string message) :
                base(message, e)
        {
        }public NotValidRequestException(string message) :
                base(message)
        {
        }
    }
}
