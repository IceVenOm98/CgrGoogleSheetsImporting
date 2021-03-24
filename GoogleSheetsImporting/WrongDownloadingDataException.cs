using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace CgrGoogleSheetsImporting
{
    class WrongDownloadingDataException : Exception
    {
        const string WrongDownloadingDataMessage = "Ошибка получения информации из гугл-таблицы. Проверьте Id таблицы и названия листов.";
        public WrongDownloadingDataException(WebException e) :
                base(WrongDownloadingDataMessage, e)
        {

        }
    }
}
