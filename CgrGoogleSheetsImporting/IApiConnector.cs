using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CgrGoogleSheetsImporting
{
    public interface IApiConnector
    {
        SpreadSheetsData GetSpreadSheetsData(string spreadsheetId);
        SheetData GetSheetData(string spreadsheetId);
        SheetData GetSheetData(string spreadsheetId, string sheetName);
    }
}
