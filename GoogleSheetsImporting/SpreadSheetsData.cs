using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CgrGoogleSheetsImporting
{
    public class SpreadSheetsData
    {
        public string SpreadsheetId { get; set; }
        public Properties Properties { get; set; }
        public List<Sheet> Sheets { get; set; }
        public List<NamedRanx> NamedRanges { get; set; }
        public string SpreadsheetUrl { get; set; }
    }
    public class Properties
    {
        public string Title { get; set; }
        public string Locale { get; set; }
        public string AutoRecalc { get; set; }
        public string TimeZone { get; set; }
        public int SheetId { get; set; }
        public int Index { get; set; }
        public string SheetType { get; set; }
        public GridProperties GridProperties { get; set; }
    }

    public class GridProperties
    {
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
    }

    public class Sheet
    {
        public Properties Properties { get; set; }
    }

    public class Range
    {
        public int SheetId { get; set; }
        public int StartColumnIndex { get; set; }
        public int EndColumnIndex { get; set; }
        public int? StartRowIndex { get; set; }
        public int? EndRowIndex { get; set; }
    }

    public class NamedRanx
    {
        public string NamedRangeId { get; set; }
        public string Name { get; set; }
        public Range Range { get; set; }
    }

}
