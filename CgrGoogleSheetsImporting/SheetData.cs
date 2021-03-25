using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CgrGoogleSheetsImporting
{
    /// <summary>
    /// Информация о листе Гугл-таблицы
    /// </summary>
    public class SheetData
    {
        public string Range { get; set; }
        public string MajorDimension { get; set; }
        public List<List<string>> Values { get; set; }
    }
}
