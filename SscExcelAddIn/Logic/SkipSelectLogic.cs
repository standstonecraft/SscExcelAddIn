using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.Logic
{
    public class SkipSelectLogic
    {
        private const string ExceptNumPrn = "[^0-9,]";
        private const string NumArrayPtn = @"^(\d+)(,*\d+)*$";
        public static void SkipSelectRange(Excel.Range range, string selector, bool isColumn = false)
        {
            Excel.Range selection = Funcs.CellSelection();
            if (selection is null)
            {
                return;
            }
            // validation
            if (!Regex.IsMatch(selector, NumArrayPtn))
            {
                return;
            }
            int uBound = isColumn ? range.Columns.Count : range.Rows.Count;
            IEnumerable<int> skipSelector = selector.Split(',').Where(s => s != "").Select(s => int.Parse(s));
            SkipFilter<int> skipFilter = new SkipFilter<int>(Enumerable.Range(1, uBound), skipSelector);
            string selected = string.Join(",", skipFilter.Select(target => isColumn ?
                     (Excel.Range)range.Columns.Item[target] :
                     (Excel.Range)range.Rows.Item[target])
                .Select(o => ((Excel.Range)o).Address));

            if (selected != "")
            {
                Excel.Worksheet activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                activeSheet.Range[selected].Select();
            }
        }
    }
}
