using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.Logic
{
    public class BorderGridTableLogic
    {
        private static readonly Excel.WorksheetFunction xlFunc = Globals.ThisAddIn.Application.WorksheetFunction;
        private const Excel.XlLineStyle xlLineStyleNone = Excel.XlLineStyle.xlLineStyleNone;
        private const Excel.XlLineStyle xlContinuous = Excel.XlLineStyle.xlContinuous;
        private const Excel.XlBordersIndex xlInsideHorizontal = Excel.XlBordersIndex.xlInsideHorizontal;
        public static void BorderGridTable()
        {
            Excel.Range selection = Funcs.CellSelection();
            if (selection == null)
            {
                return;
            }
            int rows = selection.Rows.Count;
            int colCount = 1;
            for (int c = selection.Columns.Count; c >= 1; c--)
            {
                if (xlFunc.CountA(selection.Columns[c]) > 0)
                {
                    Excel.Range range = ((Excel.Range)selection.Columns[c]).Resize[rows, colCount];
                    range.Borders.LineStyle = xlLineStyleNone;
                    range.BorderAround2(xlContinuous);
                    range.Borders[xlInsideHorizontal].LineStyle = xlContinuous;
                    colCount = 1;
                }
                else
                {
                    colCount++;
                }
            }
        }
    }
}
