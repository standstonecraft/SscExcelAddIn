using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.Logic
{

    /// <summary>
    /// 空行/空列削除
    /// </summary>
    public class RemoveEmptyRowLogic
    {
        /// <summary>
        /// 空列削除
        /// </summary>
        public static void RemoveEmptyCol()
        {
            try
            {
                Globals.ThisAddIn.Application.Interactive = false;
                Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                Excel.Range range = Funcs.CellSelection();
                if (range is null)
                {
                    return;
                }
                int colStart = range.Column;
                int colEnd = colStart + range.Columns.Count - 1;
                for (int col = colEnd; col >= colStart; col--)
                {
                    int countA = (int)Globals.ThisAddIn.Application.WorksheetFunction.CountA(sheet.Columns[col]);
                    if (countA == 0)
                    {
                        ((Excel.Range)sheet.Columns[col]).EntireColumn.Delete();
                    }
                }
            }
            finally
            {
                Globals.ThisAddIn.Application.Interactive = true;
            }
        }

        /// <summary>
        /// 空行削除
        /// </summary>
        public static void RemoveEmptyRow()
        {
            try
            {
                Globals.ThisAddIn.Application.Interactive = false;
                Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                Excel.Range range = Funcs.CellSelection();
                if (range is null)
                {
                    return;
                }
                int rowStart = range.Row;
                int rowEnd = rowStart + range.Rows.Count - 1;
                for (int row = rowEnd; row >= rowStart; row--)
                {
                    int countA = (int)Globals.ThisAddIn.Application.WorksheetFunction.CountA(sheet.Rows[row]);
                    if (countA == 0)
                    {
                        ((Excel.Range)sheet.Rows[row]).EntireRow.Delete();
                    }
                }
            }
            finally
            {
                Globals.ThisAddIn.Application.Interactive = true;
            }
        }
    }
}
