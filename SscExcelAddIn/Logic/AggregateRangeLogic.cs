using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.Logic
{
    /// <summary>
    /// 集計テーブル
    /// </summary>
    public class AggregateRangeLogic
    {
        /// <summary>
        /// 集計テーブル
        /// </summary>
        public static void AggregateRange()
        {
            Excel.Range range = Funcs.CellSelection();
            if (range is null || range.Count == 1)
            {
                return;
            }
            // ユニーク化
            object[,] rval = (object[,])range.Value2;
            int colSize = rval.GetLength(1);
            List<string> uniqueRows = new List<string>();
            for (int ridx = 1; ridx <= rval.GetLength(0); ridx++)
            {
                StringBuilder uniqueSb = new StringBuilder();
                for (int cidx = 1; cidx <= colSize; cidx++)
                {
                    string key = rval[ridx, cidx].ToString();
                    uniqueSb.Append(key);
                    uniqueSb.Append("\t");
                }

                string uniqueRow = uniqueSb.ToString();
                if (!uniqueRows.Contains(uniqueRow))
                {
                    uniqueRows.Add(uniqueRow);
                }
            }

            // キー列範囲のアドレス
            string[] colAddresses = new string[colSize];
            for (int cidx = 1; cidx <= colSize; cidx++)
            {
                colAddresses[cidx - 1] = ((Excel.Range)range.Columns[cidx]).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1);
            }

            // 結果作成
            object[,] value2s = new object[uniqueRows.Count, colSize];
            string[] formulas = new string[uniqueRows.Count];
            int dicIdx = 0;
            foreach (string uniqueRow in uniqueRows)
            {
                List<string> formulaArgs = new List<string>();
                string[] vs = uniqueRow.Split('\t');
                for (int keyIdx = 0; keyIdx < vs.Length - 1; keyIdx++)
                {
                    value2s[dicIdx, keyIdx] = vs[keyIdx];
                    formulaArgs.Add(colAddresses[keyIdx]);
                    formulaArgs.Add($"R[0]C[-{vs.Length - keyIdx - 1}]");
                }
                formulas[dicIdx] = string.Format("=COUNTIFS({0})", string.Join(",", formulaArgs));

                dicIdx++;
            }

            // 結果を貼り付け
            int offset = range.Rows.Count + 2;
            range.Offset[offset].Resize[dicIdx, colSize].Value2 = value2s;
            Excel.Range resultRange = range.Offset[offset, colSize].Resize[formulas.Length, 1];
            for (int i = 0; i < formulas.Length; i++)
            {
                ((Excel.Range)resultRange.Cells[i + 1, 1]).FormulaR1C1 = formulas[i];
            }
        }
    }
}
