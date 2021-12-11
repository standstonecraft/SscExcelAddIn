using System;
using System.Collections.Generic;
using System.Deployment.Application;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Reactive.Bindings;
using SscExcelAddIn.ComModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.Logic
{
    /// <summary>
    /// <see cref="Ribbon1"/>のロジック
    /// </summary>
    public class Ribbon1Logic
    {
        /// <summary>
        /// GitHubのリリース情報に非同期でアクセスして更新確認を行う。
        /// バージョン比較にはタグを使用する。ドット区切りであればその個数は問わない。
        /// デバッグ時は現行バージョンが確認できないためv0.0.0.1として扱う。
        /// </summary>
        /// <param name="updateNotifyCommand">新しいバージョンがある場合に起動するCommand</param>
        public static void CheckUpdate(ReactiveCommand<string> updateNotifyCommand)
        {
            _ = Task.Run(() =>
            {
                string currentVersion;
                try
                {
                    currentVersion = "v" + ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
                }
                catch (Exception)
                {
                    currentVersion = "v0.0.0.1";
                }
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Properties.Resources.ReleaseApiUrl);
                request.ContentType = "application/json; charset=utf-8";
                request.UserAgent = @"Mozilla/5.0 (iPhone; CPU iPhone OS 14_5 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) CriOS/91.0.4472.80 Mobile/15E148 Safari/604.1";

                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                using (Stream responseStream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                    dynamic json = JsonConvert.DeserializeObject(reader.ReadToEnd());
                    string publishedVersion = json.tag_name;
                    if (longVersion(currentVersion) < longVersion(publishedVersion))
                    {
                        updateNotifyCommand.Execute($"{currentVersion} => {publishedVersion}");
                    }

                }
                double longVersion(string verStr)
                {
                    string numStr = verStr.Replace("v", "");
                    double ret = 0;
                    string[] vs = numStr.Split('.');
                    for (int i = 0; i < vs.Length; i++)
                    {
                        ret += long.Parse(vs[i]) * Math.Pow(100, 4 - i);
                    }
                    return ret;
                }
            });
        }

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

        /// <summary>
        /// 行列の操作により分割された条件付き書式を統合する
        /// </summary>
        public static void MergeFormatConds()
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            while (true)
            {
                IGrouping<FormatConditionModel, FormatConditionModel> group =
                    sheet.UsedRange.FormatConditions.Cast<Excel.FormatCondition>()
                    .Select(fc => new FormatConditionModel(fc))
                    .GroupBy(fc => fc)
                    .FirstOrDefault(cg => cg.Count() > 1);
                if (group == null)
                {
                    break;
                }

                group.ElementAt(0).FormatCondition.ModifyAppliesToRange(
                    Funcs.UnionRange(group.Select(fcm => fcm.FormatCondition.AppliesTo).ToList()));
                foreach (FormatConditionModel item in group.Skip(1))
                {
                    item.FormatCondition.Delete();
                }
            }
        }
    }
}
