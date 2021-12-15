using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.Logic
{
    /// <summary>
    /// 方眼紙上に作られた表を見きれないようにリサイズする
    /// </summary>
    public class ResizeGridTableLogic
    {
        private const Excel.XlSheetVisibility xlSheetHidden = Excel.XlSheetVisibility.xlSheetHidden;
        private const Excel.XlBordersIndex xlEdgeRight = Excel.XlBordersIndex.xlEdgeRight;
        private const Excel.XlLineStyle xlLineStyleNone = Excel.XlLineStyle.xlLineStyleNone;
        private const Excel.XlDeleteShiftDirection xlShiftToLeft = Excel.XlDeleteShiftDirection.xlShiftToLeft;
        private const Excel.XlInsertShiftDirection xlShiftToRight = Excel.XlInsertShiftDirection.xlShiftToRight;
        private const Excel.XlInsertFormatOrigin xlFormatFromRightOrBelow = Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow;
        private static readonly Encoding enc = Encoding.GetEncoding("Shift_JIS");

        /// <summary>
        /// 方眼紙上に作られた表を見きれないようにリサイズする
        /// </summary>
        public static void ResizeGridTable()
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range original = Funcs.CellSelection();
            if (original == null)
            {
                return;
            }
            int row = original.Row;
            int col = original.Column;
            int cols = original.Columns.Count;
            int rows = original.Rows.Count;
            // 計算用のシート
            Excel.Worksheet tmpSh = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            tmpSh.Visible = xlSheetHidden;
            // クリップボードを使用せず書式ごとコピーする
            original.Copy(tmpSh.Cells[1, 1]);
            Excel.Range paste = Funcs.Range(tmpSh, 1, 1, rows, cols);
            // 数式の場合があるので値をコピーする
            paste.Value2 = original.Value2;

            // 値が入っている列を調べる
            // ex. +0, +0, -1, +0, -1, -1, +0
            List<int> filledList = paste.Columns.Cast<Excel.Range>()
                .Select(c => Globals.ThisAddIn.Application.WorksheetFunction.CountA(c) > 0 ? 0 : -1).ToList();
            // 値が入っていない列を消す
            IEnumerable<int> forDel = Enumerable.Range(1, filledList.Count)
                .Reverse().Where(i => filledList[i - 1] < 0);
            foreach (int fi in forDel)
            {
                Funcs.Range(tmpSh, 1, fi, rows, 1).Delete(xlShiftToLeft);
            }
            // 結合を解除
            paste = Unmerge(tmpSh, rows, cols - filledList.Count(fi => fi < 0));
            // 幅を自動調整
            paste.EntireColumn.AutoFit();
            // 最適な幅を取得
            IEnumerator<double> widths = ((Excel.Range)paste.Rows[1]).Cells.Cast<Excel.Range>().Select(r => (double)r.Width).GetEnumerator();
            // 元の領域の幅を多めに取得
            IEnumerator<dynamic> origWidths = original.Resize[1, 10000].Cast<Excel.Range>().Select(r => r.Width).GetEnumerator();
            // 最適な幅になるには元の領域だと何列ずつ必要か調べる
            // ex. +1, +4, -1, +3, -1, -1, +1
            for (int fi = 0; fi < filledList.Count; fi++)
            {
                // 値が入っている列のみ処理
                if (filledList[fi] > -1)
                {
                    widths.MoveNext();
                    double total = 0;
                    origWidths.MoveNext();
                    while ((total += origWidths.Current) < widths.Current)
                    {
                        // 必要なセル数をインクリメント
                        origWidths.MoveNext();
                        filledList[fi]++;
                        // 最適な幅に達するまで
                    }
                }
            }
            // 値が入っていない消すべき列をまとめる
            // ex. +1, +3, +0, +1, +0, +0, +1
            for (int i = filledList.Count - 1; i >= 1; i--)
            {
                // 消すべき列なら
                if (filledList[i] < 0)
                {
                    // 手前と自分どちらかが0になるまで-1ずつ手前に移す
                    // 単純に移すと、増やす列＜消す列のときに負となって自身が消す列になってしまう
                    while (filledList[i - 1] != 0 && filledList[i] != 0)
                    {
                        filledList[i - 1]--;
                        filledList[i]++;
                    }
                }
            }
            // 増える列数
            int diffCol = filledList.Sum();
            // 最後の列
            int lastCol = col + cols + diffCol - 1;
            string lastColAddress = Regex.Replace(Funcs.Range(sheet, row, lastCol).Address, @"[$\d]", "");
            MessageBoxResult answer = MessageBoxResult.OK;
            if (diffCol > 0)
            {
                // 列が増える場合は確認する
                answer = MessageBox.Show(
                    $"実行すると{lastColAddress}列に達し、その範囲の内容は上書きされます。実行しますか？",
                    "SscExcelAddIn", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            }
            if (answer == MessageBoxResult.OK)
            {
                // 元の領域を見きれないようにシフトする
                for (int fi = filledList.Count - 1; fi >= 0; fi--)
                {
                    if (filledList[fi] > 0)
                    {
                        // 値が入っている列ではその右に空の列を挿入する
                        Funcs.Range(sheet, row, col + fi + 1, rows, filledList[fi]).Insert(xlShiftToRight);
                        // 右辺の枠は不要になるので消す
                        Funcs.Range(sheet, row, col + fi, rows, 1).Borders[xlEdgeRight].LineStyle = xlLineStyleNone;
                    }
                    else if (filledList[fi] < 0)
                    {
                        // 0は無視する
                        // 値が入っていない列は削除する
                        Funcs.Range(sheet, row, col + fi, rows, -filledList[fi]).Delete();
                    }
                }
                if (diffCol > 0)
                {
                    // 増えた列を消す
                    Funcs.Range(sheet, row, lastCol + 1, rows, diffCol).Delete(xlShiftToLeft);
                }
                else if (diffCol < 0)
                {
                    // 減った列を増やす
                    Funcs.Range(sheet, row, lastCol + 1, rows, -diffCol).Insert(xlShiftToRight, xlFormatFromRightOrBelow);
                }
                // 最終的な範囲を選択する
                Funcs.Range(sheet, row, col, rows, cols + diffCol).Select();
            }
            // 計算用のシートを警告なしに削除
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            tmpSh.Delete();
            Globals.ThisAddIn.Application.DisplayAlerts = true;
        }

        /// <summary>
        /// 結合していると幅を自動調整できないので解除して文字列を分割してセットする
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rows"></param>
        /// <param name="cols"></param>
        /// <returns></returns>
        private static Excel.Range Unmerge(Excel.Worksheet sheet, int rows, int cols)
        {
            Excel.Range range;
            // 結合解除するとenumeratorが不整合になるので毎回やりなおす
            while (true)
            {
                // 継続用フラグ
                bool continueFlag = false;
                range = Funcs.Range(sheet, 1, 1, rows, cols);
                System.Collections.IEnumerator enumerator = range.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    Excel.Range cell = (Excel.Range)enumerator.Current;
                    if ((bool)cell.MergeCells)
                    {
                        // 結合セルが見つかったので継続
                        continueFlag = true;
                        int mergeCols = cell.MergeArea.Columns.Count;
                        // 複数行の場合は最も長い行を使用
                        string maxStr = ((string)cell.Text).Split()
                            .OrderByDescending(line => string.IsNullOrEmpty(line) ? 0 : enc.GetByteCount(line))
                            .FirstOrDefault();
                        cell.UnMerge();
                        string[] split = mbStrSplit(maxStr, mergeCols);
                        for (int i = 0; i < split.Length; i++)
                        {
                            string item = split[i];
                            ((Excel.Range)cell.MergeArea.Columns[i + 1]).Value2 = item;
                        }
                        break; // continue outer while
                    }
                }
                // 結合セルが見つからなかったので終了
                if (!continueFlag)
                {
                    break;
                }
            }

            return range;
        }

        /// <summary>
        /// 文字列を指定した文字数単位で指定の行数に分割する(全角文字考慮)
        /// <seealso href="http://nanoappli.com/blog/archives/1896"/>
        /// </summary>
        /// <param name="inStr">分割前文字列</param>
        /// <param name="count">行数</param>
        /// <returns>分割後文字列の配列</returns>
        private static string[] mbStrSplit(string inStr, int count)
        {
            List<string> outArray = new List<string>(); // 分割結果の保存領域
            string outStr = ""; // 現在処理中の分割後文字列

            // パラメータチェック
            if (inStr == null || count < 1)
            {
                return outArray.ToArray();
            }

            int length = Convert.ToInt32(Math.Ceiling((double)enc.GetByteCount(inStr) / count));

            //--------------------------------------
            // 全ての文字を処理するまで繰り返し
            //--------------------------------------
            for (int offset = 0; offset < inStr.Length; offset++)
            {
                //----------------------------------------------------------
                // 今回処理する文字と、その文字を含めた分割後文字列長を取得
                //----------------------------------------------------------
                string curStr = inStr[offset].ToString();
                int curTotalLength = enc.GetByteCount(outStr) + enc.GetByteCount(curStr);

                //-------------------------------------
                // この文字が、分割点になるかチェック
                //-------------------------------------
                if (curTotalLength == length)
                {
                    // 処理中の文字を含めると、ちょうどピッタリ
                    outArray.Add(outStr + curStr);
                    outStr = "";
                }
                else if (curTotalLength > length)
                {
                    // 処理中の文字を含めると、あふれる
                    outArray.Add(outStr);
                    outStr = curStr;
                }
                else
                {
                    // 処理中の文字を含めてもまだ余裕あり
                    outStr += curStr;
                }
            }

            // 最後の行の文を追加する
            if (!outStr.Equals(""))
            {
                if (outArray.Count == count)
                {
                    outArray[outArray.Count - 1] += outStr;
                }
                else
                {
                    outArray.Add(outStr);
                }
            }

            // 分割後データを配列に変換して返す
            return outArray.ToArray();
        }
    }
}
