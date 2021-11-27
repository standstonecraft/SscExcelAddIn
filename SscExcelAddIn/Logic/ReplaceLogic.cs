using System;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.Logic
{
    /// <summary>
    /// <see cref="ReplaceControl"/> のロジック
    /// </summary>
    public class ReplaceLogic
    {
        private static readonly TimeSpan MatchTimeout = TimeSpan.FromSeconds(2);
        private static readonly string HasFuncP = @"_(INC|SEQ|CAS|NAR)\(";
        private static readonly Regex HasFunc = new Regex(HasFuncP, RegexOptions.Compiled);
        /// <summary>
        /// 入れ子の一番深い関数にマッチさせる。先読みの括弧がグループのインデックスを進めてしまうので注意。
        /// 0がマッチ全体で1が先読みなので、意味のあるグループは2以降になる。
        /// </summary>
        private static readonly string NotFuncP = @"(?!" + HasFuncP + @".+?\))(.+?)";
        private static readonly Regex FuncInc = new Regex(@"_INC\(" + NotFuncP + @",(-?\d+)\)", RegexOptions.Compiled, MatchTimeout);
        private static readonly Regex FuncSeq = new Regex(@"_SEQ\(" + NotFuncP + @",(-?\d+)\)", RegexOptions.Compiled, MatchTimeout);
        private static readonly Regex FuncCas = new Regex(@"_CAS\(" + NotFuncP + @",(\p{Lu}+)\)", RegexOptions.Compiled, MatchTimeout);
        private static readonly Regex FuncNar = new Regex(@"_NAR\(" + NotFuncP + @"_NAR\)", RegexOptions.Compiled, MatchTimeout);

        /// <summary>
        /// 文字列を置換する。通常の正規表現に加え、独自の検索パターンと置換関数を実装する。
        /// </summary>
        /// <param name="input">対象文字列</param>
        /// <param name="pattern">検索文字列</param>
        /// <param name="replacement">置換文字列</param>
        /// <returns>置換後の文字列</returns>
        public static string ReplaceText(string input, string pattern, string replacement)
        {
            int dummy = 0;
            return ReplaceText(input, pattern, replacement, ref dummy);
        }

        /// <summary>
        /// 文字列を置換する。通常の正規表現に加え、独自の検索パターンと置換関数を実装する。
        /// </summary>
        /// <param name="input">対象文字列</param>
        /// <param name="pattern">検索文字列</param>
        /// <param name="replacement">置換文字列</param>
        /// <param name="seq">SEQ関数で置き換えられるシーケンス番号</param>
        /// <returns>置換後の文字列</returns>
        public static string ReplaceText(string input, string pattern, string replacement, ref int seq)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(pattern))
            {
                return input;
            }

            try
            {
                foreach (RegexPattern rp in RegexPattern.Patterns)
                {
                    pattern = pattern.Replace(rp, MatchTimeout);
                }
                // 通常の置換
                string replaced = Regex.Replace(input, pattern, replacement, RegexOptions.None, MatchTimeout);
                // 独自関数
                int loopCnt = 0;
                int seqNum = seq;
                while (HasFunc.IsMatch(replaced) && loopCnt++ < 100)
                {
                    // 関数
                    replaced = FuncInc.Replace(replaced,
                        m => new NumStr(m.Groups[2].Value).Add(m.Groups[3].Value).ToString());
                    replaced = FuncSeq.Replace(replaced,
                        m => new NumStr(m.Groups[2].Value).Set(m.Groups[3].Value).Add(seqNum).ToString());
                    replaced = FuncCas.Replace(replaced,
                        m => new NumStr(m.Groups[2].Value).SetType(m.Groups[3].Value).ToString());
                    replaced = FuncNar.Replace(replaced,
                        m => Strings.StrConv(m.Groups[2].Value, VbStrConv.Narrow));
                }
                if (Regex.IsMatch(input, pattern))
                {
                    seq += 1;
                }
                return replaced;
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="range">置換対象のセル範囲</param>
        /// <param name="patternText">検索文字列</param>
        /// <param name="replacement">置換文字列</param>
        public static void ReplaceTextRange(Excel.Range range, string patternText, string replacement)
        {
            if (range != null)
            {
                int hitCount = 0;
                if (range.Formula is object[,] formula)
                {
                    for (int r = 1; r <= formula.GetLength(0); r++)
                    {
                        for (int c = 1; c <= formula.GetLength(1); c++)
                        {
                            if (formula[r, c] != null)
                            {
                                formula[r, c] = ReplaceText(formula[r, c].ToString(), patternText, replacement, ref hitCount);
                            }
                        }
                    }
                    range.Formula = formula;
                }
                else
                {
                    range.Formula = ReplaceText(range.Formula.ToString(), patternText, replacement, ref hitCount);
                }
            }
        }
    }
}
