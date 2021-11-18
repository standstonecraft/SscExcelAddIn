using Microsoft.VisualBasic;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.Logic
{
    public class RegexLogic
    {
        public static ReadOnlyDictionary<string, string> CharPatterns = new ReadOnlyDictionary<string, string>(
                new Dictionary<string, string>{{"①-⑳", "maruNum"}, {"Ⅰ-Ⅹ", "upperRomeNum"},
                {"a-z", "lower"}, {"A-Z", "upper"}, {"ア-ン", "zenKana"}}
            );

        public static string ReplaceText(string input, string pattern, string replacement)
        {
            // 通常の置換
            string replaced = Regex.Replace(input, pattern, replacement);
            // 半角数値
            replaced = Regex.Replace(replaced, @"_INC\(([0-9]+),(-?\d+)\)",
                m => NumStrConv.AddNum("num", m.Groups[1].Value, m.Groups[2].Value));
            // 全角数値
            replaced = Regex.Replace(replaced, @"_INC\(([０-９]+),(-?\d+)\)",
                m => NumStrConv.AddNum("zenNum", m.Groups[1].Value, m.Groups[2].Value));
            // 文字
            foreach (KeyValuePair<string, string> e in CharPatterns)
            {
                replaced = Regex.Replace(replaced, string.Format(@"_INC\(([{0}]),(-?\d+)\)", e.Key),
                    m => NumStrConv.AddNum(e.Value, m.Groups[1].Value, m.Groups[2].Value));
            }
            // 全角->半角
            replaced = Regex.Replace(replaced, @"_NAR_\((.+?)_NAR_\)",
                m => Strings.StrConv(m.Groups[1].Value, VbStrConv.Narrow));
            return replaced;
        }

        public static void ReplaceTextRange(Excel.Range range, string patternText, string replacement)
        {
            if (range != null)
            {
                if (range.Formula is object[,] formula)
                {
                    for (int r = 1; r <= formula.GetLength(0); r++)
                    {
                        for (int c = 1; c <= formula.GetLength(1); c++)
                        {
                            if (formula[r, c] != null)
                            {
                                formula[r, c] = ReplaceText(formula[r, c].ToString(), patternText, replacement);
                            }
                        }
                    }
                    range.Formula = formula;
                }
                else
                {
                    range.Formula = ReplaceText(range.Formula.ToString(), patternText, replacement);
                }
            }
        }
    }
}
