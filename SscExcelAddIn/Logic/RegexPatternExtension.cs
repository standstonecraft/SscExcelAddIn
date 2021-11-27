using System;
using System.Text.RegularExpressions;

namespace SscExcelAddIn.Logic
{
    /// <summary>
    /// <see cref="RegexPattern"/> 用の拡張メソッド
    /// </summary>
    public static class RegexPatternExtension
    {
        /// <summary>
        /// <see cref="RegexPattern"/> の名前を検索文字列としてパターンに置換する
        /// </summary>
        /// <param name="str">文字列</param>
        /// <param name="rp">RegexPattern</param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public static string Replace(this string str, RegexPattern rp, TimeSpan? timeout = null)
        {
            if (timeout.HasValue)
            {
                return Regex.Replace(str, rp.Key, rp.Pattern, RegexOptions.None, timeout.Value);
            }
            else
            {
                return str.Replace(rp.Key, rp.Pattern);
            }
        }
    }

}
