using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace SscExcelAddIn.Logic
{
    /// <summary>
    /// 正規表現パターンに名前をつけて管理する
    /// </summary>
    public partial class RegexPattern
    {
        /// <summary>
        /// パターン名
        /// </summary>
        public string Key { get; set; }
        /// <summary>
        /// パターン
        /// </summary>
        public string Pattern { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="key">パターン名</param>
        /// <param name="pattern">パターン</param>
        public RegexPattern(string key, string pattern)
        {
            Key = key;
            Pattern = pattern;
            InnerPatterns.Add(this);
        }

        private static readonly List<RegexPattern> InnerPatterns = new List<RegexPattern>();

        /// <summary>
        /// パターンのリスト
        /// </summary>
        public static ReadOnlyCollection<RegexPattern> Patterns => InnerPatterns.AsReadOnly();

        /// <summary>
        /// 数字を表す文字列
        /// </summary>
        public static readonly RegexPattern NUM = new RegexPattern("_{NUM}_",
            "[-－]?[0-9０-９]+|[Ⅰ-Ⅻ①-⑳㉑-㉟㊱-㊿]");
        /// <summary>
        /// 数字と序列をを表す文字列
        /// </summary>
        public static readonly RegexPattern ALL = new RegexPattern("_{ALL}_",
            "[-－]?[0-9０-９]+|[Ⅰ-Ⅻ①-⑳㉑-㉟㊱-㊿a-zA-Z" +
            "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨララリルレロワヲン" +
            "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾗﾘﾙﾚﾛﾜｦﾝ]");
        /// <summary>
        /// 丸囲み数字を表す文字列
        /// </summary>
        public static readonly RegexPattern MARU = new RegexPattern("_{MARU}_",
            "[①-⑳㉑-㉟㊱-㊿]");
        /// <summary>
        /// 全角文字を表す文字列
        /// </summary>
        public static readonly RegexPattern ZEN = new RegexPattern("_{ZEN}_",
            @"[^\x01-\x7E\xA1-\xDF]+");
    }
}
