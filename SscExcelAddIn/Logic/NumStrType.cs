namespace SscExcelAddIn.Logic
{
    /// <summary>
    /// 番号を表す文字列の種類
    /// </summary>
    public enum NumStrType
    {
        /// <summary>未定義</summary>
        U,
        /// <summary>半角数字</summary>
        NN,
        /// <summary>全角数字</summary>
        NW,
        /// <summary>丸囲み数字</summary>
        M,
        /// <summary>大文字ローマ数字</summary>
        RU,
        /// <summary>英小文字</summary>
        ALN,
        /// <summary>英大文字</summary>
        AUN,
        /// <summary>全角カタカナ</summary>
        KW,
        /// <summary>半角カタカナ</summary>
        KN
    }

}
