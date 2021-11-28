using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SscExcelAddIn
{
    /// <summary>
    /// セル情報リスト要素
    /// </summary>
    public class CellContentModel
    {
        /// <summary>値</summary>
        public string Value { get; set; }
        /// <summary>アドレス</summary>
        public string Address { get; set; }
    }
}
