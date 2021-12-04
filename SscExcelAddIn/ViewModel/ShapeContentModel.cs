using System.Collections.Generic;

namespace SscExcelAddIn
{
    /// <summary>
    /// シェイプ情報リスト要素
    /// </summary>
    public class ShapeContentModel
    {
        /// <summary>値</summary>
        public string Value { get; set; }
        /// <summary>数式</summary>
        public string Formula { get; set; }
        /// <summary>アドレス</summary>
        public string Address { get; set; }
        /// <summary>x座標</summary>
        public double Left { get; set; }
        /// <summary>y座標</summary>
        public double Top { get; set; }
        /// <summary>Range</summary>
        public dynamic Range { get; internal set; }
        /// <summary>行</summary>
        public int Row { get; internal set; }
        /// <summary>列</summary>
        public int Column { get; internal set; }

        /// <summary>
        /// ctor
        /// </summary>
        /// <param name="range"><see cref="Range"/></param>
        public ShapeContentModel(dynamic range)
        {
            Value = range.DrawingObject.Text;
            Formula = range.DrawingObject.Formula?.Trim();
            Address = range.TopLeftCell.Address;
            Top = range.Top;
            Left = range.Left;
            Range = range;
            Row = range.TopLeftCell.Row;
            Column = range.TopLeftCell.Column;
        }

        /// <summary>
        /// 行位置優先の位置コンパレーター
        /// </summary>
        /// <returns></returns>
        public static IComparer<ShapeContentModel> RowColComparer =
            Comparer<ShapeContentModel>.Create((a, b) =>
            {
                double diff = a.Top - b.Top;
                if (diff != 0)
                {
                    return (int)(diff * 1000);
                }
                diff = a.Left - b.Left;
                return (int)(diff * 1000);
            });

        /// <summary>
        /// 列位置優先の位置コンパレーター
        /// </summary>
        /// <returns></returns>
        public static IComparer<ShapeContentModel> ColRowComparer =
            Comparer<ShapeContentModel>.Create((a, b) =>
            {
                double diff = a.Left - b.Left;
                if (diff != 0)
                {
                    return (int)(diff * 1000);
                }
                diff = a.Top - b.Top;
                return (int)(diff * 1000);
            });

        /// <inheritdoc/>
        public override string ToString()
        {
            return Value;
        }
    }
}
