using System;
using SscExcelAddIn.Logic;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.ComModel
{
    /// <summary>
    /// <see cref="Excel.Border"/> のモデル
    /// </summary>
    internal class BorderModel : IEquatable<BorderModel>
    {
        /// <summary>Excelオブジェクト</summary>
        public Excel.Border Border { get; }

        private readonly double? color;
        private readonly int? colorIndex;
        private readonly int? lineStyle;
        private readonly int? themeColor;
        private readonly double? tintAndShade;
        private readonly int? weight;

        /// <summary>
        /// ctor
        /// </summary>
        /// <param name="border"></param>
        public BorderModel(Excel.Border border)
        {
            Border = border;
            color = Funcs.OrDefault(border, e => (double)e.Color);
            colorIndex = Funcs.OrDefault(border, e => (int)e.ColorIndex);
            lineStyle = Funcs.OrDefault(border, e => (int)e.LineStyle);
            themeColor = Funcs.OrDefault(border, e => (int)e.ThemeColor);
            tintAndShade = Funcs.OrDefault(border, e => (double)e.TintAndShade);
            weight = Funcs.OrDefault(border, e => (int)e.Weight);
        }

        ///<inheritdoc/>
        public bool Equals(BorderModel other)
        {
            return other is BorderModel model &&
                   color == model.color &&
                   colorIndex == model.colorIndex &&
                   lineStyle == model.lineStyle &&
                   themeColor == model.themeColor &&
                   tintAndShade == model.tintAndShade &&
                   weight == model.weight;
        }

        ///<inheritdoc/>
        public override bool Equals(object obj)
        {
            return obj is BorderModel model ? Equals(model) : base.Equals(obj);
        }

        ///<inheritdoc/>
        public override int GetHashCode()
        {
            int hashCode = 1261321268;
            hashCode = (hashCode * -1521134295) + color.GetHashCode();
            hashCode = (hashCode * -1521134295) + colorIndex.GetHashCode();
            hashCode = (hashCode * -1521134295) + lineStyle.GetHashCode();
            hashCode = (hashCode * -1521134295) + themeColor.GetHashCode();
            hashCode = (hashCode * -1521134295) + tintAndShade.GetHashCode();
            hashCode = (hashCode * -1521134295) + weight.GetHashCode();
            return hashCode;
        }
    }
}
