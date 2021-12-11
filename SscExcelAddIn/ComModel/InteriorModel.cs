using System;
using System.Collections.Generic;
using SscExcelAddIn.Logic;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.ComModel
{
    /// <summary>
    /// <see cref="Excel.Interior"/> のモデル
    /// </summary>
    internal class InteriorModel : IEquatable<InteriorModel>
    {
        /// <summary>Excelオブジェクト</summary>
        public Excel.Interior Interior { get; }

        private readonly double? color;
        private readonly int? colorIndex;
        private readonly GradientModel gradient;
        private readonly bool? invertIfNegative;
        private readonly int? pattern;
        private readonly double? patternColor;
        private readonly int? patternColorIndex;
        private readonly int? patternThemeColor;
        private readonly double? patternTintAndShade;
        private readonly int? themeColor;
        private readonly double? tintAndShade;

        /// <summary>
        /// ctor
        /// </summary>
        /// <param name="interior"></param>
        public InteriorModel(Excel.Interior interior)
        {
            Interior = interior;
            color = Funcs.OrDefault(interior, e => (double)e.Color);
            colorIndex = Funcs.OrDefault(interior, e => (int)e.ColorIndex);
            gradient = new GradientModel(interior.Gradient);
            invertIfNegative = Funcs.OrDefault(interior, e => (bool)e.InvertIfNegative);
            pattern = Funcs.OrDefault(interior, e => (int)e.Pattern);
            patternColor = Funcs.OrDefault(interior, e => (double)e.PatternColor);
            patternColorIndex = Funcs.OrDefault(interior, e => (int)e.PatternColorIndex);
            patternThemeColor = Funcs.OrDefault(interior, e => (int)e.PatternThemeColor);
            patternTintAndShade = Funcs.OrDefault(interior, e => (double)e.PatternTintAndShade);
            themeColor = Funcs.OrDefault(interior, e => (int)e.ThemeColor);
            tintAndShade = Funcs.OrDefault(interior, e => (double)e.TintAndShade);
        }

        ///<inheritdoc/>
        public bool Equals(InteriorModel other)
        {
            return other is InteriorModel model &&
                   color == model.color &&
                   colorIndex == model.colorIndex &&
                   EqualityComparer<GradientModel>.Default.Equals(gradient, model.gradient) &&
                   invertIfNegative == model.invertIfNegative &&
                   pattern == model.pattern &&
                   patternColor == model.patternColor &&
                   patternColorIndex == model.patternColorIndex &&
                   patternThemeColor == model.patternThemeColor &&
                   patternTintAndShade == model.patternTintAndShade &&
                   themeColor == model.themeColor &&
                   tintAndShade == model.tintAndShade;
        }

        ///<inheritdoc/>
        public override bool Equals(object obj)
        {
            return obj is InteriorModel model ? Equals(model) : base.Equals(obj);
        }

        ///<inheritdoc/>
        public override int GetHashCode()
        {
            int hashCode = -1807703255;
            hashCode = (hashCode * -1521134295) + color.GetHashCode();
            hashCode = (hashCode * -1521134295) + colorIndex.GetHashCode();
            hashCode = (hashCode * -1521134295) + EqualityComparer<GradientModel>.Default.GetHashCode(gradient);
            hashCode = (hashCode * -1521134295) + invertIfNegative.GetHashCode();
            hashCode = (hashCode * -1521134295) + pattern.GetHashCode();
            hashCode = (hashCode * -1521134295) + patternColor.GetHashCode();
            hashCode = (hashCode * -1521134295) + patternColorIndex.GetHashCode();
            hashCode = (hashCode * -1521134295) + patternThemeColor.GetHashCode();
            hashCode = (hashCode * -1521134295) + patternTintAndShade.GetHashCode();
            hashCode = (hashCode * -1521134295) + themeColor.GetHashCode();
            hashCode = (hashCode * -1521134295) + tintAndShade.GetHashCode();
            return hashCode;
        }
    }
}
