using System;
using System.Collections.Generic;
using SscExcelAddIn.Logic;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.ComModel
{
    /// <summary>
    /// <see cref="Excel.Font"/> のモデル
    /// </summary>
    internal class FontModel : IEquatable<FontModel>
    {
        /// <summary>Excelオブジェクト</summary>
        public Excel.Font Font { get; }

        private readonly int? background;
        private readonly bool? bold;
        private readonly double? color;
        private readonly int? colorIndex;
        private readonly bool? italic;
        private readonly object name;
        private readonly int? size;
        private readonly bool? strikethrough;
        private readonly bool? subscript;
        private readonly bool? superscript;
        private readonly int? themeColor;
        private readonly int? themeFont;
        private readonly double? tintAndShade;
        private readonly bool? underline;

        /// <summary>
        /// ctor
        /// </summary>
        /// <param name="font"></param>
        public FontModel(Excel.Font font)
        {
            Font = font;
            background = Funcs.OrDefault(font, e => (int)e.Background);
            bold = Funcs.OrDefault(font, e => (bool)e.Bold);
            color = Funcs.OrDefault(font, e => (double)e.Color);
            colorIndex = Funcs.OrDefault(font, e => (int)e.ColorIndex);
            // font.FontStyle is dependent
            italic = Funcs.OrDefault(font, e => (bool)e.Italic);
            name = Funcs.OrDefault(font, e => (object)e.Name);
            size = Funcs.OrDefault(font, e => (int)e.Size);
            strikethrough = Funcs.OrDefault(font, e => (bool)e.Strikethrough);
            subscript = Funcs.OrDefault(font, e => (bool)e.Subscript);
            superscript = Funcs.OrDefault(font, e => (bool)e.Superscript);
            themeColor = Funcs.OrDefault(font, e => (int)e.ThemeColor);
            themeFont = Funcs.OrDefault(font, e => (int)e.ThemeFont);
            tintAndShade = Funcs.OrDefault(font, e => (double)e.TintAndShade);
            underline = Funcs.OrDefault(font, e => (bool)e.Underline);
        }

        ///<inheritdoc/>
        public bool Equals(FontModel other)
        {
            return other is FontModel model &&
                   background == model.background &&
                   bold == model.bold &&
                   color == model.color &&
                   colorIndex == model.colorIndex &&
                   italic == model.italic &&
                   EqualityComparer<object>.Default.Equals(name, model.name) &&
                   size == model.size &&
                   strikethrough == model.strikethrough &&
                   subscript == model.subscript &&
                   superscript == model.superscript &&
                   themeColor == model.themeColor &&
                   themeFont == model.themeFont &&
                   tintAndShade == model.tintAndShade &&
                   underline == model.underline;
        }

        ///<inheritdoc/>
        public override bool Equals(object obj)
        {
            return obj is FontModel model ? Equals(model) : base.Equals(obj);

        }

        ///<inheritdoc/>
        public override int GetHashCode()
        {
            int hashCode = 945841388;
            hashCode = hashCode * -1521134295 + background.GetHashCode();
            hashCode = hashCode * -1521134295 + bold.GetHashCode();
            hashCode = hashCode * -1521134295 + color.GetHashCode();
            hashCode = hashCode * -1521134295 + colorIndex.GetHashCode();
            hashCode = hashCode * -1521134295 + italic.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<object>.Default.GetHashCode(name);
            hashCode = hashCode * -1521134295 + size.GetHashCode();
            hashCode = hashCode * -1521134295 + strikethrough.GetHashCode();
            hashCode = hashCode * -1521134295 + subscript.GetHashCode();
            hashCode = hashCode * -1521134295 + superscript.GetHashCode();
            hashCode = hashCode * -1521134295 + themeColor.GetHashCode();
            hashCode = hashCode * -1521134295 + themeFont.GetHashCode();
            hashCode = hashCode * -1521134295 + tintAndShade.GetHashCode();
            hashCode = hashCode * -1521134295 + underline.GetHashCode();
            return hashCode;
        }
    }
}
