using System;
using System.Collections.Generic;
using SscExcelAddIn.Logic;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.ComModel
{
    /// <summary>
    /// <see cref="Excel.FormatCondition"/> のモデル
    /// </summary>
    internal class FormatConditionModel : IEquatable<FormatConditionModel>
    {
        /// <summary>Excelオブジェクト</summary>
        public readonly Excel.FormatCondition FormatCondition;

        private readonly Excel.XlFormatConditionType? type;
        private readonly Excel.XlFormatConditionOperator? @operator;

        // AppliedTo は無視する
        private readonly BorderModel borderTop;
        private readonly BorderModel borderRight;
        private readonly BorderModel borderBottom;
        private readonly BorderModel borderLeft;
        private readonly int? dateOperator;
        private readonly FontModel font;
        private readonly string formula1;
        private readonly string formula2;
        private readonly InteriorModel interior;
        private readonly string numberFormat;
        // Priority は無視する
        private readonly bool? ptCondition;
        private readonly int? scopeType;
        private readonly bool? stopIfTrue;
        private readonly string text;
        private readonly int? textOperator;

        /// <summary>
        /// ctor
        /// </summary>
        /// <param name="fc"></param>
        public FormatConditionModel(Excel.FormatCondition fc)
        {
            FormatCondition = fc;
            type = (Excel.XlFormatConditionType)fc.Type;
            if (type != Excel.XlFormatConditionType.xlExpression)
            {
                @operator = (Excel.XlFormatConditionOperator)fc.Operator;
            }

            borderTop = new BorderModel(fc.Borders[Excel.XlBordersIndex.xlEdgeTop]);
            borderRight = new BorderModel(fc.Borders[Excel.XlBordersIndex.xlEdgeRight]);
            borderBottom = new BorderModel(fc.Borders[Excel.XlBordersIndex.xlEdgeBottom]);
            borderLeft = new BorderModel(fc.Borders[Excel.XlBordersIndex.xlEdgeLeft]);
            dateOperator = Funcs.OrDefault(fc, e => (int)e.DateOperator);
            font = new FontModel(fc.Font);
            formula1 = (string)Globals.ThisAddIn.Application.ConvertFormula(fc.Formula1,
                Excel.XlReferenceStyle.xlA1, Excel.XlReferenceStyle.xlR1C1,
                RelativeTo: fc.AppliesTo[1, 1]);
            if (@operator == Excel.XlFormatConditionOperator.xlBetween
                || @operator == Excel.XlFormatConditionOperator.xlNotBetween)
            {
                formula2 = fc.Formula2;
            }
            interior = new InteriorModel(fc.Interior);
            numberFormat = Funcs.OrDefault(fc, e => (string)e.NumberFormat);
            ptCondition = Funcs.OrDefault(fc, e => e.PTCondition);
            scopeType = Funcs.OrDefault(fc, e => (int)e.ScopeType);
            stopIfTrue = Funcs.OrDefault(fc, e => e.StopIfTrue);
            text = Funcs.OrDefault(fc, e => e.Text);
            textOperator = Funcs.OrDefault(fc, e => (int)e.TextOperator);
        }

        ///<inheritdoc/>
        public bool Equals(FormatConditionModel other)
        {
            return other is FormatConditionModel model &&
                type == model.type &&
                @operator == model.@operator &&
                EqualityComparer<BorderModel>.Default.Equals(borderTop, model.borderTop) &&
                EqualityComparer<BorderModel>.Default.Equals(borderRight, model.borderRight) &&
                EqualityComparer<BorderModel>.Default.Equals(borderBottom, model.borderBottom) &&
                EqualityComparer<BorderModel>.Default.Equals(borderLeft, model.borderLeft) &&
                dateOperator == model.dateOperator &&
                EqualityComparer<FontModel>.Default.Equals(font, model.font) &&
                formula1 == model.formula1 &&
                formula2 == model.formula2 &&
                EqualityComparer<InteriorModel>.Default.Equals(interior, model.interior) &&
                numberFormat == model.numberFormat &&
                ptCondition == model.ptCondition &&
                scopeType == model.scopeType &&
                stopIfTrue == model.stopIfTrue &&
                text == model.text &&
                textOperator == model.textOperator;
        }

        ///<inheritdoc/>
        public override bool Equals(object obj)
        {
            return obj is FormatConditionModel model ? Equals(model) : base.Equals(obj);
        }

        ///<inheritdoc/>
        public override int GetHashCode()
        {
            int hashCode = 933748780;
            hashCode = hashCode * -1521134295 + type.GetHashCode();
            hashCode = hashCode * -1521134295 + @operator.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<BorderModel>.Default.GetHashCode(borderTop);
            hashCode = hashCode * -1521134295 + EqualityComparer<BorderModel>.Default.GetHashCode(borderRight);
            hashCode = hashCode * -1521134295 + EqualityComparer<BorderModel>.Default.GetHashCode(borderBottom);
            hashCode = hashCode * -1521134295 + EqualityComparer<BorderModel>.Default.GetHashCode(borderLeft);
            hashCode = hashCode * -1521134295 + dateOperator.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<FontModel>.Default.GetHashCode(font);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(formula1);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(formula2);
            hashCode = hashCode * -1521134295 + EqualityComparer<InteriorModel>.Default.GetHashCode(interior);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(numberFormat);
            hashCode = hashCode * -1521134295 + ptCondition.GetHashCode();
            hashCode = hashCode * -1521134295 + scopeType.GetHashCode();
            hashCode = hashCode * -1521134295 + stopIfTrue.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(text);
            hashCode = hashCode * -1521134295 + textOperator.GetHashCode();
            return hashCode;
        }
    }
}
