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

        // AppliedTo は無視する
        private readonly BorderModel borderTop;
        private readonly BorderModel borderRight;
        private readonly BorderModel borderBottom;
        private readonly BorderModel borderLeft;
        private readonly int? dateOperator;
        private readonly FontModel font;
        private readonly string formula1;
        private readonly string formula2;
        private readonly InteriorModel interir;
        private readonly string numberFormat;
        private readonly int? @operator;
        // Priority は無視する
        private readonly bool? ptCondition;
        private readonly int? scoptType;
        private readonly bool? stopIfTrue;
        private readonly string text;
        private readonly int? textOperator;
        private readonly int? type_;

        /// <summary>
        /// ctor
        /// </summary>
        /// <param name="fc"></param>
        public FormatConditionModel(Excel.FormatCondition fc)
        {
            FormatCondition = fc;
            borderTop = new BorderModel(fc.Borders[Excel.XlBordersIndex.xlEdgeTop]);
            borderRight = new BorderModel(fc.Borders[Excel.XlBordersIndex.xlEdgeRight]);
            borderBottom = new BorderModel(fc.Borders[Excel.XlBordersIndex.xlEdgeBottom]);
            borderLeft = new BorderModel(fc.Borders[Excel.XlBordersIndex.xlEdgeLeft]);
            dateOperator = Funcs.OrDefault(fc, e => (int)e.DateOperator);
            font = new FontModel(fc.Font);
            formula1 = Funcs.OrDefault(fc, e => e.Formula1);
            formula2 = Funcs.OrDefault(fc, e => e.Formula2);
            interir = new InteriorModel(fc.Interior);
            numberFormat = Funcs.OrDefault(fc, e => (string)e.NumberFormat);
            @operator = Funcs.OrDefault(fc, e => e.Operator);
            ptCondition = Funcs.OrDefault(fc, e => e.PTCondition);
            scoptType = Funcs.OrDefault(fc, e => (int)e.ScopeType);
            stopIfTrue = Funcs.OrDefault(fc, e => e.StopIfTrue);
            text = Funcs.OrDefault(fc, e => e.Text);
            textOperator = Funcs.OrDefault(fc, e => (int)e.TextOperator);
            type_ = Funcs.OrDefault(fc, e => e.Type);
        }

        ///<inheritdoc/>
        public bool Equals(FormatConditionModel other)
        {
            return other is FormatConditionModel model &&
                   EqualityComparer<BorderModel>.Default.Equals(borderTop, model.borderTop) &&
                   EqualityComparer<BorderModel>.Default.Equals(borderRight, model.borderRight) &&
                   EqualityComparer<BorderModel>.Default.Equals(borderBottom, model.borderBottom) &&
                   EqualityComparer<BorderModel>.Default.Equals(borderLeft, model.borderLeft) &&
                   dateOperator == model.dateOperator &&
                   EqualityComparer<FontModel>.Default.Equals(font, model.font) &&
                   formula1 == model.formula1 &&
                   formula2 == model.formula2 &&
                   EqualityComparer<InteriorModel>.Default.Equals(interir, model.interir) &&
                   numberFormat == model.numberFormat &&
                   @operator == model.@operator &&
                   ptCondition == model.ptCondition &&
                   scoptType == model.scoptType &&
                   stopIfTrue == model.stopIfTrue &&
                   text == model.text &&
                   textOperator == model.textOperator &&
                   type_ == model.type_;
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
            hashCode = (hashCode * -1521134295) + EqualityComparer<BorderModel>.Default.GetHashCode(borderTop);
            hashCode = (hashCode * -1521134295) + EqualityComparer<BorderModel>.Default.GetHashCode(borderRight);
            hashCode = (hashCode * -1521134295) + EqualityComparer<BorderModel>.Default.GetHashCode(borderBottom);
            hashCode = (hashCode * -1521134295) + EqualityComparer<BorderModel>.Default.GetHashCode(borderLeft);
            hashCode = (hashCode * -1521134295) + dateOperator.GetHashCode();
            hashCode = (hashCode * -1521134295) + EqualityComparer<FontModel>.Default.GetHashCode(font);
            hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(formula1);
            hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(formula2);
            hashCode = (hashCode * -1521134295) + EqualityComparer<InteriorModel>.Default.GetHashCode(interir);
            hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(numberFormat);
            hashCode = (hashCode * -1521134295) + @operator.GetHashCode();
            hashCode = (hashCode * -1521134295) + ptCondition.GetHashCode();
            hashCode = (hashCode * -1521134295) + scoptType.GetHashCode();
            hashCode = (hashCode * -1521134295) + stopIfTrue.GetHashCode();
            hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(text);
            hashCode = (hashCode * -1521134295) + textOperator.GetHashCode();
            hashCode = (hashCode * -1521134295) + type_.GetHashCode();
            return hashCode;
        }
    }
}
