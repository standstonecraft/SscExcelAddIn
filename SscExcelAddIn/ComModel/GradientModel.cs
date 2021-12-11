using System;
using SscExcelAddIn.Logic;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.ComModel
{
    /// <summary>
    /// <see cref="Excel.Interior"/> のグラデーションのモデル
    /// 実体はLinearGradient または RectangularGradient のいずれかであるが
    /// 区別の必要がないためdynamicのままで扱い、両方のプロパティを保持しようとする。
    /// </summary>
    internal class GradientModel : IEquatable<GradientModel>
    {
        /// <summary>Excelオブジェクト</summary>
        public dynamic Gradient { get; }

        /// <summary>Both</summary>
        private readonly int? colorStops;
        /// <summary>LinearGradient</summary>
        private readonly double? degree;
        /// <summary>RectangularGradient</summary>
        private readonly double? rectangleBottom;
        /// <summary>RectangularGradient</summary>
        private readonly double? rectangleLeft;
        /// <summary>RectangularGradient</summary>
        private readonly double? rectangleRight;
        /// <summary>RectangularGradient</summary>
        private readonly double? rectangleTop;

        /// <summary>
        /// ctor
        /// </summary>
        /// <param name="grad"></param>
        public GradientModel(dynamic grad)
        {
            Gradient = grad;
            colorStops = Funcs.OrDefault((object)grad, e => (int)((dynamic)e).ColorStops.Count);
            degree = Funcs.OrDefault((object)grad, e => (double)((dynamic)e).Degree);
            rectangleBottom = Funcs.OrDefault((object)grad, e => (double)((dynamic)e).RectangleBottom);
            rectangleLeft = Funcs.OrDefault((object)grad, e => (double)((dynamic)e).RectangleLeft);
            rectangleRight = Funcs.OrDefault((object)grad, e => (double)((dynamic)e).RectangleRight);
            rectangleTop = Funcs.OrDefault((object)grad, e => (double)((dynamic)e).RectangleTop);
        }

        ///<inheritdoc/>
        public bool Equals(GradientModel other)
        {
            return other is GradientModel model &&
                   colorStops == model.colorStops &&
                   degree == model.degree &&
                   rectangleBottom == model.rectangleBottom &&
                   rectangleLeft == model.rectangleLeft &&
                   rectangleRight == model.rectangleRight &&
                   rectangleTop == model.rectangleTop;
        }

        ///<inheritdoc/>
        public override bool Equals(object obj)
        {
            return obj is GradientModel model ? Equals(model) : base.Equals(obj);
        }

        ///<inheritdoc/>
        public override int GetHashCode()
        {
            int hashCode = 1602685743;
            hashCode = (hashCode * -1521134295) + colorStops.GetHashCode();
            hashCode = (hashCode * -1521134295) + degree.GetHashCode();
            hashCode = (hashCode * -1521134295) + rectangleBottom.GetHashCode();
            hashCode = (hashCode * -1521134295) + rectangleLeft.GetHashCode();
            hashCode = (hashCode * -1521134295) + rectangleRight.GetHashCode();
            hashCode = (hashCode * -1521134295) + rectangleTop.GetHashCode();
            return hashCode;
        }
    }
}
