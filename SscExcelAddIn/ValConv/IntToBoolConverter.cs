using System;
using System.Globalization;
using System.Windows.Data;

namespace SscExcelAddIn.ValConv
{
    /// <summary>
    /// int から bool に変換するコンバーター
    /// <see href="https://qiita.com/tera1707/items/47d1f1766cbe798b0c13"/>
    /// </summary>
    public class IntToBoolConverter : IValueConverter
    {
        /// <summary>
        /// 値を変換する。
        /// </summary>
        /// <param name="value">int</param>
        /// <param name="targetType"></param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns>value が0以上かどうか</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            int val = (int)value;
            int threshold = -1;
            if (parameter != null)
            {
                string para = parameter.ToString();
                threshold = int.Parse(para);
            }
            return val > threshold;
        }

        /// <summary>
        /// 未実装
        /// </summary>
        /// <param name="value"></param>
        /// <param name="targetType"></param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            //string strValue = value as string;
            //DateTime resultDateTime;
            //if (DateTime.TryParse(strValue, out resultDateTime))
            //{
            //    return resultDateTime;
            //}
            //return DependencyProperty.UnsetValue;
            throw new NotImplementedException();
        }
    }
}
