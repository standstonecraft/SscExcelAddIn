using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace SscExcelAddIn.ValConv
{
    /// <summary>
    /// 文字列を空でないかどうかの真偽値に変換する
    /// </summary>
    public class TextNotEmptyConverter : IValueConverter
    {
        /// <summary>
        /// 文字列を空でないかどうかの真偽値に変換する
        /// </summary>
        /// <param name="value">文字列</param>
        /// <param name="targetType"></param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns>空でないかどうか</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value != null && value.ToString().Length > 0;
        }

        /// <summary>
        /// 実装しない
        /// </summary>
        /// <param name="value"></param>
        /// <param name="targetType"></param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
