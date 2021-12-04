using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using SystemColors = System.Windows.SystemColors;

namespace SscExcelAddIn.Control
{
    /// <summary>
    /// プレースホルダーを表示可能なテキストボックス
    /// </summary>
    public class PlaceHolderTextBox : TextBox
    {

        private bool isPlaceHolder = true;
        private string _placeHolderText;

        /// <summary>
        /// プレースホルダーテキスト
        /// </summary>
        public string PlaceHolderText
        {
            get => _placeHolderText;
            set
            {
                _placeHolderText = value;
                setPlaceholder();
            }
        }

        /// <summary>
        /// テキスト
        /// </summary>
        public new string Text
        {
            get => isPlaceHolder ? string.Empty : base.Text;
            set => base.Text = value;
        }

        /// <summary>
        /// when the control loses focus, the placeholder is shown
        /// </summary>
        private void setPlaceholder()
        {
            if (string.IsNullOrEmpty(base.Text))
            {
                base.Text = PlaceHolderText;
                Foreground = new SolidColorBrush(Colors.Gray);
                FontStyle = FontStyles.Italic;
                isPlaceHolder = true;
            }
        }

        /// <summary>
        /// when the control is focused, the placeholder is removed
        /// </summary>
        private void removePlaceHolder()
        {

            if (isPlaceHolder)
            {
                base.Text = "";
                Foreground = SystemColors.WindowTextBrush;
                FontStyle = FontStyles.Normal;
                isPlaceHolder = false;
            }
        }
        /// <summary>
        /// ctor
        /// </summary>
        public PlaceHolderTextBox()
        {
            GotFocus += removePlaceHolder;
            LostFocus += setPlaceholder;
        }

        private void setPlaceholder(object sender, EventArgs e)
        {
            setPlaceholder();
        }

        private void removePlaceHolder(object sender, EventArgs e)
        {
            removePlaceHolder();
        }
    }
}
