using System.Windows.Controls;
using System.Windows.Input;

namespace SscExcelAddIn
{
    /// <summary>
    /// TestControl.xaml の相互作用ロジック
    /// </summary>
    public partial class TestControl : UserControl
    {
        public TestControl()
        {
            InitializeComponent();
        }

        private void ImeTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            ImeLabel.Content = (ImeTextBox.IsDirectInput(e) ? "Off" : "On") + " " + e.Key.ToString();
        }

        private void ImeTextBox_EnterKeyUp(object sender, KeyEventArgs e)
        {
            ImeLabel.Content += "WOW";
        }
    }
}
