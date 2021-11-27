using System.Windows;
using System.Windows.Controls;

namespace SscExcelAddIn
{
    /// <summary>
    /// WebControl.xaml の相互作用ロジック
    /// </summary>
    public partial class WebControl : UserControl
    {
        /// <summary>
        /// WebViewと閉じるボタンを備えるコントロール
        /// </summary>
        /// <param name="navigateString"></param>
        public WebControl(string navigateString)
        {
            InitializeComponent();
            this.TheWevBiew.NavigateToString(navigateString);
            this.TheWevBiew.Navigating += (sender, e) =>
            {
                // https://stackoverflow.com/questions/21255643/how-to-open-links-in-wpf-webview-in-default-explorer/21255951#21255951
                e.Cancel = true;
                System.Diagnostics.Process.Start(e.Uri.ToString());
            };
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
