using Microsoft.Office.Tools.Ribbon;
using System.Windows;

namespace SscExcelAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        /// <summary>
        /// 置換ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ReplaceButton_Click(object sender, RibbonControlEventArgs e)
        {
            ShowReplaceWindow();
        }

        private void AboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            ShowAboutWindow();
        }

        private void SkipSelectButton_Click(object sender, RibbonControlEventArgs e)
        {
            ShowSkipSelectWindow();
        }

        private void TestControlButton_Click(object sender, RibbonControlEventArgs e)
        {
            new Window { Content = new TestControl() }.Show();
        }
    }
}
