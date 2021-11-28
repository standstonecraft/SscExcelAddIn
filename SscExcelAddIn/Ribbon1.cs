using System.Collections.Generic;
using System.Windows;
using Microsoft.Office.Tools.Ribbon;

namespace SscExcelAddIn
{
    /// <summary>
    /// リボン
    /// </summary>
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            List<RibbonControl> sheetButtons = new List<RibbonControl> {
                ReplaceButton, ZebraButton, ShapeEditButton
            };

            EnableButtons(sheetButtons, false);

            Globals.ThisAddIn.Application.WorkbookDeactivate += book => EnableButtons(sheetButtons, false);
            Globals.ThisAddIn.Application.WorkbookActivate += book => EnableButtons(sheetButtons, true);
        }

        private static void EnableButtons(List<RibbonControl> sheetButtons, bool enabled)
        {
            foreach (RibbonControl control in sheetButtons)
            {
                control.Enabled = enabled;
            }
        }
        private static bool IsSheetShown()
        {
            return Globals.ThisAddIn.Application.ActiveSheet != null;
        }

        /// <summary>
        /// 置換ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ReplaceButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!IsSheetShown())
            {
                return;
            }
            Window window = new Window
            {
                Title = "高度な置換",
                Content = new ReplaceControl(),
                // ウィンドウサイズをコンテンツに合わせる
                SizeToContent = SizeToContent.Height,
                Width = 600,
                ResizeMode = ResizeMode.CanResizeWithGrip,
                Topmost = false,
            };
            // クローズ時にExcelを操作できるようにする
            window.Closing += (sender1, e1) =>
                    Globals.ThisAddIn.Application.Interactive = true;
            window.Show();
            // オープン時にExcelを操作できないようにする
            Globals.ThisAddIn.Application.Interactive = false;
        }

        private void AboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            Window window = new Window
            {
                Title = "About",
                Content = new WebControl(Properties.Resources.About),
                Width = 500,
                Height = 500,
            };
            window.ShowDialog();
        }

        private void SkipSelectButton_Click(object sender, RibbonControlEventArgs e)
        {
            Window window = new Window
            {
                Title = "スキップ選択",
                Content = new SkipSelectControl(),
                // ウィンドウサイズをコンテンツに合わせる
                SizeToContent = SizeToContent.Height,
                Width = 300,
                ResizeMode = ResizeMode.NoResize,
                Topmost = true,
            };
            window.Closing += (sender1, e1) => System.Windows.Threading.Dispatcher.ExitAllFrames();
            window.Show();

            /*
             * WPFのWindowを開いた際に、そのWindowのTextBoxではなぜか半角入力を受け付けてくれません。
             * https://trapemiya.hatenablog.com/entry/2020/02/07/005007
             * (セル選択はできるがセル入力はできないので注意)
             */
            System.Windows.Threading.Dispatcher.Run();
        }

        private void TestControlButton_Click(object sender, RibbonControlEventArgs e)
        {
            new Window { Content = new TestControl() }.Show();
        }

        private void ShapeEditButton_Click(object sender, RibbonControlEventArgs e)
        {
            Window window = new Window
            {
                Title = "シェイプ文字列",
                Content = new ShapeEditControl(),
                Width = 500,
                Height = 360,
                ResizeMode = ResizeMode.CanResizeWithGrip,
                Topmost = true,
            };
            // クローズ時にExcelを操作できるようにする
            window.Closing += (sender1, e1) =>
                    Globals.ThisAddIn.Application.Interactive = true;
            window.Show();
        }
    }
}
