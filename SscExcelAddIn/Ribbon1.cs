using System;
using System.Collections.Generic;
using System.Windows;
using Microsoft.Office.Tools.Ribbon;
using Reactive.Bindings;
using SscExcelAddIn.Logic;

namespace SscExcelAddIn
{
    /// <summary>
    /// リボン
    /// </summary>
    public partial class Ribbon1
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            List<RibbonComponent> sheetComponents = new List<RibbonComponent>();
            sheetComponents.AddRange(editSheetGroup.Items);
            sheetComponents.AddRange(editShapeGroup.Items);

            EnableButtons(sheetComponents, false);

            Globals.ThisAddIn.Application.WorkbookDeactivate += book => EnableButtons(sheetComponents, false);
            Globals.ThisAddIn.Application.WorkbookActivate += book => EnableButtons(sheetComponents, true);

            ResizeTextBox.Text = Properties.Settings.Default.ResizePercent.ToString();

            CheckUpdate();
        }

        /// <summary>
        /// 更新チェック
        /// </summary>
        private void CheckUpdate()
        {
            ReactiveCommand<string> updateNotifyCommand = new ReactiveCommand<string>();
            // 更新があった場合の挙動
            _ = updateNotifyCommand.Subscribe((vers) =>
              {
                  // 更新がありますボタンを可視化する
                  updateGroup.Visible = true;

                  Properties.Settings settings = Properties.Settings.Default;
                  if (settings.UpdateNotifyVersion == vers)
                  {
                      // 過去に確認ダイアログ表示済みの場合は表示しない
                      // 新たな更新があった場合は表示する
                      return;
                  }
                  settings.UpdateNotifyVersion = vers;
                  settings.Save();
                  // ダイアログを開く
                  string message = $"新しいバージョンが利用できます。配布ページにアクセスしますか？\n{vers}";
                  string caption = "更新確認";
                  MessageBoxResult messageBoxResult = MessageBox.Show(message, caption, MessageBoxButton.YesNo, MessageBoxImage.Information);
                  if (messageBoxResult == MessageBoxResult.Yes)
                  {
                      System.Diagnostics.Process.Start(Properties.Resources.ReleasePageUrl);
                  }
              });
            // 更新チェック
            Ribbon1Logic.CheckUpdate(updateNotifyCommand);
        }

        private static void EnableButtons(List<RibbonComponent> sheetButtons, bool enabled)
        {
            foreach (RibbonControl control in sheetButtons)
            {
                control.Enabled = enabled;
            }
        }

        private static bool IsSheetShown()
            => Globals.ThisAddIn.Application.ActiveSheet != null;

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

        /// <summary>
        /// Aboutボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            Window window = new Window
            {
                Title = "About",
                Content = new WebControl(Properties.Resources.About),
                Width = 600,
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
            => new Window { Content = new TestControl() }.Show();

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

        private void ResizeButton_Click(object sender, RibbonControlEventArgs e)
            => Funcs.ResizeShapes();

        private void ResizeTextBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string text = ResizeTextBox.Text;
            if (int.TryParse(text, out int percent))
            {
                Properties.Settings settings = Properties.Settings.Default;
                settings.ResizePercent = percent;
                settings.Save();
            }
        }

        private void UpdateButton_Click(object sender, RibbonControlEventArgs e)
            => System.Diagnostics.Process.Start(Properties.Resources.ReleasePageUrl);

        private void RemoveEmptyColButton_Click(object sender, RibbonControlEventArgs e)
            => Ribbon1Logic.RemoveEmptyCol();

        private void RemoveEmptyRowButton_Click(object sender, RibbonControlEventArgs e)
            => Ribbon1Logic.RemoveEmptyRow();

        private void AggregateButton_Click(object sender, RibbonControlEventArgs e)
            => Ribbon1Logic.AggregateRange();

        private void MergeFormatCondsButton_Click(object sender, RibbonControlEventArgs e)
            => Ribbon1Logic.MergeFormatConds();
    }
}
