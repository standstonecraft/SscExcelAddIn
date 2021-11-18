using System.Windows;

namespace SscExcelAddIn
{
    public partial class Ribbon1
    {
        public static void ShowReplaceWindow()
        {
            Window window = new Window
            {
                Title = "高度な置換",
                Content = new RegexControl(),
                // ウィンドウサイズをコンテンツに合わせる
                SizeToContent = SizeToContent.Height,
                Width = 600,
                ResizeMode = ResizeMode.CanResizeWithGrip,
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

        public static void ShowAboutWindow()
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

        private static void ShowSkipSelectWindow()
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
    }
}
