using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Automation.Peers;
using System.Windows.Automation.Provider;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using Button = System.Windows.Controls.Button;
using TextBox = System.Windows.Controls.TextBox;
using System.Windows;

namespace SscExcelAddIn.Logic
{
    internal static partial class Funcs
    {

        internal static void InsertTextBox(string insert, TextBox textBox, int selectStart = -1, int selectLength = -1)
        {
            string text = textBox.Text;
            int caret = textBox.CaretIndex;
            textBox.Text = text.Insert(caret, insert);
            if (selectStart < 0)
            {
                textBox.CaretIndex = caret + insert.Length;
            }
            else
            {
                textBox.CaretIndex = caret + selectStart;
                textBox.SelectionLength = selectLength;
            }

            textBox.Focus();
        }

        internal static void SurroundTextBox(string head, string tail, TextBox textBox)
        {
            textBox.SelectedText = head + textBox.SelectedText + tail;
            textBox.SelectionLength -= head.Length + tail.Length;
            textBox.SelectionStart += head.Length;
            textBox.Focus();
        }

        /// <summary>
        /// <see href="https://stackoverflow.com/questions/29806865/iinvokeprovider-invoke-in-nunit-test/29807875#29807875"/>
        /// </summary>
        public static void DoEvents()
        {
            DispatcherFrame frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background,
                    new DispatcherOperationCallback(ExitFrame), frame);
            Dispatcher.PushFrame(frame);
        }

        /// <summary>
        /// <see href="https://stackoverflow.com/questions/29806865/iinvokeprovider-invoke-in-nunit-test/29807875#29807875"/>
        /// </summary>
        public static object ExitFrame(object f)
        {
            ((DispatcherFrame)f).Continue = false;
            return null;
        }

        /// <summary>
        /// ボタンのクリックイベントを発生させる。
        /// </summary>
        /// <param name="button"></param>
        internal static void ClickButton(Button button)
        {
            ButtonAutomationPeer peer = new ButtonAutomationPeer(button);
            IInvokeProvider invokeProv = peer.GetPattern(PatternInterface.Invoke) as IInvokeProvider;
            invokeProv.Invoke();
            DoEvents();
        }

        /// <summary>
        /// セル範囲の和集合を返す。
        /// </summary>
        /// <param name="ranges"></param>
        /// <returns></returns>
        public static Excel.Range UnionRange(IEnumerable<Excel.Range> ranges)
        {
            List<Excel.Range> list = ranges.ToList();
            Excel.Range union = list[0];
            foreach (Excel.Range item in list.Skip(1))
            {
                union = Globals.ThisAddIn.Application.Union(union, item);
            }
            return union;
        }

        public static List<Excel.Range> GetSample(int size)
        {
            List<Excel.Range> sample = new List<Excel.Range>();
            Excel.Range selection = CellSelection();
            if (selection != null)
            {
                // サンプルを取得する
                IEnumerator e = selection.GetEnumerator();
                int runMax = 10000;
                int runCount = 0;
                while (e != null && e.MoveNext() && sample.Count < size && runCount < runMax)
                {
                    Excel.Range cell = (Excel.Range)e.Current;
                    if (cell.Formula != null && cell.Formula.ToString() != "")
                    {
                        sample.Add(cell);
                    }
                    runCount++;
                }
            }

            return sample;
        }

        /// <summary>
        /// 範囲が大きすぎる場合に警告する。警告した場合は真を返す。
        /// </summary>
        /// <param name="range">範囲</param>
        /// <param name="max">許容最大セル数</param>
        /// <returns>警告したかどうか</returns>
        public static bool SelectionWarning(Excel.Range range, long max)
        {
            if ((long)range.CountLarge > max)
            {
                string message = string.Format("選択範囲が広すぎます。{0:#,0}セル以内で指定してください。", max);
                MessageBox.Show(message, "SscExcelAddin", MessageBoxButton.OK, MessageBoxImage.Warning);
                return true;
            }
            return false;
        }

        /// <summary>
        /// </summary>
        /// <returns>選択された「セル範囲」。未選択の場合やシェイプが選択されている場合はnull。</returns>
        public static Excel.Range CellSelection()
        {
            dynamic selection = Globals.ThisAddIn.Application.Selection;
            return IsCellRange(selection) ? (Excel.Range)selection : null;
        }

        /// <summary>
        /// セル範囲を取得する。
        /// </summary>
        /// <param name="sheet">nullか未指定の場合、アクティブなシート</param>
        /// <param name="row">行 未指定の場合、1</param>
        /// <param name="col">列 未指定の場合、1</param>
        /// <param name="rowSize">行数 未指定の場合、1</param>
        /// <param name="colSize">列数 未指定の場合、1</param>
        /// <returns></returns>
        public static Excel.Range Range(Excel.Worksheet sheet = null, int row = 1, int col = 1, int rowSize = 1, int colSize = 1)
        {
            if (sheet == null)
            {
                sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            }
            return (Excel.Range)sheet.Range[sheet.Cells[row, col], sheet.Cells[row + rowSize - 1, col + colSize - 1]];
        }

        /// <summary>
        /// <see href="https://stackoverflow.com/questions/19943522/c-sharp-determining-the-type-of-the-selected-object-in-excel/19943737#19943737"/>
        /// </summary>
        /// <param name="thing"></param>
        /// <returns>選択範囲がセル範囲であるかどうか</returns>
        public static bool IsCellRange(dynamic thing)
        {
            if (thing is null)
            {
                return false;
            }
            Type type = GetExcelTypeForComObject(thing);
            return type == typeof(Excel.Range);
        }

        /// <summary>
        /// 
        /// <see href="https://stackoverflow.com/questions/19943522/c-sharp-determining-the-type-of-the-selected-object-in-excel/19943737#19943737"/>
        /// </summary>
        /// <param name="excelComObject"></param>
        /// <returns>Excel COMオブジェクトの型</returns>
        private static Type GetExcelTypeForComObject(object excelComObject)
        {
            if (excelComObject is null)
            {
                return null;
            }
            // get the com object and fetch its IUnknown
            IntPtr iunkwn = Marshal.GetIUnknownForObject(excelComObject);

            // enum all the types defined in the interop assembly
            System.Reflection.Assembly excelAssembly =
            System.Reflection.Assembly.GetAssembly(typeof(Excel.Range));
            Type[] excelTypes = excelAssembly.GetTypes();

            // find the first implemented interop type
            foreach (Type currType in excelTypes)
            {
                // get the iid of the current type
                Guid iid = currType.GUID;
                if (!currType.IsInterface || iid == Guid.Empty)
                {
                    // com interop type must be an interface with valid iid
                    continue;
                }

                // query supportability of current interface on object
                IntPtr ipointer = IntPtr.Zero;
                Marshal.QueryInterface(iunkwn, ref iid, out ipointer);

                if (ipointer != IntPtr.Zero)
                {
                    // yeah, that’s the one we’re after
                    return currType;
                }
            }

            // no implemented type found
            return null;
        }

        /// <summary>
        /// セル範囲にデータを書き込む。リサイズが未指定の場合、[行,列]=[値リスト件数,1]の範囲にリサイズして書き込む。
        /// isValue2 が偽の場合はFormulaプロパティに書き込む。
        /// </summary>
        /// <param name="range">セル範囲</param>
        /// <param name="values">値</param>
        /// <param name="resizeRow">リサイズする場合の行数</param>
        /// <param name="resizeCol">リサイズする場合の列数</param>
        /// <param name="isVertical">列を優先して値を書き込む(縦に書き込む)</param>
        /// <param name="isFill"></param>
        /// <param name="isValue2">Value2プロパティに書き込むかどうか</param>
        public static void WriteRange(Excel.Range range, IEnumerable<object> values, int resizeRow = -1, int resizeCol = 1, bool isVertical = true, bool isFill = false, bool isValue2 = true)
        {
            if (range is null)
            {
                throw new ArgumentNullException(nameof(range));
            }

            if (values is null)
            {
                throw new ArgumentNullException(nameof(values));
            }

            List<object> list = values.ToList();
            IEnumerator<object> enumerator = list.GetEnumerator();
            int row = resizeRow < 0 ? list.Count : resizeRow;
            int column = resizeCol < 0 ? range.Column : resizeCol;
            Excel.Range resized = resizeRow < 0 && resizeCol != 1 ? range : range.Resize[row, column];
            object[,] result = crLoop();
            if (isValue2)
            {
                resized.Value2 = result;
            }
            else
            {
                resized.Formula = result;
            }

            object[,] crLoop()
            {
                object[,] arr = new object[row, column];
                object lastValue = null;
                int outerMax = isVertical ? arr.GetLength(1) : arr.GetLength(0);
                int innerMax = isVertical ? arr.GetLength(0) : arr.GetLength(1);
                for (int outer = 0; outer < outerMax; outer++)
                {
                    for (int inner = 0; inner < innerMax; inner++)
                    {
                        int r = isVertical ? inner : outer;
                        int c = isVertical ? outer : inner;
                        if (enumerator.MoveNext())
                        {
                            lastValue = arr[r, c] = enumerator.Current.ToString();
                        }
                        else
                        {
                            if (isFill)
                            {
                                arr[r, c] = lastValue;
                            }
                            else
                            {
                                return arr;
                            }
                        }
                    }
                }
                return arr;
            }

        }

        /// <summary>
        /// 選択されたシェイプを設定値に従って拡大する
        /// </summary>
        public static void ResizeShapes()
        {
            float scale = Properties.Settings.Default.ResizePercent / 100f;
            dynamic range = Globals.ThisAddIn.Application.Selection;
            if (range == null)
            {
                return;
            }
            dynamic rangeCount = range.ShapeRange.Count;
            if (rangeCount == 1)
            {
                setScale(range, 1);
            }
            else
            {
                for (int i = 1; i <= rangeCount; i++)
                {
                    setScale(range, i);
                }
            }

            void setScale(dynamic rng, int index)
            {
                try
                {
                    Excel.Shape sr = rng.ShapeRange(index);
                    sr.ScaleHeight(scale, Microsoft.Office.Core.MsoTriState.msoFalse);
                    sr.ScaleWidth(scale, Microsoft.Office.Core.MsoTriState.msoFalse);
                }
                catch { }
            }
        }
    }
}
