using System;
using System.Runtime.InteropServices;
using System.Windows.Automation.Peers;
using System.Windows.Automation.Provider;
using Microsoft.Office.Interop.Excel;
using Button = System.Windows.Controls.Button;
using Excel = Microsoft.Office.Interop.Excel;
using TextBox = System.Windows.Controls.TextBox;

namespace SscExcelAddIn
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
        /// ボタンのクリックイベントを発生させる。
        /// </summary>
        /// <param name="button"></param>
        internal static void ClickButton(Button button)
        {
            ButtonAutomationPeer peer = new ButtonAutomationPeer(button);
            IInvokeProvider invokeProv = peer.GetPattern(PatternInterface.Invoke) as IInvokeProvider;
            invokeProv.Invoke();
        }

        /// <summary>
        /// </summary>
        /// <returns>選択された「セル範囲」。未選択の場合やシェイプが選択されている場合はnull。</returns>
        public static Excel.Range CellSelection()
        {
            dynamic selection = Globals.ThisAddIn.Application.Selection;
            return IsCellRange(selection) ? (Range)selection : null;
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
    }
}
