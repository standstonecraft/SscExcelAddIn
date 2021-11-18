using SscExcelAddIn.Logic;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn
{
    /// <summary>
    /// RegexControl.xaml の相互作用ロジック
    /// </summary>
    public partial class RegexControl : UserControl
    {
        private static IEnumerable<string[]> ButtonHandlings = new List<string[]> {
                new string[]{"HatButton","1","^" },
                new string[]{"ParenButton","2","(", ")"},
                new string[]{"ParenEscButton","2",@"\(", @"\)"},
                new string[]{ "ZenParenButton", "1",@"[（）]"},
                new string[]{"NumButton","1",@"[0-9]+"},
                new string[]{"ZenNumButton","1",@"[０-９]+"},
                new string[]{"MaruNumButton","1",@"[①-⑳]"},
                new string[]{"UpperButton","1",@"[A-Z]"},
                new string[]{"LowerButton","1",@"[a-z]"},
                new string[]{"ZenKanaButton","1",@"[ア-ン]"},
                new string[]{"ZenkakuButton","1",@"[^\x01-\x7E\xA1-\xDF]+"},
            }.AsReadOnly();

        private const int RowSize = 10;

        private ObservableCollection<PreviewData> _previewList;

        public RegexControl()
        {
            InitializeComponent();
            _previewList = new ObservableCollection<PreviewData>();
            PreviewDataGrid.ItemsSource = _previewList;
            Globals.ThisAddIn.Application.SheetSelectionChange += (object Sh, Excel.Range Target) => RefreshPreview();
            dynamic selection = Globals.ThisAddIn.Application.Selection;
            RefreshPreview();
        }

        private void PatternButton_Click(object sender, RoutedEventArgs e)
        {
            string[] handling = ButtonHandlings.Where(bh => bh[0] == ((Button)sender).Name).FirstOrDefault();
            if (handling[1] == "1")
            {
                Funcs.InsertTextBox(handling[2], PatternTextBox);
            }
            else
            {
                Funcs.SurroundTextBox(handling[2], handling[3], PatternTextBox);

            }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshPreview();
        }

        private void GoButton_Click(object sender, RoutedEventArgs e)
        {
            RegexLogic.ReplaceTextRange(Funcs.CellSelection(), PatternTextBox.Text, ReplacementTextBox.Text);
            Window.GetWindow(this).Close();
        }


        private void QuitButton_Click(object sender, RoutedEventArgs e)
        {
            Window.GetWindow(this).Close();
        }

        private void DollarButton_Click(object sender, RoutedEventArgs e)
        {
            string newDollar = "$" + NewDollarNum("${0}");
            Funcs.InsertTextBox(newDollar, ReplacementTextBox);
        }

        private void IncrementButton_Click(object sender, RoutedEventArgs e)
        {
            int newDollarNum = NewDollarNum("${0}");
            string newDollar = string.Format("_INC(${0},1)", newDollarNum);
            Funcs.InsertTextBox(newDollar, ReplacementTextBox, newDollar.Length - 2, 1);
        }

        private void HankakuButton_Click(object sender, RoutedEventArgs e)
        {
            int newDollarNum = NewDollarNum("${0}");
            string newDollar = string.Format("_NAR_(${0}_NAR_)", newDollarNum);
            Funcs.InsertTextBox(newDollar, ReplacementTextBox);
        }

        private void ShortHandButton_Click(object sender, RoutedEventArgs e)
        {
            switch (((Button)sender).Name)
            {
                case "IncFirstNumButton":
                    PatternTextBox.Text = @"(^.*?)([0-9０-９Ⅰ-Ⅹ①-⑳]+)";
                    ReplacementTextBox.Text = "";
                    Funcs.ClickButton(DollarButton);
                    Funcs.ClickButton(IncrementButton);
                    break;
                case "ExKanaHanButton":
                    PatternTextBox.Text = @"([^ァ-ヶ]+)";
                    ReplacementTextBox.Text = "";
                    Funcs.ClickButton(HankakuButton);
                    break;
                default:
                    break;
            }
        }

        private void RefreshPreview()
        {
            _previewList.Clear();
            List<Excel.Range> sample = new List<Excel.Range>();
            Excel.Range selection = Funcs.CellSelection();
            SetErrorLabel(false);
            if (selection != null)
            {
                IEnumerator e = selection.GetEnumerator();
                int runMax = 10000;
                int runCount = 0;
                while (e != null && e.MoveNext() && sample.Count < RowSize && runCount < runMax)
                {
                    Excel.Range cell = (Excel.Range)e.Current;
                    if (cell.Formula != null && cell.Formula.ToString() != "")
                    {
                        sample.Add(cell);
                    }
                    runCount++;
                }
                string pattern = PatternTextBox.Text ?? "";
                string replacement = ReplacementTextBox.Text ?? "";
                foreach (var item in sample)
                {
                    string input = item?.Formula?.ToString() ?? "";
                    string afterText = "";
                    try
                    {
                        afterText = RegexLogic.ReplaceText(input, pattern, replacement);
                    }
                    catch
                    {
                        SetErrorLabel();
                    }
                    _previewList.Add(new PreviewData
                    {
                        BeforeText = item.Formula.ToString(),
                        AfterText = afterText
                    });
                }
            }
            for (int i = sample.Count; i < RowSize; i++)
            {
                _previewList.Add(new PreviewData());
            }
        }

        private void SetErrorLabel(bool isError = true)
        {
            ErrorLabel.Content = isError ? "エラー" : "";
        }

        private class PreviewData
        {
            public string BeforeText { get; set; } = " ";
            public string AfterText { get; set; } = " ";
        }

        private int NewDollarNum(string format)
        {
            int newNum = 1;
            while (ReplacementTextBox.Text.Contains(string.Format(format, newNum)))
            {
                newNum++;
            }
            return newNum;
        }
    }
}
