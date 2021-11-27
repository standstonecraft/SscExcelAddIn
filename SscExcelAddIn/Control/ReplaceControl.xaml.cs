using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using SscExcelAddIn.Logic;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn
{
    /// <summary>
    /// ReplaceControl.xaml の相互作用ロジック
    /// </summary>
    public partial class ReplaceControl : UserControl
    {
        private static readonly ReadOnlyCollection<string[]> ButtonHandlings = new List<string[]> {
                new string[]{ "HatButton", "1", "^" },
                new string[]{ "EndButton", "1", "$" },
                new string[]{ "ParenButton", "2", "(",  ")" },
                new string[]{ "ParenEscButton", "2", @"\(",  @"\)" },
                new string[]{ "ZenParenButton",  "1", @"[（）]" },
                new string[]{ "NumButton", "1", @"[0-9]+" },
                new string[]{ "ZenNumButton", "1", @"[０-９]+" },
                new string[]{ "MaruNumButton", "1", RegexPattern.MARU.Key },
                new string[]{ "UpperButton", "1", @"[A-Z]" },
                new string[]{ "LowerButton", "1", @"[a-z]" },
                new string[]{ "ZenKanaButton", "1", @"[ァ-ヺ]" },
                new string[]{ "HanKanaButton", "1", @"[｡-ﾟ]" },
                new string[]{ "ZenkakuButton", "1", RegexPattern.ZEN.Key },
                new string[]{ "AllNumButton", "1", RegexPattern.NUM.Key },
                new string[]{ "AllCharButton", "1", RegexPattern.ALL.Key },
            }.AsReadOnly();


        //private readonly ObservableCollection<PreviewModel> previewList;
        //private readonly ObservableCollection<BatchModel> batchList;
        private readonly ReplaceViewModel vm;
        /// <summary>
        /// 
        /// </summary>
        public ReplaceControl()
        {
            InitializeComponent();

            vm = new ReplaceViewModel();
            DataContext = vm;

            //previewList = new ObservableCollection<PreviewModel>();
            //PreviewDataGrid.ItemsSource = previewList;

            //batchList = new ObservableCollection<BatchModel>();
            //BatchDataGrid.ItemsSource = batchList;
            vm.Batch.Data.CollectionChanged += BatchListChanged;
            BatchDataGrid.SelectionChanged += BatchDataGridSelectionChanged;

            NumTypeComboBox.ItemsSource = new Dictionary<string, string>
            {
                {"NN","半角数字"},
                {"NW","全角数字"},
                {"M","丸囲み数字"},
                {"RU","大文字ローマ数字"},
                {"ALN","英小文字"},
                {"AUN","英大文字"},
                {"KW","全角カタカナ"},
                {"KN","半角カタカナ"}
            };
            NumTypeComboBox.SelectedIndex = 3;

            //RefreshPreview();

            PatternTextBox.Focus();
        }

        private void BatchDataGridSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            bool enabled = BatchDataGrid.SelectedIndex > -1;

            BatchImportButton.IsEnabled = enabled;
            BatchExportButton.IsEnabled = enabled;
            BatchUpButton.IsEnabled = enabled;
            BatchDownButton.IsEnabled = enabled;

        }

        private void BatchListChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            //isBatchMode = vm.Batch.BatchList.Count(r => !string.IsNullOrEmpty(r.PatternText)) > 0;
            //GoButton.Content = isBatchMode ? "連続" : "置換";
        }

        private static List<Excel.Range> GetSample()
        {
            List<Excel.Range> sample = new List<Excel.Range>();
            Excel.Range selection = Funcs.CellSelection();
            int rowSize = 10;
            if (selection != null)
            {
                // サンプルを取得する
                IEnumerator e = selection.GetEnumerator();
                int runMax = 10000;
                int runCount = 0;
                while (e != null && e.MoveNext() && sample.Count < rowSize && runCount < runMax)
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

        private void SetErrorLabel(bool isError = true)
        {
            ErrorLabel.Content = isError ? "エラー" : "";
        }

        private int NewDollarNum(string format)
        {
            int newNum = 1;
            while (vm.ReplacementText.Value.Contains(string.Format(format, newNum)))
            {
                newNum++;
            }
            return newNum;
        }

        private void PatternButton_Click(object sender, RoutedEventArgs e)
        {
            string[] handling = ButtonHandlings.First(bh => bh[0] == ((Button)sender).Name);
            if (handling[1] == "1")
            {
                Funcs.InsertTextBox(handling[2], PatternTextBox);
            }
            else
            {
                Funcs.SurroundTextBox(handling[2], handling[3], PatternTextBox);

            }
        }

        private void ReplacementButton_Click(object sender, RoutedEventArgs e)
        {
            string newDollar;
            int newDollarNum;
            switch (((Button)sender).Name)
            {
                case "DollarButton":
                    newDollar = "$" + NewDollarNum("${0}");
                    Funcs.InsertTextBox(newDollar, ReplacementTextBox);
                    break;
                case "IncrementButton":
                    newDollarNum = NewDollarNum("${0}");
                    newDollar = string.Format("_INC(${0},1)", newDollarNum);
                    Funcs.InsertTextBox(newDollar, ReplacementTextBox, newDollar.Length - 2, 1);
                    break;
                case "HankakuButton":
                    newDollarNum = NewDollarNum("${0}");
                    newDollar = string.Format("_NAR(${0}_NAR)", newDollarNum);
                    Funcs.InsertTextBox(newDollar, ReplacementTextBox);
                    break;
                case "SeqButton":
                    newDollarNum = NewDollarNum("${0}");
                    newDollar = string.Format("_SEQ(${0},1)", newDollarNum);
                    Funcs.InsertTextBox(newDollar, ReplacementTextBox, newDollar.Length - 2, 1);
                    break;
                case "NumTypeButton":
                    newDollarNum = NewDollarNum("${0}");
                    newDollar = string.Format("_CAS(${0},{1})", newDollarNum, NumTypeComboBox.SelectedValue);
                    Funcs.InsertTextBox(newDollar, ReplacementTextBox);
                    break;
                default:
                    throw new NotSupportedException();
            }
        }

        private void ShortHandButton_Click(object sender, RoutedEventArgs e)
        {
            PatternTextBox.Text = "";
            ReplacementTextBox.Text = "";
            switch (((Button)sender).Name)
            {
                case "IncFirstNumButton":
                    PatternTextBox.Text = @"(" + RegexPattern.NUM.Key + ")(.*)";
                    Funcs.ClickButton(IncrementButton);
                    ReplacementTextBox.CaretIndex = int.MaxValue;
                    Funcs.ClickButton(DollarButton);
                    break;
                case "IncFirstCharButton":
                    PatternTextBox.Text = @"(" + RegexPattern.ALL.Key + ")(.*)";
                    Funcs.ClickButton(IncrementButton);
                    ReplacementTextBox.CaretIndex = int.MaxValue;
                    Funcs.ClickButton(DollarButton);
                    break;
                case "SeqFirstNumButton":
                    PatternTextBox.Text = @"(" + RegexPattern.NUM.Key + ")(.*)";
                    Funcs.ClickButton(SeqButton);
                    ReplacementTextBox.CaretIndex = int.MaxValue;
                    Funcs.ClickButton(DollarButton);
                    break;
                case "SeqFirstCharButton":
                    PatternTextBox.Text = @"(" + RegexPattern.ALL.Key + ")(.*)";
                    Funcs.ClickButton(SeqButton);
                    ReplacementTextBox.CaretIndex = int.MaxValue;
                    Funcs.ClickButton(DollarButton);
                    break;
                case "ExKanaHanButton":
                    PatternTextBox.Text = @"(.+)";
                    vm.ReplacementText.Value = "";
                    Funcs.ClickButton(HankakuButton);
                    break;
                default:
                    break;
            }
        }
        private void GoButton_Click(object sender, RoutedEventArgs e)
        {
            if (vm.Batch.IsBatchMode.Value)
            {
                foreach (BatchModel model in vm.Batch.Data)
                {
                    ReplaceLogic.ReplaceTextRange(Funcs.CellSelection(), model.PatternText, model.ReplacementText);
                }
            }
            else
            {
                ReplaceLogic.ReplaceTextRange(Funcs.CellSelection(), PatternTextBox.Text, vm.ReplacementText.Value);
            }
            Window.GetWindow(this).Close();
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            PatternTextBox.Text = "";
            vm.ReplacementText.Value = "";
        }

        private void QuitButton_Click(object sender, RoutedEventArgs e)
        {
            Window.GetWindow(this).Close();
        }

        private void UserControl_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F5)
            {
                vm.RefreshCommand.Execute();
            }
        }
    }
}
