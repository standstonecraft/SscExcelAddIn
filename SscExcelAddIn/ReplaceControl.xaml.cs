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
            }.AsReadOnly();

        private static readonly int RowSize = 10;

        private readonly ObservableCollection<PreviewModel> previewList;
        private readonly ObservableCollection<BatchModel> batchList;
        private bool isBatchMode = false;
        /// <summary>
        /// 
        /// </summary>
        public ReplaceControl()
        {
            InitializeComponent();

            previewList = new ObservableCollection<PreviewModel>();
            PreviewDataGrid.ItemsSource = previewList;

            batchList = new ObservableCollection<BatchModel>();
            BatchDataGrid.ItemsSource = batchList;
            batchList.CollectionChanged += BatchListChanged;
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

            RefreshPreview();

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
            isBatchMode = batchList.Count(r => !string.IsNullOrEmpty(r.PatternText)) > 0;
            GoButton.Content = isBatchMode ? "連続" : "置換";
        }

        private void RefreshPreview()
        {
            if (previewList is null)
            {
                return;
            }
            previewList.Clear();
            List<Excel.Range> sample = new List<Excel.Range>();
            Excel.Range selection = Funcs.CellSelection();
            SetErrorLabel(false);
            if (selection != null)
            {
                // サンプルを取得する
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

                // サンプルを置換する
                foreach (Excel.Range item in sample)
                {
                    string input = item?.Formula?.ToString() ?? "";
                    string afterText = "";
                    try
                    {
                        if (isBatchMode)
                        {
                            afterText = input;
                            for (int batchIdx = 0; batchIdx < batchList.Count && batchIdx < PreviewSlider.Value; batchIdx++)
                            {
                                BatchModel model = batchList[batchIdx];
                                if (!string.IsNullOrEmpty(model.PatternText))
                                {
                                    afterText = ReplaceLogic.ReplaceText(afterText, model.PatternText, model.ReplacementText);
                                }
                            }
                        }
                        else
                        {
                            afterText = ReplaceLogic.ReplaceText(input, PatternTextBox.Text ?? "", ReplacementTextBox.Text ?? "");
                        }
                    }
                    catch
                    {
                        SetErrorLabel();
                    }
                    previewList.Add(new PreviewModel
                    {
                        BeforeText = item.Formula.ToString(),
                        AfterText = afterText
                    });
                }
            }
            for (int i = sample.Count; i < RowSize; i++)
            {
                previewList.Add(new PreviewModel());
            }
        }

        private void SetErrorLabel(bool isError = true)
        {
            ErrorLabel.Content = isError ? "エラー" : "";
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

        private class PreviewModel
        {
            public string BeforeText { get; set; } = " ";
            public string AfterText { get; set; } = " ";
        }

        private class BatchModel
        {
            public string PatternText { get; set; }
            public string ReplacementText { get; set; }

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

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            //BatchImportButton.IsEnabled = PatternTextBox.Text.Length + ReplacementTextBox.Text.Length > 0;
            RefreshPreview();
        }

        private void GoButton_Click(object sender, RoutedEventArgs e)
        {
            if (isBatchMode)
            {
                foreach (BatchModel model in batchList)
                {
                    ReplaceLogic.ReplaceTextRange(Funcs.CellSelection(), model.PatternText, model.ReplacementText);
                }
            }
            else
            {
                ReplaceLogic.ReplaceTextRange(Funcs.CellSelection(), PatternTextBox.Text, ReplacementTextBox.Text);
            }
            Window.GetWindow(this).Close();
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            PatternTextBox.Text = "";
            ReplacementTextBox.Text = "";
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
            string newDollar = string.Format("_NAR(${0}_NAR)", newDollarNum);
            Funcs.InsertTextBox(newDollar, ReplacementTextBox);
        }

        private void SeqButton_Click(object sender, RoutedEventArgs e)
        {
            int newDollarNum = NewDollarNum("${0}");
            string newDollar = string.Format("_SEQ(${0},1)", newDollarNum);
            Funcs.InsertTextBox(newDollar, ReplacementTextBox, newDollar.Length - 2, 1);
        }

        private void NumTypeButton_Click(object sender, RoutedEventArgs e)
        {
            int newDollarNum = NewDollarNum("${0}");
            string newDollar = string.Format("_CAS(${0},{1})", newDollarNum, NumTypeComboBox.SelectedValue);
            Funcs.InsertTextBox(newDollar, ReplacementTextBox);
        }

        private void ShortHandButton_Click(object sender, RoutedEventArgs e)
        {
            switch (((Button)sender).Name)
            {
                case "IncFirstNumButton":
                    PatternTextBox.Text = @"(^.*?)(" + RegexPattern.NUM.Key + ")";
                    Funcs.ClickButton(DollarButton);
                    Funcs.ClickButton(IncrementButton);
                    break;
                case "ExKanaHanButton":
                    PatternTextBox.Text = @"(.+)";
                    ReplacementTextBox.Text = "";
                    Funcs.ClickButton(HankakuButton);
                    break;
                default:
                    break;
            }
        }

        private void BatchImportButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(PatternTextBox.Text))
            {
                return;
            }
            int selectedIndex = BatchDataGrid.SelectedIndex;
            if (!string.IsNullOrEmpty(PatternTextBox.Text))
            {
                BatchModel batchModel = new BatchModel
                {
                    PatternText = PatternTextBox.Text,
                    ReplacementText = ReplacementTextBox.Text
                };

                if (selectedIndex > -1)
                {
                    batchList[selectedIndex] = batchModel;
                }
                else
                {
                    batchList.Add(batchModel);
                    BatchDataGrid.SelectedIndex = batchList.Count - 1;
                    PreviewSlider.Value = batchList.Count;
                }
                PatternTextBox.Text = "";
                ReplacementTextBox.Text = "";
            }
        }

        private void BatchExportButton_Click(object sender, RoutedEventArgs e)
        {
            int selectedIndex = BatchDataGrid.SelectedIndex;
            if (selectedIndex > -1)
            {
                PatternTextBox.Text = batchList[selectedIndex].PatternText;
                ReplacementTextBox.Text = batchList[selectedIndex].ReplacementText;
            }
        }

        private void BatchUpButton_Click(object sender, RoutedEventArgs e)
        {
            int selectedIndex = BatchDataGrid.SelectedIndex;
            if (selectedIndex > 0)
            {
                batchList.Move(selectedIndex, selectedIndex - 1);
                BatchDataGrid.SelectedIndex = selectedIndex - 1;
            }
        }

        private void BatchDownButton_Click(object sender, RoutedEventArgs e)
        {
            int selectedIndex = BatchDataGrid.SelectedIndex;
            if (selectedIndex > -1 && selectedIndex < batchList.Count - 1)
            {
                batchList.Move(selectedIndex, selectedIndex + 1);
                BatchDataGrid.SelectedIndex = selectedIndex + 1;
            }
        }

        private void BatchAddButton_Click(object sender, RoutedEventArgs e)
        {
            int selectedIndex = BatchDataGrid.SelectedIndex;
            int insert = selectedIndex < 0 ? batchList.Count : selectedIndex;

            batchList.Insert(insert, new BatchModel
            {
                PatternText = PatternTextBox.Text,
                ReplacementText = ReplacementTextBox.Text
            });
            BatchDataGrid.SelectedIndex = insert;
            PreviewSlider.Value = batchList.Count;
        }

        private void BatchRemoveButton_Click(object sender, RoutedEventArgs e)
        {
            int selectedIndex = BatchDataGrid.SelectedIndex;
            if (selectedIndex > -1)
            {
                // 要素があれば行う
                batchList.RemoveAt(BatchDataGrid.SelectedIndex);
                // 行が残っていれば、
                if (isBatchMode)
                {
                    // 元の選択行か残っている行数の小さい方を選択する
                    BatchDataGrid.SelectedIndex = Math.Min(selectedIndex, batchList.Count - 1);
                }
            }
        }

        private void BatchCopyButton_Click(object sender, RoutedEventArgs e)
        {
            string tsv = string.Join("\r\n", batchList.Select(row => row.PatternText + "\t" + row.ReplacementText));
            Clipboard.SetText(tsv);
        }

        private void BatchPasteButton_Click(object sender, RoutedEventArgs e)
        {
            if (Clipboard.ContainsText())
            {
                try
                {
                    IEnumerable<string[]> data = Clipboard.GetText().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(row => row.Split('\t'));
                    foreach (string[] row in data)
                    {
                        if (row.Length != 2)
                        {
                            throw new ArgumentException();
                        }
                        if (row[0] != "")
                        {
                            batchList.Add(new BatchModel
                            {
                                PatternText = row[0],
                                ReplacementText = row[1]
                            });
                        }
                    }
                }
                catch (Exception)
                {

                    throw;
                }
            }
        }

        private void UserControl_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F5)
            {
                RefreshPreview();
            }
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshPreview();
        }

        private void PreviewSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            RefreshPreview();
        }
    }
}
