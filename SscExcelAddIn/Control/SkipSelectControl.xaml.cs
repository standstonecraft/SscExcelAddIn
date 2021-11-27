using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using SscExcelAddIn.Logic;

namespace SscExcelAddIn
{
    /// <summary>
    /// SkipSelectControl.xaml の相互作用ロジック
    /// </summary>
    public partial class SkipSelectControl : UserControl
    {
        private static readonly int GridSize = 10;
        private static readonly string ExceptNumPrn = "[^0-9,]";
        private static readonly string NumArrayPtn = @"^(\d+)(,*\d+)*$";

        /// <summary>
        /// スキップ選択のコントロール
        /// </summary>
        public SkipSelectControl()
        {
            InitializeComponent();
            List<GridDataForSkip> list = Enumerable.Range(1, GridSize)
                .Select(i => new GridDataForSkip { RowIndex = i.ToString("D2") })
                .ToList();
            PreviewDataGrid.ItemsSource = list;
            SelectorTextBox.CaretIndex = int.MaxValue;
            Loaded += (s, e) => RefreshPreview();
            SelectorTextBox.Focus();
        }
        private class GridDataForSkip
        {
            public string RowIndex { get; set; }
        }

        private void SelectorTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshPreview();
        }

        private void Radio_Changed(object sender, RoutedEventArgs e)
        {
            RefreshPreview();
        }

        private void RefreshPreview()
        {
            if (PreviewDataGrid is null)
            {
                return;
            }
            Brush selectedBrush = new SolidColorBrush(Colors.Aqua);
            Brush invalidBrush = new SolidColorBrush(Colors.Gray);
            // reset color
            for (int rowIndex = 0; rowIndex < GridSize; rowIndex++)
            {
                for (int colIndex = 0; colIndex < GridSize; colIndex++)
                {
                    PreviewDataGrid.GetCell(rowIndex, colIndex).Background = null;
                }
                PreviewDataGrid.UpdateLayout();
            }
            // validation
            if (!Regex.IsMatch(SelectorTextBox.Text, NumArrayPtn))
            {
                for (int rowIndex = 0; rowIndex < GridSize; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < GridSize; colIndex++)
                    {
                        PreviewDataGrid.GetCell(rowIndex, colIndex).Background = invalidBrush;
                    }
                    PreviewDataGrid.UpdateLayout();
                }
                PreviewDataGrid.UpdateLayout();
                return;
            }
            IEnumerable<int> skipSelector = SelectorTextBox.Text.Split(',').Where(s => s != "").Select(s => int.Parse(s));
            SkipFilter<int> skipFilter = new SkipFilter<int>(Enumerable.Range(0, GridSize), skipSelector);
            if (RowRadio.IsChecked == true)
            {
                foreach (int target in skipFilter)
                {
                    for (int colIndex = 0; colIndex < GridSize; colIndex++)
                    {
                        PreviewDataGrid.GetCell(target, colIndex).Background = selectedBrush;
                    }
                    System.Console.WriteLine(target);
                }
            }
            else
            {
                foreach (int target in skipFilter)
                {
                    for (int rowIndex = 0; rowIndex < GridSize; rowIndex++)
                    {
                        PreviewDataGrid.GetCell(rowIndex, target).Background = selectedBrush;
                    }
                }
            }
            PreviewDataGrid.UpdateLayout();
        }



        private void GoButton_Click(object sender, RoutedEventArgs e)
        {
            SkipSelectLogic.SkipSelectRange(Funcs.CellSelection(), SelectorTextBox.Text, ColRadio.IsChecked ?? false);
            Window.GetWindow(this).Close();
        }

        private void SelectorTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.R)
            {
                RowRadio.IsChecked = true;
            }
            if (e.Key == Key.C)
            {
                ColRadio.IsChecked = true;
            }
            if (Regex.IsMatch(SelectorTextBox.Text, ExceptNumPrn))
            {
                SelectorTextBox.Text = Regex.Replace(SelectorTextBox.Text, ExceptNumPrn, "");
                SelectorTextBox.CaretIndex = int.MaxValue;
            }
        }
    }
}
