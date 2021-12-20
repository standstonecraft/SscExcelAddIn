using System.Collections;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using SscExcelAddIn.Logic;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn
{
    /// <summary>
    /// ShapeEdit.xaml の相互作用ロジック
    /// </summary>
    public partial class ShapeEditControl : UserControl
    {
        private readonly ShapeEditViewModel vm;
        /// <summary>読み込むセル範囲の上限</summary>
        private const int LoadRangeMax = 2000;
        /// <summary>読み込むシェイプ数の上限</summary>
        private const int LoadShapeMax = 2000;

        /// <summary>
        /// ctor
        /// </summary>
        public ShapeEditControl()
        {
            InitializeComponent();

            vm = new ShapeEditViewModel();
            DataContext = vm;

            LoadRangeButton.Focus();
        }

        private void LoadRange()
        {
            vm.CellContents.Clear();
            IEnumerator rangeEnumerator = Funcs.CellSelection()?.GetEnumerator();
            int loopCount = 0;
            while (rangeEnumerator != null && rangeEnumerator.MoveNext() && loopCount++ < LoadRangeMax)
            {
                Excel.Range cell = (Excel.Range)rangeEnumerator.Current;
                vm.CellContents.Add(new CellContentModel
                {
                    Value = (string)cell.Value2?.ToString(),
                    Address = (string)cell.Address,
                });
            }
        }

        private void LoadShape()
        {
            dynamic range = Globals.ThisAddIn.Application.Selection;
            if (range == null)
            {
                return;
            }
            vm.ShapeContents.Clear();
            dynamic rangeCount = range.ShapeRange.Count;
            if (rangeCount == 1)
            {
                try
                {
                    dynamic sr = range.ShapeRange(1);
                    vm.ShapeContents.Add(new ShapeContentModel(sr));
                }
                catch { }
            }
            else
            {
                List<ShapeContentModel> spContents = new List<ShapeContentModel>();
                for (int i = 1; i <= rangeCount && i < LoadShapeMax; i++)
                {
                    try
                    {
                        dynamic sr = range.ShapeRange(i);
                        spContents.Add(new ShapeContentModel(sr));
                    }
                    catch { }
                }
                spContents.Sort((ColSortCheckBox.IsChecked ?? false)
                    ? ShapeContentModel.ColRowComparer
                    : ShapeContentModel.RowColComparer);
                foreach (ShapeContentModel item in spContents)
                {
                    vm.ShapeContents.Add(item);
                }
            }

        }

        private void LoadRangeButton_Click(object sender, RoutedEventArgs e)
            => LoadRange();

        private void LoadShapeButton_Click(object sender, RoutedEventArgs e)
            => LoadShape();

        private void RangeGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
            => Globals.ThisAddIn.Application.Interactive = false;

        private void RangeGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
            => Globals.ThisAddIn.Application.Interactive = true;

        private void ColSortCheckBox_Checked(object sender, RoutedEventArgs e)
            => LoadShape();

        private void EmbedButton_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < vm.CellContents.Count && i < vm.ShapeContents.Count; i++)
            {
                CellContentModel cell = vm.CellContents[i];
                vm.ShapeContents[i].Range.DrawingObject.Text = cell.Value;
            }
        }

        private void EmbedFormulaButton_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < vm.CellContents.Count && i < vm.ShapeContents.Count; i++)
            {
                CellContentModel cell = vm.CellContents[i];
                vm.ShapeContents[i].Range.DrawingObject.Formula = "=" + cell.Address;
            }
        }

        private void WriteButton_Click(object sender, RoutedEventArgs e)
        {
            dynamic selection = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            Funcs.WriteRange(selection, vm.ShapeContents);
        }

        private void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            vm.Search();
            ShapeGrid.SelectedItems.Clear();
            foreach (ShapeContentModel item in vm.SearchResults)
            {
                ShapeGrid.SelectedItems.Add(item);
            }
        }

        private void SearchNextButton_Click(object sender, RoutedEventArgs e)
        {
            ShapeContentModel hit = vm.SearchNext();
            ShapeGrid.SelectedItems.Clear();
            ShapeGrid.SelectedItems.Add(hit);
            ShapeGrid.ScrollIntoView(hit);
        }

        private void SearchPrevButton_Click(object sender, RoutedEventArgs e)
        {
            ShapeContentModel hit = vm.SearchPrev();
            ShapeGrid.SelectedItems.Clear();
            ShapeGrid.SelectedItems.Add(hit);
            ShapeGrid.ScrollIntoView(hit);
        }

        private void SearchTextBox_MouseEnter(object sender, MouseEventArgs e)
            => Globals.ThisAddIn.Application.Interactive = false;

        private void SearchTextBox_MouseLeave(object sender, MouseEventArgs e)
            => Globals.ThisAddIn.Application.Interactive = true;

        private void SearchTextBox_GotFocus(object sender, RoutedEventArgs e)
            => Globals.ThisAddIn.Application.Interactive = false;

        private void ShapeGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (e.MouseDevice.DirectlyOver is FrameworkElement element1
                && element1.Parent is DataGridCell && sender is DataGrid grid
                && grid.SelectedItems != null && grid.SelectedItems.Count == 1
                && grid.SelectedItem is ShapeContentModel shapeContent)
            {
                Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
                dynamic selected = activeSheet.Cells[shapeContent.Row, shapeContent.Column];
                Globals.ThisAddIn.Application.Goto(selected, true);
                foreach (ShapeContentModel sc in vm.ShapeContents)
                {
                    sc.Range.Select(false);
                }
            }
        }
    }
}
