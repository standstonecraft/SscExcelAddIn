using System.Linq;
using SscExcelAddIn.ComModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.Logic
{
    /// <summary>
    /// 行列の操作により分割された条件付き書式を統合する
    /// </summary>
    public class MergeFormatCondsLogic
    {
        /// <summary>
        /// 行列の操作により分割された条件付き書式を統合する
        /// </summary>
        public static void MergeFormatConds()
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            while (true)
            {
                IGrouping<FormatConditionModel, FormatConditionModel> group =
                    sheet.Cells.FormatConditions.Cast<Excel.FormatCondition>()
                    .Select(fc => new FormatConditionModel(fc))
                    .GroupBy(fc => fc)
                    .FirstOrDefault(cg => cg.Count() > 1);
                if (group == null)
                {
                    break;
                }

                group.ElementAt(0).FormatCondition.ModifyAppliesToRange(
                    Funcs.UnionRange(group.Select(fcm => fcm.FormatCondition.AppliesTo).ToList()));
                foreach (FormatConditionModel item in group.Skip(1))
                {
                    item.FormatCondition.Delete();
                }
            }
        }
    }
}
