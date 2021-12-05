using System;
using System.Deployment.Application;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Reactive.Bindings;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn.Logic
{
    public class Ribbon1Logic
    {
        /// <summary>
        /// GitHubのリリース情報に非同期でアクセスして更新確認を行う。
        /// バージョン比較にはタグを使用する。ドット区切りであればその個数は問わない。
        /// デバッグ時は現行バージョンが確認できないためv0.0.0.1として扱う。
        /// </summary>
        /// <param name="updateNotifyCommand">新しいバージョンがある場合に起動するCommand</param>
        public static void CheckUpdate(ReactiveCommand<string> updateNotifyCommand)
        {
            _ = Task.Run(() =>
            {
                string currentVersion;
                try
                {
                    currentVersion = "v" + ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
                }
                catch (Exception)
                {
                    currentVersion = "v0.0.0.1";
                }
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Properties.Resources.ReleaseApiUrl);
                request.ContentType = "application/json; charset=utf-8";
                request.UserAgent = @"Mozilla/5.0 (iPhone; CPU iPhone OS 14_5 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) CriOS/91.0.4472.80 Mobile/15E148 Safari/604.1";

                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                using (Stream responseStream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                    dynamic json = JsonConvert.DeserializeObject(reader.ReadToEnd());
                    string publishedVersion = json.tag_name;
                    if (longVersion(currentVersion) < longVersion(publishedVersion))
                    {
                        updateNotifyCommand.Execute($"{currentVersion} => {publishedVersion}");
                    }

                }
                double longVersion(string verStr)
                {
                    string numStr = verStr.Replace("v", "");
                    double ret = 0;
                    string[] vs = numStr.Split('.');
                    for (int i = 0; i < vs.Length; i++)
                    {
                        ret += long.Parse(vs[i]) * Math.Pow(100, 4 - i);
                    }
                    return ret;
                }
            });
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

        /// <summary>
        /// 空列削除
        /// </summary>
        public static void RemoveEmptyCol()
        {
            try
            {
                Globals.ThisAddIn.Application.Interactive = false;
                Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                Excel.Range range = Funcs.CellSelection();
                int colStart = range.Column;
                int colEnd = colStart + range.Columns.Count - 1;
                for (int col = colEnd; col >= colStart; col--)
                {
                    int countA = (int)Globals.ThisAddIn.Application.WorksheetFunction.CountA(sheet.Columns[col]);
                    if (countA == 0)
                    {
                        ((Excel.Range)sheet.Columns[col]).EntireColumn.Delete();
                    }
                }
            }
            finally
            {
                Globals.ThisAddIn.Application.Interactive = true;
            }
        }


        /// <summary>
        /// 空行削除
        /// </summary>
        public static void RemoveEmptyRow()
        {
            try
            {
                Globals.ThisAddIn.Application.Interactive = false;
                Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                Excel.Range range = Funcs.CellSelection();
                int rowStart = range.Row;
                int rowEnd = rowStart + range.Rows.Count - 1;
                for (int row = rowEnd; row >= rowStart; row--)
                {
                    int countA = (int)Globals.ThisAddIn.Application.WorksheetFunction.CountA(sheet.Rows[row]);
                    if (countA == 0)
                    {
                        ((Excel.Range)sheet.Rows[row]).EntireRow.Delete();
                    }
                }
            }
            finally
            {
                Globals.ThisAddIn.Application.Interactive = true;
            }
        }
    }
}
