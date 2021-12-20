using System;
using System.Collections.Generic;
using System.Windows.Controls;
using Microsoft.Office.Tools;
using Forms = System.Windows.Forms;

namespace SscExcelAddIn
{
    /// <summary>
    /// カスタム作業ウィンドウを管理する。カスタム作業ウィンドウはドキュメントフレームごとに生成されるので、
    /// 複数ウィンドウに対応するためにブック名やカスタム作業ウィンドウ名を使った辞書で管理する。
    /// <seealso href="https://stackoverflow.com/a/24732000"/>
    /// </summary>
    public class TaskPaneManager
    {
        private static readonly Dictionary<string, CustomTaskPane> _createdPanes = new Dictionary<string, CustomTaskPane>();

        /// <summary>
        /// Gets the taskpane by name (if exists for current excel window then returns existing instance, otherwise uses taskPaneCreatorFunc to create one). 
        /// </summary>
        /// <param name="nameOfControl">The name of WPF UserControl</param>
        /// <param name="taskPaneTitle">Display title of the taskpane</param>
        /// <param name="taskPaneCreatorFunc">The function that will construct the WPF UserControl if one does not already exist in the current Excel window.</param>
        public static CustomTaskPane GetTaskPane(string nameOfControl, string taskPaneTitle, Func<UserControl> taskPaneCreatorFunc)
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            string key = string.Format("{0}/{1}", nameOfControl, app.Hwnd);
            if (!_createdPanes.ContainsKey(key))
            {
                Forms.Integration.ElementHost _eh = new Forms.Integration.ElementHost
                {
                    Child = taskPaneCreatorFunc(),
                    Dock = Forms.DockStyle.Fill
                };
                Forms.UserControl formControl = new Forms.UserControl();
                formControl.Controls.Add(_eh);
                CustomTaskPane pane = Globals.ThisAddIn.CustomTaskPanes.Add(formControl, taskPaneTitle);
                pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                _createdPanes[key] = pane;
            }
            return _createdPanes[key];
        }
    }
}
