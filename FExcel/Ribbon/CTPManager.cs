using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using FExcel.FELoader;
using FExcel.FELoader.View;
using Excel = Microsoft.Office.Interop.Excel;

namespace FExcel
{
    /////////////// Helper class to manage CTP ///////////////////////////
    internal static class CTPManager
    {
        static CustomTaskPane currentCTP;
        static Dictionary<string, CustomTaskPane> _createdPanes = new Dictionary<string, CustomTaskPane>();
        static Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
       
        public static void ShowCTP()
        {
            currentCTP = GetTaskPane("LoadTaskPane", "Fast Loader", () => new FExcelLoaderUserControl());
            if (currentCTP == null)
                return;

            currentCTP.Visible = true;
            currentCTP.Width = 800;
            currentCTP.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
        }

        public static void DeleteCTP()
        {
            if (currentCTP != null)
            {
                //// Could hide instead, by calling ctp.Visible = false;
                //currentCTP.Delete();
                //currentCTP = null;
                currentCTP.Visible = false;
            }
        }

        /// <summary>
        /// Gets the taskpane by name (if exists for current excel window then returns existing instance, otherwise uses taskPaneCreatorFunc to create one). 
        /// </summary>
        /// <param name="taskPaneId">Some string to identify the taskpane</param>
        /// <param name="taskPaneTitle">Display title of the taskpane</param>
        /// <param name="taskPaneCreatorFunc">The function that will construct the taskpane if one does not already exist in the current Excel window.</param>
        public static CustomTaskPane GetTaskPane(string taskPaneId, string taskPaneTitle, Func<UserControl> taskPaneCreatorFunc)
        {
            string key = string.Format("{0}({1})", taskPaneId, xlApp.Hwnd);
            if (!_createdPanes.ContainsKey(key))
            {
                var pane = CustomTaskPaneFactory.CreateCustomTaskPane(taskPaneCreatorFunc(), taskPaneTitle);
                _createdPanes[key] = pane;
            }
            return _createdPanes[key];
        }

    }
}
