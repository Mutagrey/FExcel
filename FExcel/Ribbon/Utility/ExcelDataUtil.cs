using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace FExcel.FELoader.Utility
{
    public static class ExcelDataUtil
    {
        public static string[] TableStyles = new string[]
        {
            // Light
            "TableStyleLight1", "TableStyleLight2", "TableStyleLight3", "TableStyleLight4", "TableStyleLight5", "TableStyleLight6", "TableStyleLight7",
            "TableStyleLight8", "TableStyleLight9", "TableStyleLight10", "TableStyleLight11", "TableStyleLight12", "TableStyleLight13", "TableStyleLight14",
            "TableStyleLight15", "TableStyleLight16", "TableStyleLight17", "TableStyleLight18", "TableStyleLight19", "TableStyleLight20", "TableStyleLight21",
        };

        public static Excel.Application Application => (Excel.Application)ExcelDnaUtil.Application;
        public static Excel.Workbook ActiveWorkbook => Application.ActiveWorkbook;

        /// <summary>
        /// Получить таблицу
        /// </summary>
        /// <param name="tableName">имя таблицы Excel</param>
        /// <returns>Excel.ListObject</returns>
        public static Excel.ListObject GetListObject(string tableName)
        {
            try
            {
                foreach (Excel.Worksheet ws in ActiveWorkbook.Sheets)
                    foreach (Excel.ListObject table in ws.ListObjects)
                        if (table.Name == tableName) return table;
                return null;
            }
            catch (Exception)
            {
                return null;
            }


        }

        /// <summary>
        /// Convert this list object to a DataTable
        /// </summary>
        /// <typeparam name="T">Type Model</typeparam>
        /// <param name="items">List</param>
        /// <returns>DataTable</returns>
        public static System.Data.DataTable ToDataTable<T>(List<T> items)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable(typeof(T).Name);
            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }

        /// <summary>
        /// Получить список параметров модели
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns>Список параметров</returns>
        public static IList<string> GetModelParams<T>()
        {
            IList<string> modelParams = new List<string>();
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                modelParams.Add(prop.Name);
            }
            return modelParams;
        }

        public static Excel.Worksheet CreateSheetIfNotExists(string wsName)
        {
            try
            {
                Excel.Worksheet ws = ActiveWorkbook.Sheets.Cast<Excel.Worksheet>().Where(p => p.Name == wsName).FirstOrDefault();
                if (ws != null) return ws;
                Excel.Worksheet newWS = ActiveWorkbook.Worksheets.Add();
                if (wsName.Length > 0)
                    newWS.Name = wsName;
                return newWS;
            }
            catch (Exception ex)
            {
                MessageBox.Show("ExcelDataUtil.CreateSheetIfNotExists: " + ex.Message);
                return null;
            }

        }

        public static bool IsSheetExists(string wsName)
        {
            var ws = ActiveWorkbook.Sheets.Cast<Excel.Worksheet>().Where(p => p.Name == wsName).FirstOrDefault();
            if (ws != null) return true;
            return false;
        }

        public static bool IsTableExists(string tableName)
        {
            foreach (Excel.Worksheet ws in ActiveWorkbook.Sheets)
                foreach (Excel.ListObject listObject in ws.ListObjects)
                    if (listObject.Name == tableName)
                        return true;
            return false;
        }

        /// <summary>
        /// Создать или вернуть таблицу по модели данных
        /// </summary>
        /// <typeparam name="T">модель данных</typeparam>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="wsName">имя листа</param>
        /// <param name="rangeAddress">адресс ячейки таблицы</param>
        /// <returns></returns>
        public static Excel.ListObject CreateOrGetListObject<T>(string tableName, string wsName = "", string rangeAddress = "A1")
        {
            try
            {
                Excel.ListObject listObject = GetListObject(tableName);
                if (listObject != null)
                {
                    ((Excel.Worksheet)listObject.Parent).Activate();
                    listObject.Range.Select();
                    return listObject;
                }
                    
                Excel.Worksheet ws = CreateSheetIfNotExists(wsName);
                var modelParams = GetModelParams<T>();
                Excel.Range range = ws.Range[rangeAddress].Resize[1, modelParams.Count];

                // Print Data
                var res = new object[1, modelParams.Count];
                for (int i = 0; i < modelParams.Count; i++)
                {
                    res[0, i]=modelParams[i];
                }
                range.Value2 = res;

                listObject = ws.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, range, null, XlYesNoGuess.xlYes, range);
                listObject.Name = tableName;

                return listObject;
            }
            catch (Exception ex)
            {
                MessageBox.Show("ExcelDataUtil.CreateOrUpdateListObject: " + ex.Message);
                return null;
            }

        }


        public static Excel.ListObject CreateExcelTableFromTxtString(string tableName, string wsName = "", 
            string rangeAddress = "A1", string mockData = "", bool isSelected = false)
        {
            try
            {
                // Чтение из string txt
                if (mockData.Length == 0) return null;
                string[] lines = mockData.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                var columnsCount = lines[0].Split('\t').Count();
                var res = new object[lines.Count(), columnsCount];
                for (int i = 0; i < lines.Count(); i++)
                {
                    var line = lines[i];
                    var columns = line.Split('\t');
                    for (int j = 0; j < columnsCount; j++)
                    {
                        res[i, j] = columns[j];
                    }
                }

                // Заполнить данными
                Excel.Worksheet ws = CreateSheetIfNotExists(wsName);
                Excel.Range range = ws.Range[rangeAddress].Resize[res.GetLength(0), res.GetLength(1)];
                range.Value2 = res;

                // Создать таблицу
                var listObject = ws.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, range, null, XlYesNoGuess.xlYes, range);
                if (tableName.Length > 0 && !IsTableExists(tableName))
                    listObject.Name = tableName;

                return listObject;
            }
            catch (Exception ex)
            {
                MessageBox.Show("ExcelDataUtil.CreateOrGetExcelTableFromTxtString: " + ex.Message);
                return null;
            }

        }

    }
}
