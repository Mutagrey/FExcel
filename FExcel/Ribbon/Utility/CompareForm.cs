using FastExcelDNA;
using FExcel.FELoader.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FExcel.FELoader.Utility
{
    public static class CompareForm
    {
        private const int FIRST_ROW = 8;
        private static readonly int firstYear = Convert.ToInt16(Properties.Settings.Default.FirstYear);
        private static readonly int years = Convert.ToInt16(Properties.Settings.Default.Years);

        private static readonly object[,] headerData = new object[5, 2]
        { 
            { "Первый год", "" },
            { "Версия", "" },
            { "Вариант", "" },
            { "Процесс", "" },
            { "Кд", 6.09/7.404 } 
        };

        public static void CreateOrUpdate(LoadListModel loadListModel, IList<ParamModel> paramList)
        {
            var ws = ExcelDataUtil.CreateSheetIfNotExists(loadListModel.TableName);
            var selectedParams = GetAllExistedParamsList(loadListModel, paramList).Where(p => p.RowID > 0).ToList();
            var maxRowID = selectedParams.Select(p => p.RowID).Max();

            // Desc
            var descRowCount = (loadListModel.LoadListItems.Count + 1) * selectedParams.Count;
            var descData = new object[maxRowID + loadListModel.LoadListItems.Count + 1, 6];
            descData[0, 0] = "Параметр";
            descData[0, 1] = "Группа";
            descData[0, 2] = "ДО";
            descData[0, 3] = "Опция";
            descData[0, 4] = "Доля";
            descData[0, 5] = "Ключ";

            foreach (ParamModel paramModel in selectedParams)
            {
                var curRow = paramModel.RowID;
                for (int i = 0; i < loadListModel.LoadListItems.Count; i++)
                {
                    var fileElement = loadListModel.LoadListItems[i];
                    descData[curRow + i - FIRST_ROW, 0] = paramModel.Name;
                    descData[curRow + i - FIRST_ROW, 1] = fileElement.Group;
                    descData[curRow + i - FIRST_ROW, 2] = fileElement.OG;
                    descData[curRow + i - FIRST_ROW, 3] = fileElement.Mest;
                    descData[curRow + i - FIRST_ROW, 4] = fileElement.MFSOKoeff;
                    descData[curRow + i - FIRST_ROW, 5] = string.Join("|", new object[] { paramModel.Name, fileElement.Mest });
                }
            }
            var descRange = CreateElement(ws, descData, "B" + FIRST_ROW);

            // Years
            var yearData = new object[1, years];
            for (int i = 0; i < years; i++)
                yearData[0, i] = firstYear + i;
            var yearsRange = CreateElement(ws, yearData, "H" + FIRST_ROW);
            yearsRange.Offset[-1, 0].Value2 = "Года";

            // Header
            headerData[0, 1] = firstYear;
            headerData[1, 1] = DateTime.Now.ToString();
            headerData[2, 1] = loadListModel.TableName;
            headerData[3, 1] = "";
            var headerRange = CreateElement(ws, headerData, "B2");
            StyleRangeAsTable(headerRange, 1);
            headerRange = ws.Range["B1"];
            headerRange.Value2 = "CompareForm_v3";
            headerRange.Font.Italic = true;
            headerRange.Font.Color = Color.Gray;

            MessageBox.Show("Created Compare!");
        }

        private static void StyleRangeAsTable(Excel.Range range, int styleID = 1)
        {
            var ws = (Excel.Worksheet)range.Parent;
            if (ws == null) return;

            if (styleID > ExcelDataUtil.TableStyles.Count()) return;

            var style = ExcelDataUtil.TableStyles[styleID];
            var table = ws.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, range, false, Excel.XlYesNoGuess.xlYes, range, style);
            table.Unlist();

        }

        private static Excel.Range CreateElement(Excel.Worksheet ws, object[,] data, string rngRef)
        {
            Excel.Range range = ws.Range[rngRef];
            if (range == null) return null;
            range = range.Resize[data.GetLength(0), data.GetLength(1)];
            range.Value2 = data;
            return range;
        }


        private static IList<ParamModel> GetAllExistedParamsList(LoadListModel loadListModel, IList<ParamModel> paramsList)
        {
            var ws = ExcelDataUtil.CreateSheetIfNotExists(loadListModel.TableName);
            if (ws == null) return paramsList;

            var existedParamsDic = GetExistedParamsDic(loadListModel.TableName);

            var maxRowID = FIRST_ROW;
            if (existedParamsDic.Count > 0)
                maxRowID = existedParamsDic.Select(p => p.Value).Max();

            foreach (ParamModel model in paramsList)
            {
                if (existedParamsDic.ContainsKey(model.Name))
                {
                    model.RowID = existedParamsDic[model.Name];
                }
                else
                {
                    model.RowID = -1;
                    if (model.IsSelected)
                    {
                        model.RowID = maxRowID + 1;
                        maxRowID += loadListModel.LoadListItems.Count + 1;
                    }
                }
            }

            return paramsList;
        }

        private static Dictionary<string, int> GetExistedParamsDic(string wsName)
        {
            var dic = new Dictionary<string, int>();

            var ws = ExcelDataUtil.CreateSheetIfNotExists(wsName);
            if (ws == null) return dic;

            Excel.Range range = ws.UsedRange;
            if (range.Rows.Count > 1 && range.Columns.Count > 1)
            {
                foreach (Excel.Range row in range.Rows)
                {
                    var curParam = row[1, 2].Value2;
                    if (curParam != null)
                        if (!dic.ContainsKey(curParam) && row.Row > FIRST_ROW)
                            dic.Add(curParam, row.Row);
                }
            }

            return dic;
        }
    }
}
