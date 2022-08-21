using FExcel.Properties;
using FExcel.FELoader.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace FExcel.FELoader.Utility
{
    public class FESettingsManager
    {
        const string SETTINGS_WS_NAME = "FESettings";
        const string TEMPLATES_NAME = "TemplatesFE_TB";
        const string PARAMS_NAME = "ParamsFE_TB";
        const string TEMPLATES_RANGE= "A4";
        const string PARAMS_RANGE = "F4";
        public IList<TemplateModel> TemplateList { get; private set; }
        public IList<ParamModel> ParamList { get; private set; }

        public void CreateOrUpdateSettingsTables()
        {
            var templatesTable = ExcelDataUtil.CreateExcelTableFromTxtString(TEMPLATES_NAME, 
                SETTINGS_WS_NAME, TEMPLATES_RANGE, Resources.TemplatesTable, false);
            var paramsTable = ExcelDataUtil.CreateExcelTableFromTxtString(PARAMS_NAME,
                SETTINGS_WS_NAME, PARAMS_RANGE, Resources.ParamsTable, false);

            TemplateList = ReadTemplatesModel(templatesTable);
            ParamList = ReadParamsModel(paramsTable, TemplateList);
        }

        private IList<TemplateModel> ReadTemplatesModel(Excel.ListObject excelTable)
        {
            IList<TemplateModel> templateModels = new List<TemplateModel>();

            if (excelTable.DataBodyRange == null) return templateModels;

            for (int i = 0; i < excelTable.DataBodyRange.Rows.Count; i++)
            {
                var curRowData = excelTable.DataBodyRange.Rows[i].Value2;
                var element = new TemplateModel();
                element.Id = i;
                element.Name = curRowData[1, 1] == null ? "" : curRowData[1, 1].ToString();
                element.Mask = curRowData[1, 2] == null ? "" : curRowData[1, 2].ToString();
                element.FirstCellAddress = curRowData[1, 3] == null ? "" : curRowData[1, 3].ToString();
                templateModels.Add(element);
            }

            return templateModels;
        }

        private IList<ParamModel> ReadParamsModel(Excel.ListObject excelTable, IList<TemplateModel> templateModels)
        {
            IList<ParamModel> paramModels = new List<ParamModel>();

            if (excelTable.DataBodyRange == null) return paramModels;

            for (int i = 1; i <= excelTable.DataBodyRange.Rows.Count; i++)
            {
                var curRowData = excelTable.DataBodyRange.Rows[i].Value2;
                var element = new ParamModel();
                element.Id = i;
                element.Name = curRowData[1, 2] == null ? "" : curRowData[1, 2].ToString();
                element.IsMFSO = curRowData[1, 3] == null ? false : curRowData[1, 3].ToString() == "1" ? true : false;
                element.IsSelected = curRowData[1, 4] == null ? false : curRowData[1, 4].ToString() == "1" ? true : false;
                var formulaDic = new Dictionary<string, string>(); 
                for (int j = 5; j <= excelTable.DataBodyRange.Columns.Count; j++)
                {
                    var key = excelTable.HeaderRowRange[j].Value2;
                    if (key != null)
                    {
                        if (!formulaDic.ContainsKey(key))
                            formulaDic.Add(key, curRowData[1, j] == null ? "" : curRowData[1, j].ToString());
                    }
                }
                element.Formula = formulaDic;
                element.RowID = 0;
                paramModels.Add(element);
            }

            return paramModels;
        }
    }
}
