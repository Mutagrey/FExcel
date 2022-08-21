using FExcel.FELoader.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FExcel.FELoader.Model
{
    public class LoadListModel
    {
        public string TableName { get; set; }
        public IList<FileElement> LoadListItems { get; private set; }
        public IList<Category> Categories { get; private set; }

        public LoadListModel(string tableName)
        {
            TableName = tableName;
        }

        public void UpdateData()
        {
            var table = ExcelDataUtil.GetListObject(TableName);
            if (table == null) return;

            var data = table.DataBodyRange;
            if (data == null) return;

            var loadList = new List<FileElement>();
            var categoryDic = new Dictionary<string, IList<FileElement>>();
            for (int i = 1; i <= data.Rows.Count; i++)
            {
                var curRowData = data.Rows[i].Value2;
                if (curRowData == null) continue; // ???
                
                double parseDouble = 0;
                int parseInt = 0;

                var element = new FileElement();
                element.Id = i;
                element.FilePath = curRowData[1, 2] == null ? "" : curRowData[1, 2].ToString();
                element.BookName = curRowData[1, 3] == null ? "" : curRowData[1, 3].ToString();
                element.BookStructure = curRowData[1, 4] == null ? "" : curRowData[1, 4].ToString();
                element.SheetName = curRowData[1, 5] == null ? "" : curRowData[1, 5].ToString();
                element.TemplateName = curRowData[1, 6] == null ? "" : curRowData[1, 6].ToString();
                element.ShiftYear = curRowData[1, 7] == null ? 0 : int.TryParse(curRowData[1, 11].ToString(), out parseInt) ? parseInt : 0;
                element.Group = curRowData[1, 8] == null ? "" : curRowData[1, 8].ToString();
                element.OG = curRowData[1, 9] == null ? "" : curRowData[1, 9].ToString();
                element.Mest = curRowData[1, 10] == null ? "" : curRowData[1, 10].ToString();
                element.MFSOKoeff = curRowData[1, 11] == null ? 1 : double.TryParse(curRowData[1, 11].ToString(), out parseDouble) ? parseDouble : 1;
                element.IsSelected = false;
                loadList.Add(element);

                // Группировка по категории
                var curGroupBy = FileElement.GetGroupBy(element);
                if (!categoryDic.ContainsKey(curGroupBy))
                    categoryDic.Add(curGroupBy, new List<FileElement>());
                categoryDic[curGroupBy].Add(element);

            }
            LoadListItems = loadList;

            Categories = categoryDic.Select(p => new Category(p.Key, p.Value.Count)).ToList();
        }

        public void AddFromFiles(string[] files)
        {

        }

    }
}
