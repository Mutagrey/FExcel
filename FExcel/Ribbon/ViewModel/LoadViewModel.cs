using FExcel.Properties;
using FExcel.FELoader.Model;
using FExcel.FELoader.Utility;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FExcel.FELoader.ViewModel
{
    public class LoadViewModel: INotifyPropertyChanged
    {
        public IList<LoadListModel> LoadListModels { get; set; }
        
        public LoadViewModel()
        {
            UpdateFExcelTables();
        }

        public void UpdateFExcelTables()
        {
            try
            {
                var fexcelList = new List<LoadListModel>();
                var workbook = ExcelDataUtil.ActiveWorkbook;
                foreach (Excel.Worksheet worksheet in workbook.Sheets)
                    foreach (Excel.ListObject listObject in worksheet.ListObjects)
                    {
                        var fexcelTable = new LoadListModel(listObject.Name);
                        fexcelList.Add(fexcelTable);
                    }
                LoadListModels = fexcelList;
                OnPropertyChanged(nameof(LoadListModels));
            }
            catch 
            {

            }

        }

        public void UpdateTable(LoadListModel loadListModel)
        {
            var curLoadList = LoadListModels.Where(p => p.TableName == loadListModel.TableName).FirstOrDefault();
            if (curLoadList != null)
                curLoadList.UpdateData();
            OnPropertyChanged(nameof(LoadListModels));
        }

        public LoadListModel AddNewTable(string tableName)
        {
            var loadListTable = ExcelDataUtil.CreateExcelTableFromTxtString(tableName,
                "", "A4", Resources.LoadListTable, true);

            if (loadListTable == null) return null;

            var newLoadList = new LoadListModel(loadListTable.Name);
            //newLoadList.UpdateData();

            LoadListModels.Add(newLoadList);
            OnPropertyChanged(nameof(LoadListModels));
            return newLoadList;
        }

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        // This method is called by the Set accessor of each property.  
        // The CallerMemberName attribute that is applied to the optional propertyName  
        // parameter causes the property name of the caller to be substituted as an argument.  
        private void OnPropertyChanged([CallerMemberName] String propertyName = "")
        {
            if (PropertyChanged != null)
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }
}
