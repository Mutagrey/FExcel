using FExcel.FELoader.Model;
using FExcel.FELoader.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FExcel.FELoader.ViewModel
{
    public class SettingsViewModel
    {
        public string[] GroupBy { get { return Enum.GetNames(typeof(FileElementType)); } }
        public string SelectedGroupName { get; set; }   

        public SettingsViewModel()
        {
            SelectedGroupName = Properties.Settings.Default.GroupBy;
        }
    }
}
