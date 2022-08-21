using FExcel.FELoader.Model;
using FExcel.FELoader.Utility;
using FExcel.FELoader.ViewModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FExcel.FELoader.View
{
    public partial class SettingsForm : Form
    {
        private SettingsViewModel settingsViewModel = new SettingsViewModel();
        public SettingsForm()
        {
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            settingsViewModelBindingSource.DataSource = settingsViewModel;
        }

        private void butOK_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.GroupBy = settingsViewModel.SelectedGroupName;
            Properties.Settings.Default.Save();
            this.Hide();
        }
    }
}
