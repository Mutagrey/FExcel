using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using FExcel.FELoader.ViewModel;
using FExcel.FELoader.Model;
using FExcel.FELoader.Utility;
using System.Xml.Linq;

namespace FExcel.FELoader.View
{
    [ComVisible(true)]
    public partial class FExcelLoaderUserControl : UserControl
    {
        private readonly LoadViewModel loadViewModel;
        private readonly FESettingsManager settingsManager;
        private SettingsForm settingsForm;
        public FExcelLoaderUserControl()
        {
            InitializeComponent();
            loadViewModel = new LoadViewModel();
            settingsManager = new FESettingsManager();
        }

        protected override void OnLoad(EventArgs e)
        {
            try
            {
                loadViewModelBindingSource.DataSource = loadViewModel;

                fExcelTableModelBindingSource.DataSource = loadViewModelBindingSource;
                fExcelTableModelBindingSource.DataMember = "LoadListModels";
                dataGridViewOG.AutoGenerateColumns = false;
                dataGridViewLoad.AutoGenerateColumns = false;
                dataGridViewOG.DataSource = fExcelTableModelBindingSource;
                dataGridViewOG.DataMember = "Categories";

                var firstTable = loadViewModel.LoadListModels.FirstOrDefault();
                if (firstTable != null)
                    firstTable.UpdateData();

                settingsManager.CreateOrUpdateSettingsTables();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void cmbFETables_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                var loadList = (LoadListModel)cmbFETables.SelectedValue;
                if (loadList != null)
                    loadList.UpdateData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void butRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                var loadList = (LoadListModel)cmbFETables.SelectedValue;
                if (loadList != null)
                    loadViewModel.UpdateTable(loadList);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void butAddList_Click(object sender, EventArgs e)
        {
            try
            {
                var newLoadLst = loadViewModel.AddNewTable(cmbFETables.Text);
                if (newLoadLst != null)
                    cmbFETables.SelectedItem = newLoadLst;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void butAddFromFiles_Click(object sender, EventArgs e)
        {
            try
            {
                var loadList = (LoadListModel)cmbFETables.SelectedValue;
                if (loadList == null) return;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                    loadList.AddFromFiles(openFileDialog.FileNames);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void dataGridViewOG_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridViewOG.SelectedRows.Count == 0) return;

                var selectedOGList = dataGridViewOG.SelectedRows.Cast<DataGridViewRow>().Select(p => p.Cells[0].Value).ToList();
                var loadList = (LoadListModel)cmbFETables.SelectedValue;
                if (loadList == null) return;

                var filteredItems = loadList.LoadListItems.Where(p => selectedOGList.Contains(FileElement.GetGroupBy(p))).ToList();
                dataGridViewLoad.DataSource = filteredItems;

                UpdateStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void butSettings_Click(object sender, EventArgs e)
        {
            try
            {
                if (settingsForm == null)
                    settingsForm = new SettingsForm();
                settingsForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void butAddCompare_Click(object sender, EventArgs e)
        {
            try
            {
                var loadList = (LoadListModel)cmbFETables.SelectedValue;
                CompareForm.CreateOrUpdate(loadList, settingsManager.ParamList);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void UpdateStatus()
        {
            var categoryName = "ОГ";
            labelCategoryStatus.Text = string.Format("{0} ({1}/{2})", categoryName, dataGridViewOG.SelectedRows.Count, dataGridViewOG.RowCount);
            var loadListName = "Список для загрузки";
            labelLoadListStatus.Text = string.Format("{0} ({1}/{2})", loadListName, dataGridViewLoad.SelectedRows.Count, dataGridViewLoad.RowCount);
            
            var imageIDCategory = dataGridViewOG.SelectedRows.Count == dataGridViewOG.RowCount ? 1 
                : dataGridViewOG.SelectedRows.Count < dataGridViewOG.RowCount ? 2 : 0;
            var imageIDLoadList = dataGridViewLoad.SelectedRows.Count == dataGridViewLoad.RowCount ? 1
                : dataGridViewLoad.SelectedRows.Count < dataGridViewLoad.RowCount ? 2 : 0; ;

            butSelectCategory.Image = imageListCheck.Images[imageIDCategory];
            butSelectLoadList.Image = imageListCheck.Images[imageIDLoadList];
        }

        private void butSelectCategory_Click(object sender, EventArgs e)
        {
            try
            {
                if (butSelectCategory.Checked)
                    dataGridViewOG.SelectAll();
                else dataGridViewOG.ClearSelection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void butSelectLoadList_Click(object sender, EventArgs e)
        {
            try
            {
                if (butSelectLoadList.Checked)
                    dataGridViewLoad.SelectAll();
                else dataGridViewLoad.ClearSelection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridViewLoad_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                UpdateStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void butLoad_Click(object sender, EventArgs e)
        {

        }
    }
}
