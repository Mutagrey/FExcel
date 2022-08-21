using System;

namespace FExcel.FELoader.View
{
    partial class FExcelLoaderUserControl
    {
        /// <summary> 
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary> 
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FExcelLoaderUserControl));
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.loadViewModelBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.fExcelTableModelBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.dataGridViewOG = new System.Windows.Forms.DataGridView();
            this.dataGridViewLoad = new System.Windows.Forms.DataGridView();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.butSelectCategory = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.labelCategoryStatus = new System.Windows.Forms.ToolStripLabel();
            this.butSelectLoadList = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton2 = new System.Windows.Forms.ToolStripSeparator();
            this.labelLoadListStatus = new System.Windows.Forms.ToolStripLabel();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.toolStrip2 = new System.Windows.Forms.ToolStrip();
            this.butAddList = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.butAddFromFiles = new System.Windows.Forms.ToolStripButton();
            this.butRefresh = new System.Windows.Forms.ToolStripButton();
            this.butAddCompare = new System.Windows.Forms.ToolStripButton();
            this.butSettings = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.cmbFETables = new System.Windows.Forms.ComboBox();
            this.imageListCheck = new System.Windows.Forms.ImageList(this.components);
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.idDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.filePathDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bookNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bookStructureDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.sheetNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.templateNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.shiftYearDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.oGDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mestDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mFSOKoeffDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.isSelectedDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.categoryNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.categoryCountDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.categoryDescriptionDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.butLoad = new System.Windows.Forms.ToolStripButton();
            ((System.ComponentModel.ISupportInitialize)(this.loadViewModelBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.fExcelTableModelBindingSource)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewOG)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewLoad)).BeginInit();
            this.toolStrip1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.toolStrip2.SuspendLayout();
            this.SuspendLayout();
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.AutoSize = true;
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(560, 0);
            this.flowLayoutPanel1.TabIndex = 7;
            // 
            // loadViewModelBindingSource
            // 
            this.loadViewModelBindingSource.DataSource = typeof(FExcel.FELoader.ViewModel.LoadViewModel);
            // 
            // fExcelTableModelBindingSource
            // 
            this.fExcelTableModelBindingSource.AllowNew = true;
            this.fExcelTableModelBindingSource.DataMember = "LoadListModels";
            this.fExcelTableModelBindingSource.DataSource = this.loadViewModelBindingSource;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.splitContainer1, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.toolStrip1, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(560, 770);
            this.tableLayoutPanel1.TabIndex = 8;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(3, 79);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.dataGridViewOG);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.dataGridViewLoad);
            this.splitContainer1.Size = new System.Drawing.Size(554, 688);
            this.splitContainer1.SplitterDistance = 184;
            this.splitContainer1.TabIndex = 0;
            // 
            // dataGridViewOG
            // 
            this.dataGridViewOG.AllowUserToAddRows = false;
            this.dataGridViewOG.AllowUserToDeleteRows = false;
            this.dataGridViewOG.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(239)))), ((int)(((byte)(249)))));
            this.dataGridViewOG.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridViewOG.AutoGenerateColumns = false;
            this.dataGridViewOG.BackgroundColor = System.Drawing.Color.Gainsboro;
            this.dataGridViewOG.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.RaisedHorizontal;
            this.dataGridViewOG.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.dataGridViewOG.ColumnHeadersHeight = 30;
            this.dataGridViewOG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridViewOG.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.categoryNameDataGridViewTextBoxColumn,
            this.categoryCountDataGridViewTextBoxColumn,
            this.categoryDescriptionDataGridViewTextBoxColumn});
            this.dataGridViewOG.DataMember = "Categories";
            this.dataGridViewOG.DataSource = this.fExcelTableModelBindingSource;
            this.dataGridViewOG.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewOG.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewOG.Name = "dataGridViewOG";
            this.dataGridViewOG.ReadOnly = true;
            this.dataGridViewOG.RowHeadersVisible = false;
            this.dataGridViewOG.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewOG.Size = new System.Drawing.Size(184, 688);
            this.dataGridViewOG.TabIndex = 6;
            this.dataGridViewOG.SelectionChanged += new System.EventHandler(this.dataGridViewOG_SelectionChanged);
            // 
            // dataGridViewLoad
            // 
            this.dataGridViewLoad.AllowUserToAddRows = false;
            this.dataGridViewLoad.AllowUserToDeleteRows = false;
            this.dataGridViewLoad.AllowUserToResizeRows = false;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(239)))), ((int)(((byte)(249)))));
            this.dataGridViewLoad.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridViewLoad.AutoGenerateColumns = false;
            this.dataGridViewLoad.BackgroundColor = System.Drawing.Color.Gainsboro;
            this.dataGridViewLoad.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.RaisedHorizontal;
            this.dataGridViewLoad.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.dataGridViewLoad.ColumnHeadersHeight = 30;
            this.dataGridViewLoad.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridViewLoad.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.idDataGridViewTextBoxColumn,
            this.filePathDataGridViewTextBoxColumn,
            this.bookNameDataGridViewTextBoxColumn,
            this.bookStructureDataGridViewTextBoxColumn,
            this.sheetNameDataGridViewTextBoxColumn,
            this.templateNameDataGridViewTextBoxColumn,
            this.shiftYearDataGridViewTextBoxColumn,
            this.groupDataGridViewTextBoxColumn,
            this.oGDataGridViewTextBoxColumn,
            this.mestDataGridViewTextBoxColumn,
            this.mFSOKoeffDataGridViewTextBoxColumn,
            this.isSelectedDataGridViewCheckBoxColumn});
            this.dataGridViewLoad.DataMember = "LoadListItems";
            this.dataGridViewLoad.DataSource = this.fExcelTableModelBindingSource;
            this.dataGridViewLoad.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewLoad.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewLoad.Name = "dataGridViewLoad";
            this.dataGridViewLoad.ReadOnly = true;
            this.dataGridViewLoad.RowHeadersVisible = false;
            this.dataGridViewLoad.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewLoad.Size = new System.Drawing.Size(366, 688);
            this.dataGridViewLoad.TabIndex = 5;
            this.dataGridViewLoad.SelectionChanged += new System.EventHandler(this.dataGridViewLoad_SelectionChanged);
            // 
            // toolStrip1
            // 
            this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.butSelectCategory,
            this.toolStripSeparator2,
            this.labelCategoryStatus,
            this.butSelectLoadList,
            this.toolStripButton2,
            this.labelLoadListStatus});
            this.toolStrip1.Location = new System.Drawing.Point(0, 51);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(560, 25);
            this.toolStrip1.TabIndex = 1;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // butSelectCategory
            // 
            this.butSelectCategory.CheckOnClick = true;
            this.butSelectCategory.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.butSelectCategory.Image = ((System.Drawing.Image)(resources.GetObject("butSelectCategory.Image")));
            this.butSelectCategory.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.butSelectCategory.Name = "butSelectCategory";
            this.butSelectCategory.Size = new System.Drawing.Size(23, 22);
            this.butSelectCategory.Text = "Выделить категорию";
            this.butSelectCategory.Click += new System.EventHandler(this.butSelectCategory_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // labelCategoryStatus
            // 
            this.labelCategoryStatus.Name = "labelCategoryStatus";
            this.labelCategoryStatus.Size = new System.Drawing.Size(16, 22);
            this.labelCategoryStatus.Text = "...";
            // 
            // butSelectLoadList
            // 
            this.butSelectLoadList.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.butSelectLoadList.CheckOnClick = true;
            this.butSelectLoadList.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.butSelectLoadList.Image = ((System.Drawing.Image)(resources.GetObject("butSelectLoadList.Image")));
            this.butSelectLoadList.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.butSelectLoadList.Name = "butSelectLoadList";
            this.butSelectLoadList.Size = new System.Drawing.Size(23, 22);
            this.butSelectLoadList.Text = "Выделить список для загрузки";
            this.butSelectLoadList.Click += new System.EventHandler(this.butSelectLoadList_Click);
            // 
            // toolStripButton2
            // 
            this.toolStripButton2.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.toolStripButton2.Name = "toolStripButton2";
            this.toolStripButton2.Size = new System.Drawing.Size(6, 25);
            // 
            // labelLoadListStatus
            // 
            this.labelLoadListStatus.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.labelLoadListStatus.Name = "labelLoadListStatus";
            this.labelLoadListStatus.Size = new System.Drawing.Size(16, 22);
            this.labelLoadListStatus.Text = "...";
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.AutoSize = true;
            this.tableLayoutPanel2.ColumnCount = 2;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 27.07581F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 72.92419F));
            this.tableLayoutPanel2.Controls.Add(this.toolStrip2, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.cmbFETables, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(554, 45);
            this.tableLayoutPanel2.TabIndex = 2;
            // 
            // toolStrip2
            // 
            this.toolStrip2.AutoSize = false;
            this.toolStrip2.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.butAddList,
            this.toolStripSeparator1,
            this.butAddFromFiles,
            this.butRefresh,
            this.butAddCompare,
            this.butSettings,
            this.toolStripSeparator3,
            this.butLoad});
            this.toolStrip2.Location = new System.Drawing.Point(149, 0);
            this.toolStrip2.Name = "toolStrip2";
            this.toolStrip2.Size = new System.Drawing.Size(405, 45);
            this.toolStrip2.TabIndex = 7;
            this.toolStrip2.Text = "MainMenu";
            // 
            // butAddList
            // 
            this.butAddList.Image = ((System.Drawing.Image)(resources.GetObject("butAddList.Image")));
            this.butAddList.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.butAddList.Name = "butAddList";
            this.butAddList.Size = new System.Drawing.Size(76, 42);
            this.butAddList.Text = "Add New";
            this.butAddList.Click += new System.EventHandler(this.butAddList_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 45);
            // 
            // butAddFromFiles
            // 
            this.butAddFromFiles.Image = ((System.Drawing.Image)(resources.GetObject("butAddFromFiles.Image")));
            this.butAddFromFiles.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.butAddFromFiles.Name = "butAddFromFiles";
            this.butAddFromFiles.Size = new System.Drawing.Size(106, 42);
            this.butAddFromFiles.Text = "Add From Files";
            this.butAddFromFiles.Click += new System.EventHandler(this.butAddFromFiles_Click);
            // 
            // butRefresh
            // 
            this.butRefresh.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.butRefresh.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.butRefresh.Image = ((System.Drawing.Image)(resources.GetObject("butRefresh.Image")));
            this.butRefresh.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.butRefresh.Name = "butRefresh";
            this.butRefresh.Size = new System.Drawing.Size(23, 42);
            this.butRefresh.Text = "toolStripButton5";
            this.butRefresh.Click += new System.EventHandler(this.butRefresh_Click);
            // 
            // butAddCompare
            // 
            this.butAddCompare.Image = ((System.Drawing.Image)(resources.GetObject("butAddCompare.Image")));
            this.butAddCompare.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.butAddCompare.Name = "butAddCompare";
            this.butAddCompare.Size = new System.Drawing.Size(68, 42);
            this.butAddCompare.Text = "Add Var";
            this.butAddCompare.Click += new System.EventHandler(this.butAddCompare_Click);
            // 
            // butSettings
            // 
            this.butSettings.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.butSettings.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.butSettings.Image = ((System.Drawing.Image)(resources.GetObject("butSettings.Image")));
            this.butSettings.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.butSettings.Name = "butSettings";
            this.butSettings.Size = new System.Drawing.Size(23, 42);
            this.butSettings.Text = "toolStripButton5";
            this.butSettings.Click += new System.EventHandler(this.butSettings_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 45);
            // 
            // cmbFETables
            // 
            this.cmbFETables.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFETables.DataSource = this.fExcelTableModelBindingSource;
            this.cmbFETables.DisplayMember = "TableName";
            this.cmbFETables.FormattingEnabled = true;
            this.cmbFETables.Location = new System.Drawing.Point(3, 12);
            this.cmbFETables.Name = "cmbFETables";
            this.cmbFETables.Size = new System.Drawing.Size(143, 21);
            this.cmbFETables.TabIndex = 4;
            this.cmbFETables.SelectedIndexChanged += new System.EventHandler(this.cmbFETables_SelectedIndexChanged);
            // 
            // imageListCheck
            // 
            this.imageListCheck.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListCheck.ImageStream")));
            this.imageListCheck.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListCheck.Images.SetKeyName(0, "CheckBoxUnchecked.png");
            this.imageListCheck.Images.SetKeyName(1, "CheckBoxChecked.png");
            this.imageListCheck.Images.SetKeyName(2, "CheckBoxMixed.png");
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "Excel файлы|*.xls*|Все файлы|*.*";
            this.openFileDialog.Multiselect = true;
            this.openFileDialog.Title = "Excel Files";
            // 
            // idDataGridViewTextBoxColumn
            // 
            this.idDataGridViewTextBoxColumn.DataPropertyName = "Id";
            this.idDataGridViewTextBoxColumn.HeaderText = "Id";
            this.idDataGridViewTextBoxColumn.Name = "idDataGridViewTextBoxColumn";
            this.idDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // filePathDataGridViewTextBoxColumn
            // 
            this.filePathDataGridViewTextBoxColumn.DataPropertyName = "FilePath";
            this.filePathDataGridViewTextBoxColumn.HeaderText = "FilePath";
            this.filePathDataGridViewTextBoxColumn.Name = "filePathDataGridViewTextBoxColumn";
            this.filePathDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // bookNameDataGridViewTextBoxColumn
            // 
            this.bookNameDataGridViewTextBoxColumn.DataPropertyName = "BookName";
            this.bookNameDataGridViewTextBoxColumn.HeaderText = "BookName";
            this.bookNameDataGridViewTextBoxColumn.Name = "bookNameDataGridViewTextBoxColumn";
            this.bookNameDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // bookStructureDataGridViewTextBoxColumn
            // 
            this.bookStructureDataGridViewTextBoxColumn.DataPropertyName = "BookStructure";
            this.bookStructureDataGridViewTextBoxColumn.HeaderText = "BookStructure";
            this.bookStructureDataGridViewTextBoxColumn.Name = "bookStructureDataGridViewTextBoxColumn";
            this.bookStructureDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // sheetNameDataGridViewTextBoxColumn
            // 
            this.sheetNameDataGridViewTextBoxColumn.DataPropertyName = "SheetName";
            this.sheetNameDataGridViewTextBoxColumn.HeaderText = "SheetName";
            this.sheetNameDataGridViewTextBoxColumn.Name = "sheetNameDataGridViewTextBoxColumn";
            this.sheetNameDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // templateNameDataGridViewTextBoxColumn
            // 
            this.templateNameDataGridViewTextBoxColumn.DataPropertyName = "TemplateName";
            this.templateNameDataGridViewTextBoxColumn.HeaderText = "TemplateName";
            this.templateNameDataGridViewTextBoxColumn.Name = "templateNameDataGridViewTextBoxColumn";
            this.templateNameDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // shiftYearDataGridViewTextBoxColumn
            // 
            this.shiftYearDataGridViewTextBoxColumn.DataPropertyName = "ShiftYear";
            this.shiftYearDataGridViewTextBoxColumn.HeaderText = "ShiftYear";
            this.shiftYearDataGridViewTextBoxColumn.Name = "shiftYearDataGridViewTextBoxColumn";
            this.shiftYearDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // groupDataGridViewTextBoxColumn
            // 
            this.groupDataGridViewTextBoxColumn.DataPropertyName = "Group";
            this.groupDataGridViewTextBoxColumn.HeaderText = "Group";
            this.groupDataGridViewTextBoxColumn.Name = "groupDataGridViewTextBoxColumn";
            this.groupDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // oGDataGridViewTextBoxColumn
            // 
            this.oGDataGridViewTextBoxColumn.DataPropertyName = "OG";
            this.oGDataGridViewTextBoxColumn.HeaderText = "OG";
            this.oGDataGridViewTextBoxColumn.Name = "oGDataGridViewTextBoxColumn";
            this.oGDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // mestDataGridViewTextBoxColumn
            // 
            this.mestDataGridViewTextBoxColumn.DataPropertyName = "Mest";
            this.mestDataGridViewTextBoxColumn.HeaderText = "Mest";
            this.mestDataGridViewTextBoxColumn.Name = "mestDataGridViewTextBoxColumn";
            this.mestDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // mFSOKoeffDataGridViewTextBoxColumn
            // 
            this.mFSOKoeffDataGridViewTextBoxColumn.DataPropertyName = "MFSOKoeff";
            this.mFSOKoeffDataGridViewTextBoxColumn.HeaderText = "MFSOKoeff";
            this.mFSOKoeffDataGridViewTextBoxColumn.Name = "mFSOKoeffDataGridViewTextBoxColumn";
            this.mFSOKoeffDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // isSelectedDataGridViewCheckBoxColumn
            // 
            this.isSelectedDataGridViewCheckBoxColumn.DataPropertyName = "IsSelected";
            this.isSelectedDataGridViewCheckBoxColumn.HeaderText = "IsSelected";
            this.isSelectedDataGridViewCheckBoxColumn.Name = "isSelectedDataGridViewCheckBoxColumn";
            this.isSelectedDataGridViewCheckBoxColumn.ReadOnly = true;
            // 
            // categoryNameDataGridViewTextBoxColumn
            // 
            this.categoryNameDataGridViewTextBoxColumn.DataPropertyName = "CategoryName";
            this.categoryNameDataGridViewTextBoxColumn.HeaderText = "CategoryName";
            this.categoryNameDataGridViewTextBoxColumn.Name = "categoryNameDataGridViewTextBoxColumn";
            this.categoryNameDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // categoryCountDataGridViewTextBoxColumn
            // 
            this.categoryCountDataGridViewTextBoxColumn.DataPropertyName = "CategoryCount";
            this.categoryCountDataGridViewTextBoxColumn.HeaderText = "CategoryCount";
            this.categoryCountDataGridViewTextBoxColumn.Name = "categoryCountDataGridViewTextBoxColumn";
            this.categoryCountDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // categoryDescriptionDataGridViewTextBoxColumn
            // 
            this.categoryDescriptionDataGridViewTextBoxColumn.DataPropertyName = "CategoryDescription";
            this.categoryDescriptionDataGridViewTextBoxColumn.HeaderText = "CategoryDescription";
            this.categoryDescriptionDataGridViewTextBoxColumn.Name = "categoryDescriptionDataGridViewTextBoxColumn";
            this.categoryDescriptionDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // butLoad
            // 
            this.butLoad.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.butLoad.Image = ((System.Drawing.Image)(resources.GetObject("butLoad.Image")));
            this.butLoad.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.butLoad.Name = "butLoad";
            this.butLoad.Size = new System.Drawing.Size(53, 42);
            this.butLoad.Text = "Load";
            // 
            // FExcelLoaderUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.flowLayoutPanel1);
            this.Name = "FExcelLoaderUserControl";
            this.Size = new System.Drawing.Size(560, 770);
            ((System.ComponentModel.ISupportInitialize)(this.loadViewModelBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.fExcelTableModelBindingSource)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewOG)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewLoad)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.toolStrip2.ResumeLayout(false);
            this.toolStrip2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.DataGridViewTextBoxColumn paramsDicDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn templatesDicDataGridViewTextBoxColumn;
        private System.Windows.Forms.BindingSource loadViewModelBindingSource;
        private System.Windows.Forms.BindingSource fExcelTableModelBindingSource;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.DataGridView dataGridViewOG;
        private System.Windows.Forms.DataGridView dataGridViewLoad;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton butSelectCategory;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripLabel labelCategoryStatus;
        private System.Windows.Forms.ToolStripButton butSelectLoadList;
        private System.Windows.Forms.ToolStripSeparator toolStripButton2;
        private System.Windows.Forms.ToolStripLabel labelLoadListStatus;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.ToolStrip toolStrip2;
        private System.Windows.Forms.ToolStripButton butAddList;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton butAddFromFiles;
        private System.Windows.Forms.ToolStripButton butRefresh;
        private System.Windows.Forms.ComboBox cmbFETables;
        private System.Windows.Forms.ImageList imageListCheck;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn7;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn8;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn9;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn10;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn11;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn12;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn13;
        private System.Windows.Forms.ToolStripButton butAddCompare;
        private System.Windows.Forms.ToolStripButton butSettings;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.DataGridViewTextBoxColumn categoryNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn categoryCountDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn categoryDescriptionDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn idDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn filePathDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn bookNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn bookStructureDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn sheetNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn templateNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn shiftYearDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn groupDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn oGDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn mestDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn mFSOKoeffDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn isSelectedDataGridViewCheckBoxColumn;
        private System.Windows.Forms.ToolStripButton butLoad;
    }
}
