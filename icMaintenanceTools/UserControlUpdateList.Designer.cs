namespace ICApiAddin.icMaintenanceTools
{
    partial class UserControlUpdateList
    {
        /// <summary> 
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary> 
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を 
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.tableLayoutPanelBase = new System.Windows.Forms.TableLayoutPanel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tableLayoutPanelTabPage1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanelPage1_1 = new System.Windows.Forms.TableLayoutPanel();
            this.dataGridViewUpdateList = new System.Windows.Forms.DataGridView();
            this.UpdateDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.UpdateTitle = new System.Windows.Forms.DataGridViewLinkColumn();
            this.updateAbout = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.linkLabelIroncadUpdate = new System.Windows.Forms.LinkLabel();
            this.tableLayoutPanelBase.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tableLayoutPanelTabPage1.SuspendLayout();
            this.tableLayoutPanelPage1_1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewUpdateList)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanelBase
            // 
            this.tableLayoutPanelBase.ColumnCount = 1;
            this.tableLayoutPanelBase.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanelBase.Controls.Add(this.tabControl1, 0, 0);
            this.tableLayoutPanelBase.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanelBase.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanelBase.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tableLayoutPanelBase.Name = "tableLayoutPanelBase";
            this.tableLayoutPanelBase.RowCount = 1;
            this.tableLayoutPanelBase.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanelBase.Size = new System.Drawing.Size(1119, 985);
            this.tableLayoutPanelBase.TabIndex = 2;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(4, 5);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1111, 975);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.tableLayoutPanelTabPage1);
            this.tabPage1.Location = new System.Drawing.Point(4, 28);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage1.Size = new System.Drawing.Size(1103, 943);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "IRONCAD更新プログラム情報";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tableLayoutPanelTabPage1
            // 
            this.tableLayoutPanelTabPage1.ColumnCount = 1;
            this.tableLayoutPanelTabPage1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanelTabPage1.Controls.Add(this.tableLayoutPanelPage1_1, 0, 0);
            this.tableLayoutPanelTabPage1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanelTabPage1.Location = new System.Drawing.Point(4, 5);
            this.tableLayoutPanelTabPage1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tableLayoutPanelTabPage1.Name = "tableLayoutPanelTabPage1";
            this.tableLayoutPanelTabPage1.RowCount = 1;
            this.tableLayoutPanelTabPage1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanelTabPage1.Size = new System.Drawing.Size(1095, 933);
            this.tableLayoutPanelTabPage1.TabIndex = 0;
            // 
            // tableLayoutPanelPage1_1
            // 
            this.tableLayoutPanelPage1_1.ColumnCount = 1;
            this.tableLayoutPanelPage1_1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanelPage1_1.Controls.Add(this.dataGridViewUpdateList, 0, 1);
            this.tableLayoutPanelPage1_1.Controls.Add(this.linkLabelIroncadUpdate, 0, 0);
            this.tableLayoutPanelPage1_1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanelPage1_1.Location = new System.Drawing.Point(15, 15);
            this.tableLayoutPanelPage1_1.Margin = new System.Windows.Forms.Padding(15);
            this.tableLayoutPanelPage1_1.Name = "tableLayoutPanelPage1_1";
            this.tableLayoutPanelPage1_1.RowCount = 2;
            this.tableLayoutPanelPage1_1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanelPage1_1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanelPage1_1.Size = new System.Drawing.Size(1065, 903);
            this.tableLayoutPanelPage1_1.TabIndex = 0;
            // 
            // dataGridViewUpdateList
            // 
            this.dataGridViewUpdateList.AllowUserToAddRows = false;
            this.dataGridViewUpdateList.AllowUserToDeleteRows = false;
            this.dataGridViewUpdateList.AllowUserToResizeRows = false;
            this.dataGridViewUpdateList.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridViewUpdateList.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dataGridViewUpdateList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewUpdateList.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.UpdateDate,
            this.UpdateTitle,
            this.updateAbout});
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewUpdateList.DefaultCellStyle = dataGridViewCellStyle4;
            this.dataGridViewUpdateList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewUpdateList.Location = new System.Drawing.Point(4, 24);
            this.dataGridViewUpdateList.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dataGridViewUpdateList.Name = "dataGridViewUpdateList";
            this.dataGridViewUpdateList.RowTemplate.Height = 65;
            this.dataGridViewUpdateList.RowTemplate.ReadOnly = true;
            this.dataGridViewUpdateList.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewUpdateList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewUpdateList.Size = new System.Drawing.Size(1057, 874);
            this.dataGridViewUpdateList.TabIndex = 9;
            this.dataGridViewUpdateList.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewUpdateList_CellContentClick);
            // 
            // UpdateDate
            // 
            this.UpdateDate.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Meiryo UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.UpdateDate.DefaultCellStyle = dataGridViewCellStyle1;
            this.UpdateDate.HeaderText = "リリース日";
            this.UpdateDate.Name = "UpdateDate";
            this.UpdateDate.Width = 94;
            // 
            // UpdateTitle
            // 
            this.UpdateTitle.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Meiryo UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.UpdateTitle.DefaultCellStyle = dataGridViewCellStyle2;
            this.UpdateTitle.HeaderText = "更新プログラム";
            this.UpdateTitle.LinkColor = System.Drawing.Color.DodgerBlue;
            this.UpdateTitle.Name = "UpdateTitle";
            this.UpdateTitle.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.UpdateTitle.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.UpdateTitle.VisitedLinkColor = System.Drawing.Color.DodgerBlue;
            this.UpdateTitle.Width = 124;
            // 
            // updateAbout
            // 
            this.updateAbout.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.updateAbout.DefaultCellStyle = dataGridViewCellStyle3;
            this.updateAbout.HeaderText = "概要";
            this.updateAbout.Name = "updateAbout";
            // 
            // linkLabelIroncadUpdate
            // 
            this.linkLabelIroncadUpdate.AutoSize = true;
            this.linkLabelIroncadUpdate.Dock = System.Windows.Forms.DockStyle.Left;
            this.linkLabelIroncadUpdate.Location = new System.Drawing.Point(4, 0);
            this.linkLabelIroncadUpdate.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.linkLabelIroncadUpdate.Name = "linkLabelIroncadUpdate";
            this.linkLabelIroncadUpdate.Size = new System.Drawing.Size(208, 19);
            this.linkLabelIroncadUpdate.TabIndex = 10;
            this.linkLabelIroncadUpdate.TabStop = true;
            this.linkLabelIroncadUpdate.Text = "IRONCAD公式サイトはこちらから";
            this.linkLabelIroncadUpdate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelIroncadUpdate_LinkClicked);
            // 
            // UserControlUpdateList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.tableLayoutPanelBase);
            this.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "UserControlUpdateList";
            this.Size = new System.Drawing.Size(1119, 985);
            this.VisibleChanged += new System.EventHandler(this.UserControlUpdateList_VisibleChanged);
            this.tableLayoutPanelBase.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tableLayoutPanelTabPage1.ResumeLayout(false);
            this.tableLayoutPanelPage1_1.ResumeLayout(false);
            this.tableLayoutPanelPage1_1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewUpdateList)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanelBase;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanelTabPage1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanelPage1_1;
        private System.Windows.Forms.DataGridView dataGridViewUpdateList;
        private System.Windows.Forms.DataGridViewTextBoxColumn UpdateDate;
        private System.Windows.Forms.DataGridViewLinkColumn UpdateTitle;
        private System.Windows.Forms.DataGridViewTextBoxColumn updateAbout;
        private System.Windows.Forms.LinkLabel linkLabelIroncadUpdate;
    }
}
