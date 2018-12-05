namespace MySpireOffice2
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnBuildTable4 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtPeopleInfoTablePath = new System.Windows.Forms.TextBox();
            this.btnLoadPeopleInfoTable = new System.Windows.Forms.Button();
            this.txtSrcTable5FilePath = new System.Windows.Forms.TextBox();
            this.btnLoadSrcTable5 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnBuildTable4
            // 
            this.btnBuildTable4.Location = new System.Drawing.Point(30, 367);
            this.btnBuildTable4.Name = "btnBuildTable4";
            this.btnBuildTable4.Size = new System.Drawing.Size(310, 23);
            this.btnBuildTable4.TabIndex = 0;
            this.btnBuildTable4.Text = "生成";
            this.btnBuildTable4.UseVisualStyleBackColor = true;
            this.btnBuildTable4.Click += new System.EventHandler(this.btnBuildTable4_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtPeopleInfoTablePath);
            this.groupBox1.Controls.Add(this.btnLoadPeopleInfoTable);
            this.groupBox1.Controls.Add(this.txtSrcTable5FilePath);
            this.groupBox1.Controls.Add(this.btnLoadSrcTable5);
            this.groupBox1.Controls.Add(this.btnBuildTable4);
            this.groupBox1.Location = new System.Drawing.Point(12, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(382, 411);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "生成家庭人员调查表";
            // 
            // txtPeopleInfoTablePath
            // 
            this.txtPeopleInfoTablePath.Location = new System.Drawing.Point(6, 167);
            this.txtPeopleInfoTablePath.Name = "txtPeopleInfoTablePath";
            this.txtPeopleInfoTablePath.ReadOnly = true;
            this.txtPeopleInfoTablePath.Size = new System.Drawing.Size(370, 21);
            this.txtPeopleInfoTablePath.TabIndex = 4;
            // 
            // btnLoadPeopleInfoTable
            // 
            this.btnLoadPeopleInfoTable.Location = new System.Drawing.Point(30, 121);
            this.btnLoadPeopleInfoTable.Name = "btnLoadPeopleInfoTable";
            this.btnLoadPeopleInfoTable.Size = new System.Drawing.Size(310, 23);
            this.btnLoadPeopleInfoTable.TabIndex = 3;
            this.btnLoadPeopleInfoTable.Text = "导入人口信息采集表";
            this.btnLoadPeopleInfoTable.UseVisualStyleBackColor = true;
            this.btnLoadPeopleInfoTable.Click += new System.EventHandler(this.btnLoadSrcTablePeopleInfo_Click);
            // 
            // txtSrcTable5FilePath
            // 
            this.txtSrcTable5FilePath.Location = new System.Drawing.Point(6, 77);
            this.txtSrcTable5FilePath.Name = "txtSrcTable5FilePath";
            this.txtSrcTable5FilePath.ReadOnly = true;
            this.txtSrcTable5FilePath.Size = new System.Drawing.Size(370, 21);
            this.txtSrcTable5FilePath.TabIndex = 2;
            // 
            // btnLoadSrcTable5
            // 
            this.btnLoadSrcTable5.Location = new System.Drawing.Point(30, 37);
            this.btnLoadSrcTable5.Name = "btnLoadSrcTable5";
            this.btnLoadSrcTable5.Size = new System.Drawing.Size(310, 23);
            this.btnLoadSrcTable5.TabIndex = 1;
            this.btnLoadSrcTable5.Text = "导入原始表5";
            this.btnLoadSrcTable5.UseVisualStyleBackColor = true;
            this.btnLoadSrcTable5.Click += new System.EventHandler(this.btnLoadSrcTable5_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "村成员Excel生成";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.Button btnBuildTable4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtSrcTable5FilePath;
        private System.Windows.Forms.Button btnLoadSrcTable5;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;

        #endregion

        private System.Windows.Forms.Button btnLoadPeopleInfoTable;
        private System.Windows.Forms.TextBox txtPeopleInfoTablePath;
    }
}

