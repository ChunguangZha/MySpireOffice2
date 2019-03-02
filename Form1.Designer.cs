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
            this.label1 = new System.Windows.Forms.Label();
            this.txtGroup = new System.Windows.Forms.TextBox();
            this.btnLoadSymbols = new System.Windows.Forms.Button();
            this.btnBuild3Table = new System.Windows.Forms.Button();
            this.txtPeopleInfoTablePath = new System.Windows.Forms.TextBox();
            this.btnLoadPeopleInfoTable = new System.Windows.Forms.Button();
            this.txtSrcTable5FilePath = new System.Windows.Forms.TextBox();
            this.btnLoadSrcTable5 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnFormatDate = new System.Windows.Forms.Button();
            this.btnLoad人口登记表 = new System.Windows.Forms.Button();
            this.txt人口登记表路径 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnBuildTable4
            // 
            this.btnBuildTable4.Location = new System.Drawing.Point(30, 367);
            this.btnBuildTable4.Name = "btnBuildTable4";
            this.btnBuildTable4.Size = new System.Drawing.Size(310, 23);
            this.btnBuildTable4.TabIndex = 0;
            this.btnBuildTable4.Text = "生成家庭人员调查表";
            this.btnBuildTable4.UseVisualStyleBackColor = true;
            this.btnBuildTable4.Click += new System.EventHandler(this.btnBuildTable4家庭人员调查表_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txt人口登记表路径);
            this.groupBox1.Controls.Add(this.btnLoad人口登记表);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txtGroup);
            this.groupBox1.Controls.Add(this.btnLoadSymbols);
            this.groupBox1.Controls.Add(this.btnBuild3Table);
            this.groupBox1.Controls.Add(this.btnBuildTable4);
            this.groupBox1.Location = new System.Drawing.Point(12, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(382, 411);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "生成家庭人员调查表";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(185, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(17, 12);
            this.label1.TabIndex = 7;
            this.label1.Text = "屯";
            // 
            // txtGroup
            // 
            this.txtGroup.Location = new System.Drawing.Point(79, 38);
            this.txtGroup.Name = "txtGroup";
            this.txtGroup.Size = new System.Drawing.Size(100, 21);
            this.txtGroup.TabIndex = 6;
            // 
            // btnLoadSymbols
            // 
            this.btnLoadSymbols.Location = new System.Drawing.Point(30, 249);
            this.btnLoadSymbols.Name = "btnLoadSymbols";
            this.btnLoadSymbols.Size = new System.Drawing.Size(310, 23);
            this.btnLoadSymbols.TabIndex = 5;
            this.btnLoadSymbols.Text = "导入符号";
            this.btnLoadSymbols.UseVisualStyleBackColor = true;
            this.btnLoadSymbols.Click += new System.EventHandler(this.btnLoadSymbols_Click);
            // 
            // btnBuild3Table
            // 
            this.btnBuild3Table.Location = new System.Drawing.Point(30, 338);
            this.btnBuild3Table.Name = "btnBuild3Table";
            this.btnBuild3Table.Size = new System.Drawing.Size(310, 23);
            this.btnBuild3Table.TabIndex = 3;
            this.btnBuild3Table.Text = "生成人口摸底调查表";
            this.btnBuild3Table.UseVisualStyleBackColor = true;
            this.btnBuild3Table.Click += new System.EventHandler(this.btnBuild3Table人口摸底调查表_Click);
            // 
            // txtPeopleInfoTablePath
            // 
            this.txtPeopleInfoTablePath.Location = new System.Drawing.Point(418, 242);
            this.txtPeopleInfoTablePath.Name = "txtPeopleInfoTablePath";
            this.txtPeopleInfoTablePath.ReadOnly = true;
            this.txtPeopleInfoTablePath.Size = new System.Drawing.Size(370, 21);
            this.txtPeopleInfoTablePath.TabIndex = 4;
            // 
            // btnLoadPeopleInfoTable
            // 
            this.btnLoadPeopleInfoTable.Location = new System.Drawing.Point(454, 213);
            this.btnLoadPeopleInfoTable.Name = "btnLoadPeopleInfoTable";
            this.btnLoadPeopleInfoTable.Size = new System.Drawing.Size(310, 23);
            this.btnLoadPeopleInfoTable.TabIndex = 3;
            this.btnLoadPeopleInfoTable.Text = "导入人口信息采集表";
            this.btnLoadPeopleInfoTable.UseVisualStyleBackColor = true;
            this.btnLoadPeopleInfoTable.Click += new System.EventHandler(this.btnLoadSrcTablePeopleInfo_Click);
            // 
            // txtSrcTable5FilePath
            // 
            this.txtSrcTable5FilePath.Location = new System.Drawing.Point(418, 149);
            this.txtSrcTable5FilePath.Name = "txtSrcTable5FilePath";
            this.txtSrcTable5FilePath.ReadOnly = true;
            this.txtSrcTable5FilePath.Size = new System.Drawing.Size(370, 21);
            this.txtSrcTable5FilePath.TabIndex = 2;
            // 
            // btnLoadSrcTable5
            // 
            this.btnLoadSrcTable5.Location = new System.Drawing.Point(454, 120);
            this.btnLoadSrcTable5.Name = "btnLoadSrcTable5";
            this.btnLoadSrcTable5.Size = new System.Drawing.Size(310, 23);
            this.btnLoadSrcTable5.TabIndex = 1;
            this.btnLoadSrcTable5.Text = "导入户籍信息表";
            this.btnLoadSrcTable5.UseVisualStyleBackColor = true;
            this.btnLoadSrcTable5.Click += new System.EventHandler(this.btnLoad户籍信息表_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnFormatDate
            // 
            this.btnFormatDate.Location = new System.Drawing.Point(524, 80);
            this.btnFormatDate.Name = "btnFormatDate";
            this.btnFormatDate.Size = new System.Drawing.Size(189, 23);
            this.btnFormatDate.TabIndex = 2;
            this.btnFormatDate.Text = "Format Date";
            this.btnFormatDate.UseVisualStyleBackColor = true;
            this.btnFormatDate.Click += new System.EventHandler(this.btnFormatDate_Click);
            // 
            // btnLoad人口登记表
            // 
            this.btnLoad人口登记表.Location = new System.Drawing.Point(30, 76);
            this.btnLoad人口登记表.Name = "btnLoad人口登记表";
            this.btnLoad人口登记表.Size = new System.Drawing.Size(310, 23);
            this.btnLoad人口登记表.TabIndex = 8;
            this.btnLoad人口登记表.Text = "导入人口登记表";
            this.btnLoad人口登记表.UseVisualStyleBackColor = true;
            this.btnLoad人口登记表.Click += new System.EventHandler(this.btnLoad人口登记表_Click);
            // 
            // txt人口登记表路径
            // 
            this.txt人口登记表路径.Location = new System.Drawing.Point(30, 106);
            this.txt人口登记表路径.Name = "txt人口登记表路径";
            this.txt人口登记表路径.ReadOnly = true;
            this.txt人口登记表路径.Size = new System.Drawing.Size(310, 21);
            this.txt人口登记表路径.TabIndex = 9;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnFormatDate);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnLoadPeopleInfoTable);
            this.Controls.Add(this.btnLoadSrcTable5);
            this.Controls.Add(this.txtPeopleInfoTablePath);
            this.Controls.Add(this.txtSrcTable5FilePath);
            this.Name = "Form1";
            this.Text = "村成员Excel生成";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Button btnBuildTable4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtSrcTable5FilePath;
        private System.Windows.Forms.Button btnLoadSrcTable5;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;

        #endregion

        private System.Windows.Forms.Button btnLoadPeopleInfoTable;
        private System.Windows.Forms.TextBox txtPeopleInfoTablePath;
        private System.Windows.Forms.Button btnFormatDate;
        private System.Windows.Forms.Button btnBuild3Table;
        private System.Windows.Forms.Button btnLoadSymbols;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtGroup;
        private System.Windows.Forms.TextBox txt人口登记表路径;
        private System.Windows.Forms.Button btnLoad人口登记表;
    }
}

