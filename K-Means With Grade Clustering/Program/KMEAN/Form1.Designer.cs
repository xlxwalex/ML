namespace KMEAN
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
            this.LoadDataXLs = new System.Windows.Forms.Button();
            this.XlsDataSh = new System.Windows.Forms.ListView();
            this.Col1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Col2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Col3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.PlanBok = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Col4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Col5 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Col6 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Col7 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.LOF = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.kMEANBtn = new System.Windows.Forms.Button();
            this.OutPExBtn = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.XLsSaveDlg = new System.Windows.Forms.SaveFileDialog();
            this.SuspendLayout();
            // 
            // LoadDataXLs
            // 
            this.LoadDataXLs.Location = new System.Drawing.Point(810, 779);
            this.LoadDataXLs.Name = "LoadDataXLs";
            this.LoadDataXLs.Size = new System.Drawing.Size(173, 73);
            this.LoadDataXLs.TabIndex = 1;
            this.LoadDataXLs.Text = "加载Excel";
            this.LoadDataXLs.UseVisualStyleBackColor = true;
            this.LoadDataXLs.Click += new System.EventHandler(this.LoadDataXLs_Click);
            // 
            // XlsDataSh
            // 
            this.XlsDataSh.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Col1,
            this.Col2,
            this.Col3,
            this.PlanBok,
            this.Col4,
            this.Col5,
            this.Col6,
            this.Col7,
            this.LOF});
            this.XlsDataSh.Location = new System.Drawing.Point(12, 12);
            this.XlsDataSh.Name = "XlsDataSh";
            this.XlsDataSh.Size = new System.Drawing.Size(1161, 761);
            this.XlsDataSh.TabIndex = 2;
            this.XlsDataSh.UseCompatibleStateImageBehavior = false;
            this.XlsDataSh.View = System.Windows.Forms.View.Details;
            // 
            // Col1
            // 
            this.Col1.Text = "编号";
            this.Col1.Width = 45;
            // 
            // Col2
            // 
            this.Col2.Text = "学号";
            this.Col2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.Col2.Width = 63;
            // 
            // Col3
            // 
            this.Col3.Text = "姓名";
            this.Col3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // PlanBok
            // 
            this.PlanBok.Text = "计划书";
            // 
            // Col4
            // 
            this.Col4.Text = "评价";
            this.Col4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.Col4.Width = 55;
            // 
            // Col5
            // 
            this.Col5.Text = "报告";
            this.Col5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.Col5.Width = 55;
            // 
            // Col6
            // 
            this.Col6.Text = "互评";
            this.Col6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.Col6.Width = 55;
            // 
            // Col7
            // 
            this.Col7.Text = "记录得分";
            this.Col7.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // LOF
            // 
            this.LOF.Text = "LOF";
            this.LOF.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // kMEANBtn
            // 
            this.kMEANBtn.Enabled = false;
            this.kMEANBtn.Location = new System.Drawing.Point(1000, 779);
            this.kMEANBtn.Name = "kMEANBtn";
            this.kMEANBtn.Size = new System.Drawing.Size(173, 73);
            this.kMEANBtn.TabIndex = 3;
            this.kMEANBtn.Text = "LOF";
            this.kMEANBtn.UseVisualStyleBackColor = true;
            this.kMEANBtn.Click += new System.EventHandler(this.kMEANBtn_Click);
            // 
            // OutPExBtn
            // 
            this.OutPExBtn.Enabled = false;
            this.OutPExBtn.Location = new System.Drawing.Point(1000, 779);
            this.OutPExBtn.Name = "OutPExBtn";
            this.OutPExBtn.Size = new System.Drawing.Size(173, 73);
            this.OutPExBtn.TabIndex = 4;
            this.OutPExBtn.Text = "导出Excel";
            this.OutPExBtn.UseVisualStyleBackColor = true;
            this.OutPExBtn.Click += new System.EventHandler(this.OutPExBtn_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.BackColor = System.Drawing.SystemColors.Menu;
            this.richTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextBox1.Location = new System.Drawing.Point(1179, 12);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(263, 840);
            this.richTextBox1.TabIndex = 5;
            this.richTextBox1.Text = "准备就绪:\n";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1452, 860);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.kMEANBtn);
            this.Controls.Add(this.XlsDataSh);
            this.Controls.Add(this.LoadDataXLs);
            this.Controls.Add(this.OutPExBtn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "K-Means   Xlxw.Per";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button LoadDataXLs;
        private System.Windows.Forms.ListView XlsDataSh;
        private System.Windows.Forms.ColumnHeader Col1;
        private System.Windows.Forms.ColumnHeader Col2;
        private System.Windows.Forms.ColumnHeader Col3;
        private System.Windows.Forms.ColumnHeader Col4;
        private System.Windows.Forms.ColumnHeader Col5;
        private System.Windows.Forms.ColumnHeader Col6;
        private System.Windows.Forms.ColumnHeader Col7;
        private System.Windows.Forms.Button kMEANBtn;
        private System.Windows.Forms.Button OutPExBtn;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.ColumnHeader LOF;
        private System.Windows.Forms.ColumnHeader PlanBok;
        private System.Windows.Forms.SaveFileDialog XLsSaveDlg;
    }
}

