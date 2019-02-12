namespace WindowsFormsApp1
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
            this.dgv_Message = new System.Windows.Forms.DataGridView();
            this.btn_Remove = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.TextBoxProcess = new System.Windows.Forms.TextBox();
            this.btn_ToExcel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Message)).BeginInit();
            this.SuspendLayout();
            // 
            // dgv_Message
            // 
            this.dgv_Message.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_Message.Location = new System.Drawing.Point(12, 27);
            this.dgv_Message.Name = "dgv_Message";
            this.dgv_Message.RowTemplate.Height = 27;
            this.dgv_Message.Size = new System.Drawing.Size(490, 189);
            this.dgv_Message.TabIndex = 0;
            // 
            // btn_Remove
            // 
            this.btn_Remove.Font = new System.Drawing.Font("宋体", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_Remove.Location = new System.Drawing.Point(104, 267);
            this.btn_Remove.Name = "btn_Remove";
            this.btn_Remove.Size = new System.Drawing.Size(247, 65);
            this.btn_Remove.TabIndex = 1;
            this.btn_Remove.Text = "删除选定列";
            this.btn_Remove.UseVisualStyleBackColor = true;
            this.btn_Remove.Click += new System.EventHandler(this.btn_Remove_Click);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("宋体", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.Location = new System.Drawing.Point(453, 267);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(255, 65);
            this.button1.TabIndex = 2;
            this.button1.Text = "导出到Word文档";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // TextBoxProcess
            // 
            this.TextBoxProcess.Location = new System.Drawing.Point(541, 45);
            this.TextBoxProcess.Multiline = true;
            this.TextBoxProcess.Name = "TextBoxProcess";
            this.TextBoxProcess.Size = new System.Drawing.Size(263, 171);
            this.TextBoxProcess.TabIndex = 3;
            this.TextBoxProcess.Visible = false;
            // 
            // btn_ToExcel
            // 
            this.btn_ToExcel.Font = new System.Drawing.Font("宋体", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_ToExcel.Location = new System.Drawing.Point(453, 350);
            this.btn_ToExcel.Name = "btn_ToExcel";
            this.btn_ToExcel.Size = new System.Drawing.Size(255, 55);
            this.btn_ToExcel.TabIndex = 4;
            this.btn_ToExcel.Text = "导出到Excel文件";
            this.btn_ToExcel.UseVisualStyleBackColor = true;
            this.btn_ToExcel.Click += new System.EventHandler(this.btn_ToExcel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(850, 450);
            this.Controls.Add(this.btn_ToExcel);
            this.Controls.Add(this.TextBoxProcess);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btn_Remove);
            this.Controls.Add(this.dgv_Message);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Message)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgv_Message;
        private System.Windows.Forms.Button btn_Remove;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox TextBoxProcess;
        private System.Windows.Forms.Button btn_ToExcel;
    }
}

