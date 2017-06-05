namespace EJTroubleReader
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {            
            this.btnEnter = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.btnExcel = new System.Windows.Forms.Button();
            this.lblPath = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.path_dir = new System.Windows.Forms.TextBox();
            this.excel_dir = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog2 = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // btnEnter
            // 
            this.btnEnter.Location = new System.Drawing.Point(77, 157);
            this.btnEnter.Name = "btnEnter";
            this.btnEnter.Size = new System.Drawing.Size(107, 55);
            this.btnEnter.TabIndex = 2;
            this.btnEnter.Text = "Enter";
            this.btnEnter.UseVisualStyleBackColor = true;
            this.btnEnter.Click += new System.EventHandler(this.btnEnter_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(339, 46);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(99, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Press to view Excel";
            // 
            // btnExcel
            // 
            this.btnExcel.Location = new System.Drawing.Point(342, 157);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(129, 55);
            this.btnExcel.TabIndex = 5;
            this.btnExcel.Text = "Extract to Excel";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // lblPath
            // 
            this.lblPath.AutoSize = true;
            this.lblPath.Location = new System.Drawing.Point(74, 46);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(86, 13);
            this.lblPath.TabIndex = 6;
            this.lblPath.Text = "Path or Directory";
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.ShowNewFolderButton = false;
            // 
            // path_dir
            // 
            this.path_dir.Location = new System.Drawing.Point(12, 81);
            this.path_dir.Multiline = true;
            this.path_dir.Name = "path_dir";
            this.path_dir.ReadOnly = true;
            this.path_dir.Size = new System.Drawing.Size(224, 52);
            this.path_dir.TabIndex = 7;
            this.path_dir.Text = "Browse your directory.";
            this.path_dir.Click += new System.EventHandler(this.path_dir_Click);
            // 
            // excel_dir
            // 
            this.excel_dir.Location = new System.Drawing.Point(293, 81);
            this.excel_dir.Multiline = true;
            this.excel_dir.Name = "excel_dir";
            this.excel_dir.ReadOnly = true;
            this.excel_dir.Size = new System.Drawing.Size(225, 52);
            this.excel_dir.TabIndex = 8;
            this.excel_dir.Text = "Browse your directory.";
            this.excel_dir.Click += new System.EventHandler(this.excel_dir_Click);
            // 
            // folderBrowserDialog2
            // 
            this.folderBrowserDialog2.ShowNewFolderButton = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(540, 273);
            this.Controls.Add(this.excel_dir);
            this.Controls.Add(this.path_dir);
            this.Controls.Add(this.lblPath);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnEnter);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnEnter;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.TextBox path_dir;
        private System.Windows.Forms.TextBox excel_dir;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog2;
    }
}

