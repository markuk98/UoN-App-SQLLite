namespace UoN_App_SQLLite
{
    partial class FormDataRecovery
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
            this.dataGridViewDBFiles = new System.Windows.Forms.DataGridView();
            this.C1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.C2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.C3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDBFiles)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridViewDBFiles
            // 
            this.dataGridViewDBFiles.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridViewDBFiles.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewDBFiles.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.C1,
            this.C2,
            this.C3});
            this.dataGridViewDBFiles.Location = new System.Drawing.Point(12, 12);
            this.dataGridViewDBFiles.Name = "dataGridViewDBFiles";
            this.dataGridViewDBFiles.RowTemplate.Height = 24;
            this.dataGridViewDBFiles.Size = new System.Drawing.Size(316, 257);
            this.dataGridViewDBFiles.TabIndex = 0;
            this.dataGridViewDBFiles.DoubleClick += new System.EventHandler(this.dataGridViewDBFiles_DoubleClick);
            // 
            // C1
            // 
            this.C1.HeaderText = "File Location";
            this.C1.Name = "C1";
            this.C1.ReadOnly = true;
            this.C1.Visible = false;
            // 
            // C2
            // 
            this.C2.HeaderText = "Last used";
            this.C2.Name = "C2";
            this.C2.ReadOnly = true;
            // 
            // C3
            // 
            this.C3.HeaderText = "New File Name";
            this.C3.Name = "C3";
            this.C3.Visible = false;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(340, 281);
            this.Controls.Add(this.dataGridViewDBFiles);
            this.Name = "Form2";
            this.Text = "Database Recovery";
            this.Load += new System.EventHandler(this.Form2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDBFiles)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion


        public System.Windows.Forms.DataGridView dataGridViewDBFiles;
        public System.Windows.Forms.DataGridViewTextBoxColumn C1;
        public System.Windows.Forms.DataGridViewTextBoxColumn C2;
        public System.Windows.Forms.DataGridViewTextBoxColumn C3;
    }
}