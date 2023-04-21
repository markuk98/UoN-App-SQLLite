namespace UoN_App_SQLLite
{
    partial class FormLoanManage
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
            this.buttonDeviceActivity = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panelDeviceActivity = new System.Windows.Forms.Panel();
            this.DGVDeviceList = new System.Windows.Forms.DataGridView();
            this.DGVDeviceActivity = new System.Windows.Forms.DataGridView();
            this.panelSpacer1 = new System.Windows.Forms.Panel();
            this.panel1.SuspendLayout();
            this.panelDeviceActivity.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGVDeviceList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGVDeviceActivity)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonDeviceActivity
            // 
            this.buttonDeviceActivity.Location = new System.Drawing.Point(61, 65);
            this.buttonDeviceActivity.Name = "buttonDeviceActivity";
            this.buttonDeviceActivity.Size = new System.Drawing.Size(239, 60);
            this.buttonDeviceActivity.TabIndex = 0;
            this.buttonDeviceActivity.Text = "Device Activity";
            this.buttonDeviceActivity.UseVisualStyleBackColor = true;
            this.buttonDeviceActivity.Click += new System.EventHandler(this.buttonDeviceActivity_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.buttonDeviceActivity);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(425, 1213);
            this.panel1.TabIndex = 1;
            // 
            // panelDeviceActivity
            // 
            this.panelDeviceActivity.Controls.Add(this.DGVDeviceActivity);
            this.panelDeviceActivity.Controls.Add(this.panelSpacer1);
            this.panelDeviceActivity.Controls.Add(this.DGVDeviceList);
            this.panelDeviceActivity.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelDeviceActivity.Location = new System.Drawing.Point(425, 0);
            this.panelDeviceActivity.Name = "panelDeviceActivity";
            this.panelDeviceActivity.Size = new System.Drawing.Size(1369, 1213);
            this.panelDeviceActivity.TabIndex = 2;
            this.panelDeviceActivity.Visible = false;
            // 
            // DGVDeviceList
            // 
            this.DGVDeviceList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGVDeviceList.Dock = System.Windows.Forms.DockStyle.Left;
            this.DGVDeviceList.Location = new System.Drawing.Point(0, 0);
            this.DGVDeviceList.Name = "DGVDeviceList";
            this.DGVDeviceList.RowTemplate.Height = 33;
            this.DGVDeviceList.Size = new System.Drawing.Size(355, 1213);
            this.DGVDeviceList.TabIndex = 0;
            this.DGVDeviceList.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DGVDeviceList_CellClick);
            // 
            // DGVDeviceActivity
            // 
            this.DGVDeviceActivity.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGVDeviceActivity.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGVDeviceActivity.Location = new System.Drawing.Point(428, 0);
            this.DGVDeviceActivity.Name = "DGVDeviceActivity";
            this.DGVDeviceActivity.RowTemplate.Height = 33;
            this.DGVDeviceActivity.Size = new System.Drawing.Size(941, 1213);
            this.DGVDeviceActivity.TabIndex = 1;
            // 
            // panelSpacer1
            // 
            this.panelSpacer1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelSpacer1.Location = new System.Drawing.Point(355, 0);
            this.panelSpacer1.Name = "panelSpacer1";
            this.panelSpacer1.Size = new System.Drawing.Size(73, 1213);
            this.panelSpacer1.TabIndex = 2;
            // 
            // FormLoanManage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1794, 1213);
            this.Controls.Add(this.panelDeviceActivity);
            this.Controls.Add(this.panel1);
            this.Name = "FormLoanManage";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.FormLoanManage_Load);
            this.panel1.ResumeLayout(false);
            this.panelDeviceActivity.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DGVDeviceList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGVDeviceActivity)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonDeviceActivity;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panelDeviceActivity;
        private System.Windows.Forms.Panel panelSpacer1;
        private System.Windows.Forms.DataGridView DGVDeviceActivity;
        private System.Windows.Forms.DataGridView DGVDeviceList;
    }
}