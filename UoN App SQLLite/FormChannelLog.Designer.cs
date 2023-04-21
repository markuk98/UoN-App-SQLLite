namespace UoN_App_SQLLite
{
    partial class FormChannelLog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormChannelLog));
            this.buttonOK = new System.Windows.Forms.Button();
            this.webUpdateLog = new System.Windows.Forms.WebBrowser();
            this.SuspendLayout();
            // 
            // buttonOK
            // 
            this.buttonOK.BackColor = System.Drawing.Color.MidnightBlue;
            this.buttonOK.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.buttonOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonOK.ForeColor = System.Drawing.Color.White;
            this.buttonOK.Location = new System.Drawing.Point(0, 851);
            this.buttonOK.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(1735, 48);
            this.buttonOK.TabIndex = 1;
            this.buttonOK.Text = "Close";
            this.buttonOK.UseVisualStyleBackColor = false;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // webUpdateLog
            // 
            this.webUpdateLog.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webUpdateLog.Location = new System.Drawing.Point(0, 0);
            this.webUpdateLog.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.webUpdateLog.MinimumSize = new System.Drawing.Size(29, 31);
            this.webUpdateLog.Name = "webUpdateLog";
            this.webUpdateLog.Size = new System.Drawing.Size(1735, 851);
            this.webUpdateLog.TabIndex = 2;
            // 
            // FormChannelLog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1735, 899);
            this.Controls.Add(this.webUpdateLog);
            this.Controls.Add(this.buttonOK);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "FormChannelLog";
            this.Text = "Update Notes";
            this.Load += new System.EventHandler(this.ChannelLog_Load);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.WebBrowser webUpdateLog;
    }
}