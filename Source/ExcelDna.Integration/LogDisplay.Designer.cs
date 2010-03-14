namespace ExcelDna.Logging
{
    partial class LogDisplayForm
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
			this.listBoxErrors = new System.Windows.Forms.ListBox();
			this.btnSaveErrors = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// listBoxErrors
			// 
			this.listBoxErrors.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
						| System.Windows.Forms.AnchorStyles.Left)
						| System.Windows.Forms.AnchorStyles.Right)));
			this.listBoxErrors.FormattingEnabled = true;
			this.listBoxErrors.HorizontalScrollbar = true;
			this.listBoxErrors.ItemHeight = 15;
			this.listBoxErrors.Location = new System.Drawing.Point(0, 32);
			this.listBoxErrors.Name = "listBoxErrors";
			this.listBoxErrors.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
			this.listBoxErrors.Size = new System.Drawing.Size(741, 304);
			this.listBoxErrors.TabIndex = 0;
			// 
			// btnSaveErrors
			// 
			this.btnSaveErrors.Location = new System.Drawing.Point(4, 4);
			this.btnSaveErrors.Name = "btnSaveErrors";
			this.btnSaveErrors.Size = new System.Drawing.Size(80, 24);
			this.btnSaveErrors.TabIndex = 1;
			this.btnSaveErrors.Text = "Save ...";
			this.btnSaveErrors.UseVisualStyleBackColor = true;
			this.btnSaveErrors.Click += new System.EventHandler(this.btnSaveErrors_Click);
			// 
			// LogDisplayForm
			// 
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
			this.ClientSize = new System.Drawing.Size(741, 336);
			this.Controls.Add(this.btnSaveErrors);
			this.Controls.Add(this.listBoxErrors);
			this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Name = "LogDisplayForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "ExcelDna Error Display";
			this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox listBoxErrors;
		private System.Windows.Forms.Button btnSaveErrors;

    }
}