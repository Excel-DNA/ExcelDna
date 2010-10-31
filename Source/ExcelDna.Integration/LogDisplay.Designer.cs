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
            this.btnClear = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.btnCopy = new System.Windows.Forms.Button();
            this.logMessages = new System.Windows.Forms.ListView();
            this.logMessageText = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnSaveErrors = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(175, 3);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(80, 24);
            this.btnClear.TabIndex = 3;
            this.btnClear.Text = "C&lear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tableLayoutPanel1.ColumnCount = 4;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this.btnCopy, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.logMessages, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.btnSaveErrors, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.btnClear, 2, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(5, 5);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(622, 284);
            this.tableLayoutPanel1.TabIndex = 3;
            // 
            // btnCopy
            // 
            this.btnCopy.Location = new System.Drawing.Point(89, 3);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(80, 24);
            this.btnCopy.TabIndex = 4;
            this.btnCopy.Text = "&Copy";
            this.btnCopy.UseVisualStyleBackColor = true;
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // logMessages
            // 
            this.logMessages.BackColor = System.Drawing.SystemColors.Window;
            this.logMessages.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.logMessageText});
            this.tableLayoutPanel1.SetColumnSpan(this.logMessages, 4);
            this.logMessages.Dock = System.Windows.Forms.DockStyle.Fill;
            this.logMessages.FullRowSelect = true;
            this.logMessages.GridLines = true;
            this.logMessages.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.logMessages.Location = new System.Drawing.Point(3, 33);
            this.logMessages.Name = "logMessages";
            this.logMessages.Size = new System.Drawing.Size(616, 248);
            this.logMessages.TabIndex = 1;
            this.logMessages.UseCompatibleStateImageBehavior = false;
            this.logMessages.View = System.Windows.Forms.View.Details;
            this.logMessages.VirtualMode = true;
            this.logMessages.RetrieveVirtualItem += new System.Windows.Forms.RetrieveVirtualItemEventHandler(this.logMessages_RetrieveVirtualItem);
            // 
            // logMessageText
            // 
            this.logMessageText.Width = 585;
            // 
            // btnSaveErrors
            // 
            this.btnSaveErrors.Location = new System.Drawing.Point(3, 3);
            this.btnSaveErrors.Name = "btnSaveErrors";
            this.btnSaveErrors.Size = new System.Drawing.Size(80, 24);
            this.btnSaveErrors.TabIndex = 2;
            this.btnSaveErrors.Text = "&Save ...";
            this.btnSaveErrors.UseVisualStyleBackColor = true;
            this.btnSaveErrors.Click += new System.EventHandler(this.btnSaveErrors_Click);
            // 
            // LogDisplayForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(632, 294);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(450, 270);
            this.Name = "LogDisplayForm";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Excel-Dna Error Display";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.LogDisplayForm_FormClosing);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.LogDisplayForm_KeyDown);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.ListView logMessages;
        private System.Windows.Forms.Button btnSaveErrors;
        private System.Windows.Forms.ColumnHeader logMessageText;
        private System.Windows.Forms.Button btnCopy;

    }
}