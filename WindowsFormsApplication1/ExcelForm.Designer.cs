namespace ExcelTranslator
{
    partial class ExcelForm
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.PathBox = new System.Windows.Forms.TextBox();
            this.FolderDialog = new System.Windows.Forms.Button();
            this.translateButton = new System.Windows.Forms.Button();
            this.FolderDialog2 = new System.Windows.Forms.Button();
            this.PathBox2 = new System.Windows.Forms.TextBox();
            this.pBar = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.pLable = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // PathBox
            // 
            this.PathBox.Location = new System.Drawing.Point(12, 42);
            this.PathBox.Name = "PathBox";
            this.PathBox.Size = new System.Drawing.Size(436, 20);
            this.PathBox.TabIndex = 0;
            // 
            // FolderDialog
            // 
            this.FolderDialog.Location = new System.Drawing.Point(441, 41);
            this.FolderDialog.Name = "FolderDialog";
            this.FolderDialog.Size = new System.Drawing.Size(32, 23);
            this.FolderDialog.TabIndex = 1;
            this.FolderDialog.Text = "...";
            this.FolderDialog.UseVisualStyleBackColor = true;
            this.FolderDialog.Click += new System.EventHandler(this.button1_Click);
            // 
            // translateButton
            // 
            this.translateButton.Location = new System.Drawing.Point(406, 305);
            this.translateButton.Name = "translateButton";
            this.translateButton.Size = new System.Drawing.Size(68, 32);
            this.translateButton.TabIndex = 2;
            this.translateButton.Text = "Translate";
            this.translateButton.UseVisualStyleBackColor = true;
            this.translateButton.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // FolderDialog2
            // 
            this.FolderDialog2.Location = new System.Drawing.Point(442, 115);
            this.FolderDialog2.Name = "FolderDialog2";
            this.FolderDialog2.Size = new System.Drawing.Size(32, 23);
            this.FolderDialog2.TabIndex = 4;
            this.FolderDialog2.Text = "...";
            this.FolderDialog2.UseVisualStyleBackColor = true;
            this.FolderDialog2.Click += new System.EventHandler(this.button1_Click_2);
            // 
            // PathBox2
            // 
            this.PathBox2.Location = new System.Drawing.Point(11, 116);
            this.PathBox2.Name = "PathBox2";
            this.PathBox2.Size = new System.Drawing.Size(437, 20);
            this.PathBox2.TabIndex = 3;
            this.PathBox2.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // pBar
            // 
            this.pBar.Location = new System.Drawing.Point(11, 253);
            this.pBar.Name = "pBar";
            this.pBar.Size = new System.Drawing.Size(463, 17);
            this.pBar.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Excel Input Address";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 97);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(80, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Output Address";
            // 
            // pLable
            // 
            this.pLable.AutoSize = true;
            this.pLable.Location = new System.Drawing.Point(11, 277);
            this.pLable.Name = "pLable";
            this.pLable.Size = new System.Drawing.Size(16, 13);
            this.pLable.TabIndex = 8;
            this.pLable.Text = "...";
            // 
            // ExcelForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(486, 348);
            this.Controls.Add(this.pLable);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pBar);
            this.Controls.Add(this.FolderDialog2);
            this.Controls.Add(this.PathBox2);
            this.Controls.Add(this.translateButton);
            this.Controls.Add(this.FolderDialog);
            this.Controls.Add(this.PathBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "ExcelForm";
            this.ShowIcon = false;
            this.Text = "Excel Translator";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox PathBox;
        private System.Windows.Forms.Button FolderDialog;
        private System.Windows.Forms.Button translateButton;
        private System.Windows.Forms.Button FolderDialog2;
        private System.Windows.Forms.TextBox PathBox2;
        private System.Windows.Forms.ProgressBar pBar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label pLable;
    }
}

