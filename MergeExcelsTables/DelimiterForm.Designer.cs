using System;

namespace MergeExcelsTables
{
    partial class DelimiterForm
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
            this.okBtn = new System.Windows.Forms.Button();
            this.label = new System.Windows.Forms.Label();
            this.delimiterTextBox = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // okBtn
            // 
            this.okBtn.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.okBtn.Location = new System.Drawing.Point(0, 77);
            this.okBtn.Name = "okBtn";
            this.okBtn.Size = new System.Drawing.Size(230, 23);
            this.okBtn.TabIndex = 0;
            this.okBtn.Text = "OK";
            this.okBtn.UseVisualStyleBackColor = true;
            this.okBtn.Click += new System.EventHandler(this.okBtn_Click);
            // 
            // label
            // 
            this.label.AutoSize = true;
            this.label.Font = new System.Drawing.Font("Calibri", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label.Location = new System.Drawing.Point(7, 6);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(216, 36);
            this.label.TabIndex = 2;
            this.label.Text = "Enter delimiter\n(or if tab you can leave it emptye)";
            this.label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // delimiterTextBox
            // 
            this.delimiterTextBox.AccessibleName = "delimiterTextBox";
            this.delimiterTextBox.AllowDrop = true;
            this.delimiterTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.delimiterTextBox.Location = new System.Drawing.Point(66, 50);
            this.delimiterTextBox.Name = "delimiterTextBox";
            this.delimiterTextBox.Size = new System.Drawing.Size(100, 21);
            this.delimiterTextBox.TabIndex = 1;
            this.delimiterTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // DelimiterForm
            // 
            this.BackColor = System.Drawing.Color.MediumSeaGreen;
            this.ClientSize = new System.Drawing.Size(230, 100);
            this.Controls.Add(this.delimiterTextBox);
            this.Controls.Add(this.label);
            this.Controls.Add(this.okBtn);
            this.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "DelimiterForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Delimiter";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button okBtn;
        private System.Windows.Forms.Label label;
        private System.Windows.Forms.TextBox delimiterTextBox;
    }
}