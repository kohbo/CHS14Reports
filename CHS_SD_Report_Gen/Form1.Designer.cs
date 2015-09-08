using System;

namespace CHSReportGen
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
            this.log = new System.Windows.Forms.TextBox();
            this.chk_autosave = new System.Windows.Forms.CheckBox();
            this.chk_supwarning = new System.Windows.Forms.CheckBox();
            this.btn_run = new System.Windows.Forms.Button();
            this.sel_report = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.endDate = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.Weeks = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.Weeks)).BeginInit();
            this.SuspendLayout();
            // 
            // log
            // 
            this.log.BackColor = System.Drawing.SystemColors.Desktop;
            this.log.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.log.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.log.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.log.Location = new System.Drawing.Point(10, 119);
            this.log.Multiline = true;
            this.log.Name = "log";
            this.log.ReadOnly = true;
            this.log.Size = new System.Drawing.Size(350, 160);
            this.log.TabIndex = 2;
            // 
            // chk_autosave
            // 
            this.chk_autosave.AutoSize = true;
            this.chk_autosave.Enabled = false;
            this.chk_autosave.Location = new System.Drawing.Point(253, 11);
            this.chk_autosave.Name = "chk_autosave";
            this.chk_autosave.Size = new System.Drawing.Size(109, 17);
            this.chk_autosave.TabIndex = 3;
            this.chk_autosave.Text = "Auto-save Report";
            this.chk_autosave.UseVisualStyleBackColor = true;
            // 
            // chk_supwarning
            // 
            this.chk_supwarning.AutoSize = true;
            this.chk_supwarning.Checked = true;
            this.chk_supwarning.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chk_supwarning.Location = new System.Drawing.Point(253, 32);
            this.chk_supwarning.Name = "chk_supwarning";
            this.chk_supwarning.Size = new System.Drawing.Size(118, 17);
            this.chk_supwarning.TabIndex = 4;
            this.chk_supwarning.Text = "Suppress Warnings";
            this.chk_supwarning.UseVisualStyleBackColor = true;
            // 
            // btn_run
            // 
            this.btn_run.Location = new System.Drawing.Point(11, 90);
            this.btn_run.Name = "btn_run";
            this.btn_run.Size = new System.Drawing.Size(349, 23);
            this.btn_run.TabIndex = 6;
            this.btn_run.Text = "Run Report";
            this.btn_run.UseVisualStyleBackColor = true;
            this.btn_run.Click += new System.EventHandler(this.btn_run_Click);
            // 
            // sel_report
            // 
            this.sel_report.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.sel_report.FormattingEnabled = true;
            this.sel_report.Items.AddRange(new object[] {
            "SLA",
            "RCA",
            "PK"});
            this.sel_report.Location = new System.Drawing.Point(57, 8);
            this.sel_report.Name = "sel_report";
            this.sel_report.Size = new System.Drawing.Size(96, 21);
            this.sel_report.TabIndex = 0;
            this.sel_report.SelectedIndexChanged += new System.EventHandler(this.SelectedReportChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(39, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Report";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 42);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(26, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "End";
            // 
            // endDate
            // 
            this.endDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.endDate.Location = new System.Drawing.Point(54, 38);
            this.endDate.Name = "endDate";
            this.endDate.Size = new System.Drawing.Size(99, 20);
            this.endDate.TabIndex = 11;
            this.endDate.Value = DateTime.Today;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 66);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Weeks";
            // 
            // Weeks
            // 
            this.Weeks.Location = new System.Drawing.Point(54, 64);
            this.Weeks.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.Weeks.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.Weeks.Name = "Weeks";
            this.Weeks.Size = new System.Drawing.Size(99, 20);
            this.Weeks.TabIndex = 13;
            this.Weeks.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(374, 291);
            this.Controls.Add(this.Weeks);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.endDate);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.sel_report);
            this.Controls.Add(this.btn_run);
            this.Controls.Add(this.chk_supwarning);
            this.Controls.Add(this.chk_autosave);
            this.Controls.Add(this.log);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.ShowInTaskbar = false;
            this.Text = "Report Generator";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.Weeks)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox log;
        private System.Windows.Forms.CheckBox chk_autosave;
        private System.Windows.Forms.CheckBox chk_supwarning;
        private System.Windows.Forms.Button btn_run;
        private System.Windows.Forms.ComboBox sel_report;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker endDate;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown Weeks;
    }
}

