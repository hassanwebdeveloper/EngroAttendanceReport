﻿namespace AttendanceReport
{
    partial class LMSAttendanceReportForm
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
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dtpToDate = new System.Windows.Forms.DateTimePicker();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dtpFromDate = new System.Windows.Forms.DateTimePicker();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.tbxCnic = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.cbxCompany = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.cbxCadre = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.tbxName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cbxSections = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cbxDepartments = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.dtpNdtStart = new System.Windows.Forms.DateTimePicker();
            this.dtpNdtEnd = new System.Windows.Forms.DateTimePicker();
            this.dtpNdtLunchStart = new System.Windows.Forms.DateTimePicker();
            this.label7 = new System.Windows.Forms.Label();
            this.dtpNdtLunchEnd = new System.Windows.Forms.DateTimePicker();
            this.label6 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.dtpFdtStart = new System.Windows.Forms.DateTimePicker();
            this.dtpFdtEnd = new System.Windows.Forms.DateTimePicker();
            this.dtpFdtLunchStart = new System.Windows.Forms.DateTimePicker();
            this.label8 = new System.Windows.Forms.Label();
            this.dtpFdtLunchEnd = new System.Windows.Forms.DateTimePicker();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dtpToDate);
            this.groupBox2.Location = new System.Drawing.Point(341, 150);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(211, 43);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "To Date";
            // 
            // dtpToDate
            // 
            this.dtpToDate.CustomFormat = "";
            this.dtpToDate.Location = new System.Drawing.Point(6, 17);
            this.dtpToDate.Name = "dtpToDate";
            this.dtpToDate.Size = new System.Drawing.Size(198, 20);
            this.dtpToDate.TabIndex = 1;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dtpFromDate);
            this.groupBox1.Location = new System.Drawing.Point(12, 150);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(210, 43);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "From Date";
            // 
            // dtpFromDate
            // 
            this.dtpFromDate.CustomFormat = "";
            this.dtpFromDate.Location = new System.Drawing.Point(6, 17);
            this.dtpFromDate.Name = "dtpFromDate";
            this.dtpFromDate.Size = new System.Drawing.Size(196, 20);
            this.dtpFromDate.TabIndex = 1;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.tbxCnic);
            this.groupBox4.Controls.Add(this.label11);
            this.groupBox4.Controls.Add(this.cbxCompany);
            this.groupBox4.Controls.Add(this.label10);
            this.groupBox4.Controls.Add(this.cbxCadre);
            this.groupBox4.Controls.Add(this.label9);
            this.groupBox4.Controls.Add(this.tbxName);
            this.groupBox4.Controls.Add(this.label5);
            this.groupBox4.Controls.Add(this.cbxSections);
            this.groupBox4.Controls.Add(this.label4);
            this.groupBox4.Controls.Add(this.cbxDepartments);
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Location = new System.Drawing.Point(12, 12);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(540, 132);
            this.groupBox4.TabIndex = 6;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Filter by";
            // 
            // tbxCnic
            // 
            this.tbxCnic.Location = new System.Drawing.Point(79, 94);
            this.tbxCnic.Name = "tbxCnic";
            this.tbxCnic.Size = new System.Drawing.Size(185, 20);
            this.tbxCnic.TabIndex = 21;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(6, 97);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(72, 13);
            this.label11.TabIndex = 20;
            this.label11.Text = "CNIC Number";
            // 
            // cbxCompany
            // 
            this.cbxCompany.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxCompany.FormattingEnabled = true;
            this.cbxCompany.Location = new System.Drawing.Point(344, 93);
            this.cbxCompany.Name = "cbxCompany";
            this.cbxCompany.Size = new System.Drawing.Size(188, 21);
            this.cbxCompany.TabIndex = 19;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(276, 96);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(51, 13);
            this.label10.TabIndex = 18;
            this.label10.Text = "Company";
            // 
            // cbxCadre
            // 
            this.cbxCadre.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxCadre.FormattingEnabled = true;
            this.cbxCadre.Location = new System.Drawing.Point(344, 54);
            this.cbxCadre.Name = "cbxCadre";
            this.cbxCadre.Size = new System.Drawing.Size(188, 21);
            this.cbxCadre.TabIndex = 17;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(276, 58);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(35, 13);
            this.label9.TabIndex = 16;
            this.label9.Text = "Cadre";
            // 
            // tbxName
            // 
            this.tbxName.Location = new System.Drawing.Point(79, 55);
            this.tbxName.Name = "tbxName";
            this.tbxName.Size = new System.Drawing.Size(185, 20);
            this.tbxName.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 58);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(35, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Name";
            // 
            // cbxSections
            // 
            this.cbxSections.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxSections.FormattingEnabled = true;
            this.cbxSections.Location = new System.Drawing.Point(344, 13);
            this.cbxSections.Name = "cbxSections";
            this.cbxSections.Size = new System.Drawing.Size(188, 21);
            this.cbxSections.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(276, 16);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(43, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Section";
            // 
            // cbxDepartments
            // 
            this.cbxDepartments.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxDepartments.FormattingEnabled = true;
            this.cbxDepartments.Location = new System.Drawing.Point(79, 13);
            this.cbxDepartments.Name = "cbxDepartments";
            this.cbxDepartments.Size = new System.Drawing.Size(185, 21);
            this.cbxDepartments.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 16);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Department";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.dtpNdtStart);
            this.groupBox3.Controls.Add(this.dtpNdtEnd);
            this.groupBox3.Controls.Add(this.dtpNdtLunchStart);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Controls.Add(this.dtpNdtLunchEnd);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Location = new System.Drawing.Point(12, 199);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(264, 137);
            this.groupBox3.TabIndex = 7;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Normal Duty Timing";
            // 
            // dtpNdtStart
            // 
            this.dtpNdtStart.CustomFormat = "";
            this.dtpNdtStart.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtpNdtStart.Location = new System.Drawing.Point(144, 25);
            this.dtpNdtStart.Name = "dtpNdtStart";
            this.dtpNdtStart.Size = new System.Drawing.Size(103, 20);
            this.dtpNdtStart.TabIndex = 28;
            this.dtpNdtStart.Value = new System.DateTime(2017, 8, 24, 7, 30, 0, 0);
            // 
            // dtpNdtEnd
            // 
            this.dtpNdtEnd.CustomFormat = "";
            this.dtpNdtEnd.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtpNdtEnd.Location = new System.Drawing.Point(144, 51);
            this.dtpNdtEnd.Name = "dtpNdtEnd";
            this.dtpNdtEnd.Size = new System.Drawing.Size(103, 20);
            this.dtpNdtEnd.TabIndex = 27;
            this.dtpNdtEnd.Value = new System.DateTime(2017, 8, 24, 16, 30, 0, 0);
            // 
            // dtpNdtLunchStart
            // 
            this.dtpNdtLunchStart.CustomFormat = "";
            this.dtpNdtLunchStart.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtpNdtLunchStart.Location = new System.Drawing.Point(144, 77);
            this.dtpNdtLunchStart.Name = "dtpNdtLunchStart";
            this.dtpNdtLunchStart.Size = new System.Drawing.Size(103, 20);
            this.dtpNdtLunchStart.TabIndex = 26;
            this.dtpNdtLunchStart.Value = new System.DateTime(2017, 8, 24, 12, 0, 0, 0);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 106);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(59, 13);
            this.label7.TabIndex = 25;
            this.label7.Text = "Lunch End";
            // 
            // dtpNdtLunchEnd
            // 
            this.dtpNdtLunchEnd.CustomFormat = "";
            this.dtpNdtLunchEnd.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtpNdtLunchEnd.Location = new System.Drawing.Point(144, 103);
            this.dtpNdtLunchEnd.Name = "dtpNdtLunchEnd";
            this.dtpNdtLunchEnd.Size = new System.Drawing.Size(103, 20);
            this.dtpNdtLunchEnd.TabIndex = 1;
            this.dtpNdtLunchEnd.Value = new System.DateTime(2017, 8, 24, 13, 0, 0, 0);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 80);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(62, 13);
            this.label6.TabIndex = 24;
            this.label6.Text = "Lunch Start";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 54);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 13);
            this.label2.TabIndex = 23;
            this.label2.Text = "End Time";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 22;
            this.label1.Text = "Start Time";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.dtpFdtStart);
            this.groupBox5.Controls.Add(this.dtpFdtEnd);
            this.groupBox5.Controls.Add(this.dtpFdtLunchStart);
            this.groupBox5.Controls.Add(this.label8);
            this.groupBox5.Controls.Add(this.dtpFdtLunchEnd);
            this.groupBox5.Controls.Add(this.label12);
            this.groupBox5.Controls.Add(this.label13);
            this.groupBox5.Controls.Add(this.label14);
            this.groupBox5.Location = new System.Drawing.Point(288, 199);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(264, 137);
            this.groupBox5.TabIndex = 29;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Friday Duty Timing";
            // 
            // dtpFdtStart
            // 
            this.dtpFdtStart.CustomFormat = "";
            this.dtpFdtStart.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtpFdtStart.Location = new System.Drawing.Point(144, 25);
            this.dtpFdtStart.Name = "dtpFdtStart";
            this.dtpFdtStart.Size = new System.Drawing.Size(103, 20);
            this.dtpFdtStart.TabIndex = 28;
            this.dtpFdtStart.Value = new System.DateTime(2017, 8, 24, 7, 30, 0, 0);
            // 
            // dtpFdtEnd
            // 
            this.dtpFdtEnd.CustomFormat = "";
            this.dtpFdtEnd.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtpFdtEnd.Location = new System.Drawing.Point(144, 51);
            this.dtpFdtEnd.Name = "dtpFdtEnd";
            this.dtpFdtEnd.Size = new System.Drawing.Size(103, 20);
            this.dtpFdtEnd.TabIndex = 27;
            this.dtpFdtEnd.Value = new System.DateTime(2017, 8, 24, 17, 30, 0, 0);
            // 
            // dtpFdtLunchStart
            // 
            this.dtpFdtLunchStart.CustomFormat = "";
            this.dtpFdtLunchStart.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtpFdtLunchStart.Location = new System.Drawing.Point(144, 77);
            this.dtpFdtLunchStart.Name = "dtpFdtLunchStart";
            this.dtpFdtLunchStart.Size = new System.Drawing.Size(103, 20);
            this.dtpFdtLunchStart.TabIndex = 26;
            this.dtpFdtLunchStart.Value = new System.DateTime(2017, 8, 24, 12, 30, 0, 0);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 106);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(59, 13);
            this.label8.TabIndex = 25;
            this.label8.Text = "Lunch End";
            // 
            // dtpFdtLunchEnd
            // 
            this.dtpFdtLunchEnd.CustomFormat = "";
            this.dtpFdtLunchEnd.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtpFdtLunchEnd.Location = new System.Drawing.Point(144, 103);
            this.dtpFdtLunchEnd.Name = "dtpFdtLunchEnd";
            this.dtpFdtLunchEnd.Size = new System.Drawing.Size(103, 20);
            this.dtpFdtLunchEnd.TabIndex = 1;
            this.dtpFdtLunchEnd.Value = new System.DateTime(2017, 8, 24, 14, 30, 0, 0);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(6, 80);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(62, 13);
            this.label12.TabIndex = 24;
            this.label12.Text = "Lunch Start";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(6, 54);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(52, 13);
            this.label13.TabIndex = 23;
            this.label13.Text = "End Time";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(6, 28);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(55, 13);
            this.label14.TabIndex = 22;
            this.label14.Text = "Start Time";
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(439, 342);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(113, 48);
            this.btnGenerate.TabIndex = 30;
            this.btnGenerate.Text = "Generate Report";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "pdf";
            this.saveFileDialog1.Filter = "PDF|*.pdf|Excel|*.xlsx";
            this.saveFileDialog1.Title = "Select Path To Save Report";
            this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // LMSAttendanceReportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(564, 399);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "LMSAttendanceReportForm";
            this.Text = "LMSAttendanceReportForm";
            this.groupBox2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DateTimePicker dtpToDate;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DateTimePicker dtpFromDate;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TextBox tbxCnic;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ComboBox cbxCompany;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox cbxCadre;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox tbxName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cbxSections;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cbxDepartments;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.DateTimePicker dtpNdtStart;
        private System.Windows.Forms.DateTimePicker dtpNdtEnd;
        private System.Windows.Forms.DateTimePicker dtpNdtLunchStart;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.DateTimePicker dtpNdtLunchEnd;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.DateTimePicker dtpFdtStart;
        private System.Windows.Forms.DateTimePicker dtpFdtEnd;
        private System.Windows.Forms.DateTimePicker dtpFdtLunchStart;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.DateTimePicker dtpFdtLunchEnd;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}