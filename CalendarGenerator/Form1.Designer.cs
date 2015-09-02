namespace CalendarGenerator
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.txtMonth = new System.Windows.Forms.TextBox();
            this.txtYear = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtAscTerms = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtHolidayTerms = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnHolidayColour = new System.Windows.Forms.Button();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.btnAscColour = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Arial", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 21);
            this.label1.TabIndex = 2;
            this.label1.Text = "Month";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Arial", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 49);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(85, 21);
            this.label2.TabIndex = 3;
            this.label2.Text = "Year";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // button1
            // 
            this.button1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.button1.Location = new System.Drawing.Point(0, 183);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(409, 41);
            this.button1.TabIndex = 4;
            this.button1.Text = "Generate";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtMonth
            // 
            this.txtMonth.Font = new System.Drawing.Font("Arial", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMonth.Location = new System.Drawing.Point(103, 10);
            this.txtMonth.Name = "txtMonth";
            this.txtMonth.Size = new System.Drawing.Size(102, 27);
            this.txtMonth.TabIndex = 5;
            // 
            // txtYear
            // 
            this.txtYear.Font = new System.Drawing.Font("Arial", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtYear.Location = new System.Drawing.Point(103, 43);
            this.txtYear.Name = "txtYear";
            this.txtYear.Size = new System.Drawing.Size(102, 27);
            this.txtYear.TabIndex = 6;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnAscColour);
            this.panel1.Controls.Add(this.btnHolidayColour);
            this.panel1.Controls.Add(this.txtAscTerms);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.txtHolidayTerms);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.txtYear);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.txtMonth);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(409, 183);
            this.panel1.TabIndex = 7;
            // 
            // txtAscTerms
            // 
            this.txtAscTerms.Font = new System.Drawing.Font("Arial", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtAscTerms.Location = new System.Drawing.Point(103, 132);
            this.txtAscTerms.Name = "txtAscTerms";
            this.txtAscTerms.Size = new System.Drawing.Size(294, 27);
            this.txtAscTerms.TabIndex = 10;
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Arial", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(12, 138);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(85, 21);
            this.label4.TabIndex = 9;
            this.label4.Text = "ASC Terms";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtHolidayTerms
            // 
            this.txtHolidayTerms.Font = new System.Drawing.Font("Arial", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtHolidayTerms.Location = new System.Drawing.Point(103, 98);
            this.txtHolidayTerms.Name = "txtHolidayTerms";
            this.txtHolidayTerms.Size = new System.Drawing.Size(294, 27);
            this.txtHolidayTerms.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Arial", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(12, 104);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(85, 21);
            this.label3.TabIndex = 7;
            this.label3.Text = "Holiday Terms";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnHolidayColour
            // 
            this.btnHolidayColour.Location = new System.Drawing.Point(295, 10);
            this.btnHolidayColour.Name = "btnHolidayColour";
            this.btnHolidayColour.Size = new System.Drawing.Size(102, 27);
            this.btnHolidayColour.TabIndex = 11;
            this.btnHolidayColour.Text = "Holiday Colour";
            this.btnHolidayColour.UseVisualStyleBackColor = true;
            this.btnHolidayColour.Click += new System.EventHandler(this.button2_Click);
            // 
            // btnAscColour
            // 
            this.btnAscColour.Location = new System.Drawing.Point(295, 43);
            this.btnAscColour.Name = "btnAscColour";
            this.btnAscColour.Size = new System.Drawing.Size(102, 27);
            this.btnAscColour.TabIndex = 12;
            this.btnAscColour.Text = "Asc Colour";
            this.btnAscColour.UseVisualStyleBackColor = true;
            this.btnAscColour.Click += new System.EventHandler(this.btnAscColour_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(409, 224);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtMonth;
        private System.Windows.Forms.TextBox txtYear;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox txtAscTerms;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtHolidayTerms;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnHolidayColour;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.Button btnAscColour;
    }
}

