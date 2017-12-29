namespace ftpCourtTransfer_New
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
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.label1 = new System.Windows.Forms.Label();
            this.copyTo = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.fromDate = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.toDate = new System.Windows.Forms.TextBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.startAt = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // webBrowser1
            // 
            this.webBrowser1.Location = new System.Drawing.Point(-5, 12);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(899, 257);
            this.webBrowser1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(753, 299);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "העתקה לספריה";
            // 
            // copyTo
            // 
            this.copyTo.Location = new System.Drawing.Point(575, 296);
            this.copyTo.Name = "copyTo";
            this.copyTo.Size = new System.Drawing.Size(161, 20);
            this.copyTo.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(753, 329);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(117, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "מתאריך או שם ספריה";
            // 
            // fromDate
            // 
            this.fromDate.Location = new System.Drawing.Point(575, 329);
            this.fromDate.Name = "fromDate";
            this.fromDate.Size = new System.Drawing.Size(161, 20);
            this.fromDate.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(413, 331);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(57, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "עד תאריך";
            // 
            // toDate
            // 
            this.toDate.Location = new System.Drawing.Point(247, 329);
            this.toDate.Name = "toDate";
            this.toDate.Size = new System.Drawing.Size(160, 20);
            this.toDate.TabIndex = 6;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(822, 357);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(15, 14);
            this.checkBox1.TabIndex = 7;
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(753, 358);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(45, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "התחל ב";
            // 
            // startAt
            // 
            this.startAt.Location = new System.Drawing.Point(575, 358);
            this.startAt.Name = "startAt";
            this.startAt.Size = new System.Drawing.Size(161, 20);
            this.startAt.TabIndex = 9;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(484, 401);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(181, 23);
            this.button1.TabIndex = 10;
            this.button1.Text = "לחץ כאן להעתקת קבצים";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(247, 401);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(187, 23);
            this.button2.TabIndex = 11;
            this.button2.Text = "לחץ כאן להשוואה  בין ספריות";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(895, 440);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.startAt);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.toDate);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.fromDate);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.copyTo);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.webBrowser1);
            this.Name = "Form1";
            this.Text = "הורדת קבצים מאתר בתי המשפט";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox copyTo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox fromDate;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox toDate;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox startAt;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}

