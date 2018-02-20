namespace WindowsFormsApp1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.status = new System.Windows.Forms.Label();
            this.filesList = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.generujRaport = new System.Windows.Forms.Button();
            this.month = new System.Windows.Forms.ComboBox();
            this.excelProcessingProgress = new System.Windows.Forms.ProgressBar();
            this.generate_extra_rachunek = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // status
            // 
            this.status.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.status.AutoSize = true;
            this.status.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.status.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.status.Location = new System.Drawing.Point(56, 161);
            this.status.Name = "status";
            this.status.Size = new System.Drawing.Size(215, 32);
            this.status.TabIndex = 0;
            this.status.Text = "Upuść pliki tutaj";
            // 
            // filesList
            // 
            this.filesList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.filesList.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.filesList.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.filesList.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.filesList.FormattingEnabled = true;
            this.filesList.Location = new System.Drawing.Point(397, 54);
            this.filesList.Name = "filesList";
            this.filesList.Size = new System.Drawing.Size(311, 244);
            this.filesList.TabIndex = 1;
            this.filesList.SelectedIndexChanged += new System.EventHandler(this.filesList_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(392, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(112, 25);
            this.label1.TabIndex = 2;
            this.label1.Text = "Lista plików";
            // 
            // generujRaport
            // 
            this.generujRaport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.generujRaport.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.generujRaport.Location = new System.Drawing.Point(572, 324);
            this.generujRaport.Name = "generujRaport";
            this.generujRaport.Size = new System.Drawing.Size(135, 34);
            this.generujRaport.TabIndex = 3;
            this.generujRaport.Text = "Generuj raport";
            this.generujRaport.UseVisualStyleBackColor = true;
            this.generujRaport.Click += new System.EventHandler(this.button1_Click);
            // 
            // month
            // 
            this.month.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.month.FormattingEnabled = true;
            this.month.ItemHeight = 16;
            this.month.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12"});
            this.month.Location = new System.Drawing.Point(397, 324);
            this.month.Name = "month";
            this.month.Size = new System.Drawing.Size(169, 24);
            this.month.TabIndex = 5;
            this.month.Text = "Wybierz miesiąc";
            // 
            // excelProcessingProgress
            // 
            this.excelProcessingProgress.Location = new System.Drawing.Point(62, 219);
            this.excelProcessingProgress.Name = "excelProcessingProgress";
            this.excelProcessingProgress.Size = new System.Drawing.Size(209, 23);
            this.excelProcessingProgress.TabIndex = 6;
            // 
            // generate_extra_rachunek
            // 
            this.generate_extra_rachunek.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.generate_extra_rachunek.Location = new System.Drawing.Point(13, 324);
            this.generate_extra_rachunek.Name = "generate_extra_rachunek";
            this.generate_extra_rachunek.Size = new System.Drawing.Size(328, 33);
            this.generate_extra_rachunek.TabIndex = 7;
            this.generate_extra_rachunek.Text = "Wygeneruj przykładowy Rachunek Extra";
            this.generate_extra_rachunek.UseVisualStyleBackColor = true;
            this.generate_extra_rachunek.Click += new System.EventHandler(this.generate_extra_rachunek_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.ClientSize = new System.Drawing.Size(720, 370);
            this.Controls.Add(this.generate_extra_rachunek);
            this.Controls.Add(this.excelProcessingProgress);
            this.Controls.Add(this.month);
            this.Controls.Add(this.generujRaport);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.filesList);
            this.Controls.Add(this.status);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Generator Raportów";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label status;
        private System.Windows.Forms.CheckedListBox filesList;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button generujRaport;
        private System.Windows.Forms.ComboBox month;
        private System.Windows.Forms.ProgressBar excelProcessingProgress;
        private System.Windows.Forms.Button generate_extra_rachunek;
    }
}

