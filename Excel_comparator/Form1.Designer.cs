namespace Excel_comparator
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.ChooseFile1 = new System.Windows.Forms.Button();
            this.textFilePath1 = new System.Windows.Forms.TextBox();
            this.ChooseFile2 = new System.Windows.Forms.Button();
            this.textFilePath2 = new System.Windows.Forms.TextBox();
            this.buttonCompare = new System.Windows.Forms.Button();
            this.labelNewPeople = new System.Windows.Forms.Label();
            this.labelMissingPeople = new System.Windows.Forms.Label();
            this.labelFile1 = new System.Windows.Forms.Label();
            this.labelFile2 = new System.Windows.Forms.Label();
            this.listNewPeople = new System.Windows.Forms.ListBox();
            this.listMissingPeople = new System.Windows.Forms.ListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // ChooseFile1
            // 
            this.ChooseFile1.Location = new System.Drawing.Point(326, 24);
            this.ChooseFile1.Name = "ChooseFile1";
            this.ChooseFile1.Size = new System.Drawing.Size(124, 20);
            this.ChooseFile1.TabIndex = 0;
            this.ChooseFile1.Text = "Выбрать  файл 1";
            this.ChooseFile1.UseVisualStyleBackColor = true;
            this.ChooseFile1.Click += new System.EventHandler(this.ChooseFile1_Click);
            // 
            // textFilePath1
            // 
            this.textFilePath1.Location = new System.Drawing.Point(19, 42);
            this.textFilePath1.Name = "textFilePath1";
            this.textFilePath1.Size = new System.Drawing.Size(314, 20);
            this.textFilePath1.TabIndex = 1;
            // 
            // ChooseFile2
            // 
            this.ChooseFile2.Location = new System.Drawing.Point(340, 117);
            this.ChooseFile2.Name = "ChooseFile2";
            this.ChooseFile2.Size = new System.Drawing.Size(123, 20);
            this.ChooseFile2.TabIndex = 2;
            this.ChooseFile2.Text = "Выбрать файл 2";
            this.ChooseFile2.UseVisualStyleBackColor = true;
            this.ChooseFile2.Click += new System.EventHandler(this.ChooseFile2_Click);
            // 
            // textFilePath2
            // 
            this.textFilePath2.Location = new System.Drawing.Point(19, 117);
            this.textFilePath2.Name = "textFilePath2";
            this.textFilePath2.Size = new System.Drawing.Size(314, 20);
            this.textFilePath2.TabIndex = 3;
            // 
            // buttonCompare
            // 
            this.buttonCompare.Location = new System.Drawing.Point(161, 160);
            this.buttonCompare.Name = "buttonCompare";
            this.buttonCompare.Size = new System.Drawing.Size(146, 23);
            this.buttonCompare.TabIndex = 4;
            this.buttonCompare.Text = "Сравнить файлы";
            this.buttonCompare.UseVisualStyleBackColor = true;
            this.buttonCompare.Click += new System.EventHandler(this.button1_Click);
            // 
            // labelNewPeople
            // 
            this.labelNewPeople.AutoSize = true;
            this.labelNewPeople.Location = new System.Drawing.Point(12, 222);
            this.labelNewPeople.Name = "labelNewPeople";
            this.labelNewPeople.Size = new System.Drawing.Size(123, 13);
            this.labelNewPeople.TabIndex = 7;
            this.labelNewPeople.Text = "Новые люди в файле 2";
            // 
            // labelMissingPeople
            // 
            this.labelMissingPeople.AutoSize = true;
            this.labelMissingPeople.Location = new System.Drawing.Point(251, 219);
            this.labelMissingPeople.Name = "labelMissingPeople";
            this.labelMissingPeople.Size = new System.Drawing.Size(169, 13);
            this.labelMissingPeople.TabIndex = 8;
            this.labelMissingPeople.Text = "Отсутствующие люди в файле 2";
            // 
            // labelFile1
            // 
            this.labelFile1.AutoSize = true;
            this.labelFile1.Location = new System.Drawing.Point(6, 0);
            this.labelFile1.Name = "labelFile1";
            this.labelFile1.Size = new System.Drawing.Size(76, 13);
            this.labelFile1.TabIndex = 9;
            this.labelFile1.Text = "Первый файл";
            // 
            // labelFile2
            // 
            this.labelFile2.AutoSize = true;
            this.labelFile2.Location = new System.Drawing.Point(6, 0);
            this.labelFile2.Name = "labelFile2";
            this.labelFile2.Size = new System.Drawing.Size(72, 13);
            this.labelFile2.TabIndex = 10;
            this.labelFile2.Text = "Второй файл";
            this.labelFile2.Click += new System.EventHandler(this.labelFile2_Click);
            // 
            // listNewPeople
            // 
            this.listNewPeople.FormattingEnabled = true;
            this.listNewPeople.Location = new System.Drawing.Point(13, 239);
            this.listNewPeople.Name = "listNewPeople";
            this.listNewPeople.Size = new System.Drawing.Size(199, 199);
            this.listNewPeople.TabIndex = 11;
            // 
            // listMissingPeople
            // 
            this.listMissingPeople.FormattingEnabled = true;
            this.listMissingPeople.Location = new System.Drawing.Point(254, 239);
            this.listMissingPeople.Name = "listMissingPeople";
            this.listMissingPeople.Size = new System.Drawing.Size(215, 199);
            this.listMissingPeople.TabIndex = 12;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.labelFile1);
            this.groupBox1.Controls.Add(this.ChooseFile1);
            this.groupBox1.Location = new System.Drawing.Point(13, 18);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(456, 61);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "groupBox1";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.labelFile2);
            this.groupBox2.Location = new System.Drawing.Point(12, 93);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(456, 61);
            this.groupBox2.TabIndex = 14;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "groupBox2";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(481, 450);
            this.Controls.Add(this.textFilePath2);
            this.Controls.Add(this.ChooseFile2);
            this.Controls.Add(this.textFilePath1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.listMissingPeople);
            this.Controls.Add(this.listNewPeople);
            this.Controls.Add(this.labelMissingPeople);
            this.Controls.Add(this.labelNewPeople);
            this.Controls.Add(this.buttonCompare);
            this.Controls.Add(this.groupBox2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Excel comparator v0.0.1337.322 pre-alpha release";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ChooseFile1;
        private System.Windows.Forms.TextBox textFilePath1;
        private System.Windows.Forms.Button ChooseFile2;
        private System.Windows.Forms.TextBox textFilePath2;
        private System.Windows.Forms.Button buttonCompare;
        private System.Windows.Forms.Label labelNewPeople;
        private System.Windows.Forms.Label labelMissingPeople;
        private System.Windows.Forms.Label labelFile1;
        private System.Windows.Forms.Label labelFile2;
        private System.Windows.Forms.ListBox listNewPeople;
        private System.Windows.Forms.ListBox listMissingPeople;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
    }
}

