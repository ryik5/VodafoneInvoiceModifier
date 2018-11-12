namespace VodafoneInvoiceModifier
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.buttonOpen = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.StatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.label1 = new System.Windows.Forms.Label();
            this.labelAccount = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.labelDate = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.labelBill = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.labelSummaryNumbers = new System.Windows.Forms.Label();
            this.textBoxP1 = new System.Windows.Forms.TextBox();
            this.textBoxP2 = new System.Windows.Forms.TextBox();
            this.textBoxP3 = new System.Windows.Forms.TextBox();
            this.textBoxP4 = new System.Windows.Forms.TextBox();
            this.textBoxP5 = new System.Windows.Forms.TextBox();
            this.textBoxP6 = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBoxP8 = new System.Windows.Forms.TextBox();
            this.textBoxP7 = new System.Windows.Forms.TextBox();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.buttonAbout = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.buttonExit = new System.Windows.Forms.Button();
            this.buttonReport1 = new System.Windows.Forms.Button();
            this.buttonReport2 = new System.Windows.Forms.Button();
            this.buttonClear = new System.Windows.Forms.Button();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonOpen
            // 
            this.buttonOpen.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.Control;
            this.buttonOpen.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.buttonOpen.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOpen.Location = new System.Drawing.Point(10, 11);
            this.buttonOpen.Margin = new System.Windows.Forms.Padding(2);
            this.buttonOpen.Name = "buttonOpen";
            this.buttonOpen.Size = new System.Drawing.Size(44, 24);
            this.buttonOpen.TabIndex = 0;
            this.buttonOpen.Text = "Open";
            this.toolTip1.SetToolTip(this.buttonOpen, "Открыть счет Vodafon в текстовом формате");
            this.buttonOpen.UseVisualStyleBackColor = true;
            this.buttonOpen.Click += new System.EventHandler(this.Open_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Font = new System.Drawing.Font("Courier New", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBox1.Location = new System.Drawing.Point(10, 46);
            this.textBox1.Margin = new System.Windows.Forms.Padding(2);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox1.Size = new System.Drawing.Size(594, 273);
            this.textBox1.TabIndex = 2;
            this.textBox1.WordWrap = false;
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.StatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 321);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(1, 0, 10, 0);
            this.statusStrip1.Size = new System.Drawing.Size(801, 22);
            this.statusStrip1.TabIndex = 3;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // StatusLabel1
            // 
            this.StatusLabel1.Name = "StatusLabel1";
            this.StatusLabel1.Size = new System.Drawing.Size(73, 17);
            this.StatusLabel1.Text = "StatusLabel1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(285, 11);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Лиц. счет";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // labelAccount
            // 
            this.labelAccount.AutoSize = true;
            this.labelAccount.BackColor = System.Drawing.Color.Transparent;
            this.labelAccount.Location = new System.Drawing.Point(335, 11);
            this.labelAccount.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelAccount.Name = "labelAccount";
            this.labelAccount.Size = new System.Drawing.Size(25, 13);
            this.labelAccount.TabIndex = 6;
            this.labelAccount.Text = "л.с.";
            this.labelAccount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Location = new System.Drawing.Point(424, 30);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(48, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Период:";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // labelDate
            // 
            this.labelDate.AutoSize = true;
            this.labelDate.BackColor = System.Drawing.Color.Transparent;
            this.labelDate.Location = new System.Drawing.Point(467, 30);
            this.labelDate.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelDate.Name = "labelDate";
            this.labelDate.Size = new System.Drawing.Size(33, 13);
            this.labelDate.TabIndex = 8;
            this.labelDate.Text = "Дата";
            this.labelDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(286, 30);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(49, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "№ счета";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // labelBill
            // 
            this.labelBill.AutoSize = true;
            this.labelBill.BackColor = System.Drawing.Color.Transparent;
            this.labelBill.Location = new System.Drawing.Point(331, 30);
            this.labelBill.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelBill.Name = "labelBill";
            this.labelBill.Size = new System.Drawing.Size(29, 13);
            this.labelBill.TabIndex = 10;
            this.labelBill.Text = "счет";
            this.labelBill.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.BackColor = System.Drawing.Color.Transparent;
            this.label7.Location = new System.Drawing.Point(423, 11);
            this.label7.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(69, 13);
            this.label7.TabIndex = 11;
            this.label7.Text = "Контрактов:";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // labelSummaryNumbers
            // 
            this.labelSummaryNumbers.AutoSize = true;
            this.labelSummaryNumbers.BackColor = System.Drawing.Color.Transparent;
            this.labelSummaryNumbers.Location = new System.Drawing.Point(489, 11);
            this.labelSummaryNumbers.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelSummaryNumbers.Name = "labelSummaryNumbers";
            this.labelSummaryNumbers.Size = new System.Drawing.Size(65, 13);
            this.labelSummaryNumbers.TabIndex = 12;
            this.labelSummaryNumbers.Text = "контрактов";
            this.labelSummaryNumbers.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // textBoxP1
            // 
            this.textBoxP1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxP1.Location = new System.Drawing.Point(618, 65);
            this.textBoxP1.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxP1.Name = "textBoxP1";
            this.textBoxP1.Size = new System.Drawing.Size(168, 20);
            this.textBoxP1.TabIndex = 13;
            this.toolTip1.SetToolTip(this.textBoxP1, "p1 - Контракт");
            // 
            // textBoxP2
            // 
            this.textBoxP2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxP2.Location = new System.Drawing.Point(618, 96);
            this.textBoxP2.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxP2.Name = "textBoxP2";
            this.textBoxP2.Size = new System.Drawing.Size(168, 20);
            this.textBoxP2.TabIndex = 14;
            this.toolTip1.SetToolTip(this.textBoxP2, "p2 - Номер телефону");
            // 
            // textBoxP3
            // 
            this.textBoxP3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxP3.Location = new System.Drawing.Point(618, 127);
            this.textBoxP3.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxP3.Name = "textBoxP3";
            this.textBoxP3.Size = new System.Drawing.Size(168, 20);
            this.textBoxP3.TabIndex = 15;
            this.toolTip1.SetToolTip(this.textBoxP3, "p3 - Тарифний Пакет");
            // 
            // textBoxP4
            // 
            this.textBoxP4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxP4.Location = new System.Drawing.Point(618, 158);
            this.textBoxP4.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxP4.Name = "textBoxP4";
            this.textBoxP4.Size = new System.Drawing.Size(168, 20);
            this.textBoxP4.TabIndex = 16;
            this.toolTip1.SetToolTip(this.textBoxP4, "p4 - ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА");
            // 
            // textBoxP5
            // 
            this.textBoxP5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxP5.Location = new System.Drawing.Point(618, 189);
            this.textBoxP5.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxP5.Name = "textBoxP5";
            this.textBoxP5.Size = new System.Drawing.Size(168, 20);
            this.textBoxP5.TabIndex = 17;
            this.toolTip1.SetToolTip(this.textBoxP5, "p5 - ПОСЛУГИ МІЖНАРОДНОГО РОУМІНГУ");
            // 
            // textBoxP6
            // 
            this.textBoxP6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxP6.Location = new System.Drawing.Point(618, 220);
            this.textBoxP6.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxP6.Name = "textBoxP6";
            this.textBoxP6.Size = new System.Drawing.Size(168, 20);
            this.textBoxP6.TabIndex = 18;
            this.toolTip1.SetToolTip(this.textBoxP6, "p6 - ЗНИЖКИ");
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Location = new System.Drawing.Point(608, 46);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(184, 273);
            this.groupBox1.TabIndex = 20;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Парсеры счета Vodafone";
            this.toolTip1.SetToolTip(this.groupBox1, "Теги для парсинга счета Vodafone");
            // 
            // textBoxP8
            // 
            this.textBoxP8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxP8.Location = new System.Drawing.Point(618, 282);
            this.textBoxP8.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxP8.Name = "textBoxP8";
            this.textBoxP8.Size = new System.Drawing.Size(168, 20);
            this.textBoxP8.TabIndex = 23;
            this.toolTip1.SetToolTip(this.textBoxP8, "pStop - Маркер окончания списка контрактов");
            // 
            // textBoxP7
            // 
            this.textBoxP7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxP7.Location = new System.Drawing.Point(618, 251);
            this.textBoxP7.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxP7.Name = "textBoxP7";
            this.textBoxP7.Size = new System.Drawing.Size(168, 20);
            this.textBoxP7.TabIndex = 22;
            this.toolTip1.SetToolTip(this.textBoxP7, "p7 - ЗАГАЛОМ ЗА КОНТРАКТОМ (БЕЗ ПДВ ТА ПФ)");
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.notifyIcon1.BalloonTipText = "Парсер счета Vodafone ©RYIK 2016-2018";
            this.notifyIcon1.BalloonTipTitle = "Developed by © Yuri Ryabchenko";
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "Парсер счета Vodafone и их экспорт в Excel ©RYIK 2016-2018";
            this.notifyIcon1.Visible = true;
            // 
            // buttonAbout
            // 
            this.buttonAbout.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonAbout.AutoSize = true;
            this.buttonAbout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonAbout.BackColor = System.Drawing.SystemColors.Control;
            this.buttonAbout.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.Control;
            this.buttonAbout.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.buttonAbout.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonAbout.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonAbout.Location = new System.Drawing.Point(705, 11);
            this.buttonAbout.Margin = new System.Windows.Forms.Padding(2);
            this.buttonAbout.Name = "buttonAbout";
            this.buttonAbout.Size = new System.Drawing.Size(87, 25);
            this.buttonAbout.TabIndex = 21;
            this.buttonAbout.Text = "О программе";
            this.toolTip1.SetToolTip(this.buttonAbout, "Описание назначения программы");
            this.buttonAbout.UseVisualStyleBackColor = true;
            this.buttonAbout.Click += new System.EventHandler(this.AboutSoft);
            // 
            // buttonExit
            // 
            this.buttonExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonExit.Location = new System.Drawing.Point(608, 11);
            this.buttonExit.Margin = new System.Windows.Forms.Padding(2);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Size = new System.Drawing.Size(75, 25);
            this.buttonExit.TabIndex = 22;
            this.buttonExit.Text = "Выход";
            this.toolTip1.SetToolTip(this.buttonExit, "Выйти из программы и сохранить парсеры текстового счета");
            this.buttonExit.UseVisualStyleBackColor = true;
            this.buttonExit.Click += new System.EventHandler(this.ApplicationExit);
            // 
            // buttonReport1
            // 
            this.buttonReport1.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.Control;
            this.buttonReport1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.buttonReport1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonReport1.Location = new System.Drawing.Point(57, 11);
            this.buttonReport1.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
            this.buttonReport1.Name = "buttonReport1";
            this.buttonReport1.Size = new System.Drawing.Size(90, 24);
            this.buttonReport1.TabIndex = 23;
            this.buttonReport1.Text = "Отчет для бух.";
            this.toolTip1.SetToolTip(this.buttonReport1, "Сгенерировать отчет в EXCEL для бухгалтерии");
            this.buttonReport1.UseVisualStyleBackColor = true;
            this.buttonReport1.Click += new System.EventHandler(this.buttonReport1_Click);
            // 
            // buttonReport2
            // 
            this.buttonReport2.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.Control;
            this.buttonReport2.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.buttonReport2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonReport2.Location = new System.Drawing.Point(149, 11);
            this.buttonReport2.Margin = new System.Windows.Forms.Padding(2);
            this.buttonReport2.Name = "buttonReport2";
            this.buttonReport2.Size = new System.Drawing.Size(87, 24);
            this.buttonReport2.TabIndex = 24;
            this.buttonReport2.Text = "Отчет полный";
            this.toolTip1.SetToolTip(this.buttonReport2, "Сгенерировать полный отчет в EXCEL");
            this.buttonReport2.UseVisualStyleBackColor = true;
            this.buttonReport2.Click += new System.EventHandler(this.buttonReport2_Click);
            // 
            // buttonClear
            // 
            this.buttonClear.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.Control;
            this.buttonClear.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.buttonClear.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonClear.Location = new System.Drawing.Point(238, 11);
            this.buttonClear.Margin = new System.Windows.Forms.Padding(2);
            this.buttonClear.Name = "buttonClear";
            this.buttonClear.Size = new System.Drawing.Size(44, 24);
            this.buttonClear.TabIndex = 25;
            this.buttonClear.Text = "Clear";
            this.toolTip1.SetToolTip(this.buttonClear, "Убрать весь текст из окна просмотра");
            this.buttonClear.UseVisualStyleBackColor = true;
            this.buttonClear.Click += new System.EventHandler(this.buttonClear_Click);
            // 
            // Form1
            // 
            this.AccessibleDescription = "Парсер счетов МТС в текстовом формате и экспорт результата в Excel";
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(801, 343);
            this.Controls.Add(this.buttonClear);
            this.Controls.Add(this.buttonReport2);
            this.Controls.Add(this.buttonReport1);
            this.Controls.Add(this.buttonExit);
            this.Controls.Add(this.buttonAbout);
            this.Controls.Add(this.labelSummaryNumbers);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.labelBill);
            this.Controls.Add(this.labelDate);
            this.Controls.Add(this.labelAccount);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.buttonOpen);
            this.Controls.Add(this.textBoxP3);
            this.Controls.Add(this.textBoxP1);
            this.Controls.Add(this.textBoxP4);
            this.Controls.Add(this.textBoxP5);
            this.Controls.Add(this.textBoxP6);
            this.Controls.Add(this.textBoxP2);
            this.Controls.Add(this.textBoxP7);
            this.Controls.Add(this.textBoxP8);
            this.Controls.Add(this.groupBox1);
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "VodafoneInvoiceModifier ©RYIK 2016-2018";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonOpen;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel StatusLabel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelAccount;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label labelDate;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label labelBill;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label labelSummaryNumbers;
        private System.Windows.Forms.TextBox textBoxP1;
        private System.Windows.Forms.TextBox textBoxP2;
        private System.Windows.Forms.TextBox textBoxP3;
        private System.Windows.Forms.TextBox textBoxP4;
        private System.Windows.Forms.TextBox textBoxP5;
        private System.Windows.Forms.TextBox textBoxP6;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.Button buttonAbout;
        private System.Windows.Forms.TextBox textBoxP7;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.TextBox textBoxP8;
        private System.Windows.Forms.Button buttonExit;
        private System.Windows.Forms.Button buttonReport1;
        private System.Windows.Forms.Button buttonReport2;
        private System.Windows.Forms.Button buttonClear;
    }
}

