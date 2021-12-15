
namespace SmetaAndGraphs
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
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabPageSmeta = new System.Windows.Forms.TabPage();
            this.butStartChoise = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.numericFront = new System.Windows.Forms.NumericUpDown();
            this.groupBoxSmeta = new System.Windows.Forms.GroupBox();
            this.radioButExpert = new System.Windows.Forms.RadioButton();
            this.radioButTehnadzor = new System.Windows.Forms.RadioButton();
            this.butExitWorkSmeta = new System.Windows.Forms.Button();
            this.butStartWorkSmet = new System.Windows.Forms.Button();
            this.butSelectSave = new System.Windows.Forms.Button();
            this.butSelectFolderAct = new System.Windows.Forms.Button();
            this.butSelectFolderSmet = new System.Windows.Forms.Button();
            this.tabPageGraph = new System.Windows.Forms.TabPage();
            this.labelColor = new System.Windows.Forms.Label();
            this.changeColor = new System.Windows.Forms.DomainUpDown();
            this.butStartGraph = new System.Windows.Forms.Button();
            this.labeAmountDays = new System.Windows.Forms.Label();
            this.labelAmountPeople = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.numericDays = new System.Windows.Forms.NumericUpDown();
            this.numericPeople = new System.Windows.Forms.NumericUpDown();
            this.buttonExitGraph = new System.Windows.Forms.Button();
            this.buttonStartGraph = new System.Windows.Forms.Button();
            this.labelSelectData = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.radioButAmountDays = new System.Windows.Forms.RadioButton();
            this.radioButAmountPeople = new System.Windows.Forms.RadioButton();
            this.butSaveGraph = new System.Windows.Forms.Button();
            this.butSelectOneSmeta = new System.Windows.Forms.Button();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.fontDialog1 = new System.Windows.Forms.FontDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.tabControl.SuspendLayout();
            this.tabPageSmeta.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericFront)).BeginInit();
            this.groupBoxSmeta.SuspendLayout();
            this.tabPageGraph.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericDays)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericPeople)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl
            // 
            this.tabControl.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl.Controls.Add(this.tabPageSmeta);
            this.tabControl.Controls.Add(this.tabPageGraph);
            this.tabControl.Location = new System.Drawing.Point(1, 3);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(791, 449);
            this.tabControl.TabIndex = 0;
            // 
            // tabPageSmeta
            // 
            this.tabPageSmeta.BackColor = System.Drawing.Color.WhiteSmoke;
            this.tabPageSmeta.Controls.Add(this.butStartChoise);
            this.tabPageSmeta.Controls.Add(this.label2);
            this.tabPageSmeta.Controls.Add(this.numericFront);
            this.tabPageSmeta.Controls.Add(this.groupBoxSmeta);
            this.tabPageSmeta.Controls.Add(this.butExitWorkSmeta);
            this.tabPageSmeta.Controls.Add(this.butStartWorkSmet);
            this.tabPageSmeta.Controls.Add(this.butSelectSave);
            this.tabPageSmeta.Controls.Add(this.butSelectFolderAct);
            this.tabPageSmeta.Controls.Add(this.butSelectFolderSmet);
            this.tabPageSmeta.Location = new System.Drawing.Point(4, 22);
            this.tabPageSmeta.Name = "tabPageSmeta";
            this.tabPageSmeta.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageSmeta.Size = new System.Drawing.Size(783, 423);
            this.tabPageSmeta.TabIndex = 0;
            this.tabPageSmeta.Text = "Работа со сметами";
            // 
            // butStartChoise
            // 
            this.butStartChoise.Location = new System.Drawing.Point(49, 149);
            this.butStartChoise.Name = "butStartChoise";
            this.butStartChoise.Size = new System.Drawing.Size(166, 23);
            this.butStartChoise.TabIndex = 23;
            this.butStartChoise.Text = "Начать работу";
            this.butStartChoise.UseVisualStyleBackColor = true;
            this.butStartChoise.Click += new System.EventHandler(this.butStartChoise_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(29, 105);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(88, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "Размер шрифта";
            // 
            // numericFront
            // 
            this.numericFront.Location = new System.Drawing.Point(166, 105);
            this.numericFront.Maximum = new decimal(new int[] {
            12,
            0,
            0,
            0});
            this.numericFront.Minimum = new decimal(new int[] {
            8,
            0,
            0,
            0});
            this.numericFront.Name = "numericFront";
            this.numericFront.Size = new System.Drawing.Size(120, 20);
            this.numericFront.TabIndex = 9;
            this.numericFront.Value = new decimal(new int[] {
            8,
            0,
            0,
            0});
            this.numericFront.ValueChanged += new System.EventHandler(this.numericFront_ValueChanged);
            // 
            // groupBoxSmeta
            // 
            this.groupBoxSmeta.Controls.Add(this.radioButExpert);
            this.groupBoxSmeta.Controls.Add(this.radioButTehnadzor);
            this.groupBoxSmeta.Location = new System.Drawing.Point(32, 196);
            this.groupBoxSmeta.Name = "groupBoxSmeta";
            this.groupBoxSmeta.Size = new System.Drawing.Size(209, 50);
            this.groupBoxSmeta.TabIndex = 8;
            this.groupBoxSmeta.TabStop = false;
            this.groupBoxSmeta.Text = "Выберите режим работы";
            this.groupBoxSmeta.Visible = false;
            // 
            // radioButExpert
            // 
            this.radioButExpert.AutoSize = true;
            this.radioButExpert.Location = new System.Drawing.Point(17, 19);
            this.radioButExpert.Name = "radioButExpert";
            this.radioButExpert.Size = new System.Drawing.Size(66, 17);
            this.radioButExpert.TabIndex = 1;
            this.radioButExpert.TabStop = true;
            this.radioButExpert.Text = "эксперт";
            this.radioButExpert.UseVisualStyleBackColor = true;
            this.radioButExpert.CheckedChanged += new System.EventHandler(this.radioButSelect_CheckedChanged);
            // 
            // radioButTehnadzor
            // 
            this.radioButTehnadzor.AutoSize = true;
            this.radioButTehnadzor.Location = new System.Drawing.Point(106, 19);
            this.radioButTehnadzor.Name = "radioButTehnadzor";
            this.radioButTehnadzor.Size = new System.Drawing.Size(77, 17);
            this.radioButTehnadzor.TabIndex = 0;
            this.radioButTehnadzor.TabStop = true;
            this.radioButTehnadzor.Text = "технадзор";
            this.radioButTehnadzor.UseVisualStyleBackColor = true;
            this.radioButTehnadzor.CheckedChanged += new System.EventHandler(this.radioButSelect_CheckedChanged);
            // 
            // butExitWorkSmeta
            // 
            this.butExitWorkSmeta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.butExitWorkSmeta.Location = new System.Drawing.Point(678, 362);
            this.butExitWorkSmeta.Name = "butExitWorkSmeta";
            this.butExitWorkSmeta.Size = new System.Drawing.Size(75, 23);
            this.butExitWorkSmeta.TabIndex = 7;
            this.butExitWorkSmeta.Text = "Выход";
            this.butExitWorkSmeta.UseVisualStyleBackColor = true;
            this.butExitWorkSmeta.Click += new System.EventHandler(this.butExitWorkSmeta_Click);
            // 
            // butStartWorkSmet
            // 
            this.butStartWorkSmet.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.butStartWorkSmet.Location = new System.Drawing.Point(32, 362);
            this.butStartWorkSmet.Name = "butStartWorkSmet";
            this.butStartWorkSmet.Size = new System.Drawing.Size(75, 23);
            this.butStartWorkSmet.TabIndex = 6;
            this.butStartWorkSmet.Text = "Старт";
            this.butStartWorkSmet.UseVisualStyleBackColor = true;
            this.butStartWorkSmet.Visible = false;
            this.butStartWorkSmet.Click += new System.EventHandler(this.butStartWorkSmet_Click);
            // 
            // butSelectSave
            // 
            this.butSelectSave.Location = new System.Drawing.Point(383, 29);
            this.butSelectSave.Name = "butSelectSave";
            this.butSelectSave.Size = new System.Drawing.Size(160, 52);
            this.butSelectSave.TabIndex = 5;
            this.butSelectSave.Text = "Выбрать папку для сохранения ведомости";
            this.butSelectSave.UseVisualStyleBackColor = true;
            this.butSelectSave.Click += new System.EventHandler(this.butSelectSave_Click);
            // 
            // butSelectFolderAct
            // 
            this.butSelectFolderAct.Location = new System.Drawing.Point(193, 29);
            this.butSelectFolderAct.Name = "butSelectFolderAct";
            this.butSelectFolderAct.Size = new System.Drawing.Size(160, 52);
            this.butSelectFolderAct.TabIndex = 4;
            this.butSelectFolderAct.Text = "Выбрать папку с актами КС-2";
            this.butSelectFolderAct.UseVisualStyleBackColor = true;
            this.butSelectFolderAct.Click += new System.EventHandler(this.butSelectFolderAct_Click);
            // 
            // butSelectFolderSmet
            // 
            this.butSelectFolderSmet.Location = new System.Drawing.Point(10, 29);
            this.butSelectFolderSmet.Name = "butSelectFolderSmet";
            this.butSelectFolderSmet.Size = new System.Drawing.Size(160, 52);
            this.butSelectFolderSmet.TabIndex = 3;
            this.butSelectFolderSmet.Text = "Выбрать папку со сметами";
            this.butSelectFolderSmet.UseVisualStyleBackColor = true;
            this.butSelectFolderSmet.Click += new System.EventHandler(this.butSelectFolderSmet_Click);
            // 
            // tabPageGraph
            // 
            this.tabPageGraph.BackColor = System.Drawing.Color.WhiteSmoke;
            this.tabPageGraph.Controls.Add(this.labelColor);
            this.tabPageGraph.Controls.Add(this.changeColor);
            this.tabPageGraph.Controls.Add(this.butStartGraph);
            this.tabPageGraph.Controls.Add(this.labeAmountDays);
            this.tabPageGraph.Controls.Add(this.labelAmountPeople);
            this.tabPageGraph.Controls.Add(this.dateTimePicker1);
            this.tabPageGraph.Controls.Add(this.numericDays);
            this.tabPageGraph.Controls.Add(this.numericPeople);
            this.tabPageGraph.Controls.Add(this.buttonExitGraph);
            this.tabPageGraph.Controls.Add(this.buttonStartGraph);
            this.tabPageGraph.Controls.Add(this.labelSelectData);
            this.tabPageGraph.Controls.Add(this.groupBox2);
            this.tabPageGraph.Controls.Add(this.butSaveGraph);
            this.tabPageGraph.Controls.Add(this.butSelectOneSmeta);
            this.tabPageGraph.Location = new System.Drawing.Point(4, 22);
            this.tabPageGraph.Name = "tabPageGraph";
            this.tabPageGraph.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageGraph.Size = new System.Drawing.Size(783, 423);
            this.tabPageGraph.TabIndex = 1;
            this.tabPageGraph.Text = "График производства работ";
            // 
            // labelColor
            // 
            this.labelColor.AutoSize = true;
            this.labelColor.Location = new System.Drawing.Point(421, 258);
            this.labelColor.Name = "labelColor";
            this.labelColor.Size = new System.Drawing.Size(129, 13);
            this.labelColor.TabIndex = 28;
            this.labelColor.Text = "Выберите цвет графика";
            this.labelColor.Visible = false;
            // 
            // changeColor
            // 
            this.changeColor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.changeColor.Items.Add("красный");
            this.changeColor.Items.Add("зеленый");
            this.changeColor.Items.Add("синий");
            this.changeColor.Items.Add("желтый");
            this.changeColor.Items.Add("оранжевый");
            this.changeColor.Items.Add("голубой");
            this.changeColor.Items.Add("коричневый");
            this.changeColor.Items.Add("черный");
            this.changeColor.Location = new System.Drawing.Point(424, 274);
            this.changeColor.Name = "changeColor";
            this.changeColor.Size = new System.Drawing.Size(120, 20);
            this.changeColor.TabIndex = 27;
            this.changeColor.Text = "красный";
            this.changeColor.Visible = false;
            this.changeColor.SelectedItemChanged += new System.EventHandler(this.changeColor_SelectedItemChanged);
            // 
            // butStartGraph
            // 
            this.butStartGraph.Location = new System.Drawing.Point(167, 98);
            this.butStartGraph.Name = "butStartGraph";
            this.butStartGraph.Size = new System.Drawing.Size(148, 23);
            this.butStartGraph.TabIndex = 26;
            this.butStartGraph.Text = "Начать работу";
            this.butStartGraph.UseVisualStyleBackColor = true;
            this.butStartGraph.Click += new System.EventHandler(this.butStartGraph_Click);
            // 
            // labeAmountDays
            // 
            this.labeAmountDays.Location = new System.Drawing.Point(266, 220);
            this.labeAmountDays.Name = "labeAmountDays";
            this.labeAmountDays.Size = new System.Drawing.Size(150, 30);
            this.labeAmountDays.TabIndex = 25;
            this.labeAmountDays.Text = "количество дней на работу";
            this.labeAmountDays.Visible = false;
            // 
            // labelAmountPeople
            // 
            this.labelAmountPeople.AutoSize = true;
            this.labelAmountPeople.Location = new System.Drawing.Point(15, 222);
            this.labelAmountPeople.Name = "labelAmountPeople";
            this.labelAmountPeople.Size = new System.Drawing.Size(109, 13);
            this.labelAmountPeople.TabIndex = 24;
            this.labelAmountPeople.Text = "количество человек";
            this.labelAmountPeople.Visible = false;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(84, 274);
            this.dateTimePicker1.MaxDate = new System.DateTime(2025, 12, 31, 0, 0, 0, 0);
            this.dateTimePicker1.MinDate = new System.DateTime(1999, 1, 1, 0, 0, 0, 0);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker1.TabIndex = 23;
            this.dateTimePicker1.Visible = false;
            this.dateTimePicker1.CloseUp += new System.EventHandler(this.dateTimePicker1_CloseUp);
            // 
            // numericDays
            // 
            this.numericDays.Location = new System.Drawing.Point(424, 220);
            this.numericDays.Name = "numericDays";
            this.numericDays.Size = new System.Drawing.Size(120, 20);
            this.numericDays.TabIndex = 18;
            this.numericDays.Visible = false;
            this.numericDays.ValueChanged += new System.EventHandler(this.numericDays_ValueChanged);
            // 
            // numericPeople
            // 
            this.numericPeople.Location = new System.Drawing.Point(130, 220);
            this.numericPeople.Name = "numericPeople";
            this.numericPeople.Size = new System.Drawing.Size(120, 20);
            this.numericPeople.TabIndex = 17;
            this.numericPeople.Visible = false;
            this.numericPeople.ValueChanged += new System.EventHandler(this.numericPeople_ValueChanged);
            // 
            // buttonExitGraph
            // 
            this.buttonExitGraph.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonExitGraph.Location = new System.Drawing.Point(676, 366);
            this.buttonExitGraph.Name = "buttonExitGraph";
            this.buttonExitGraph.Size = new System.Drawing.Size(75, 23);
            this.buttonExitGraph.TabIndex = 16;
            this.buttonExitGraph.Text = "Выход";
            this.buttonExitGraph.UseVisualStyleBackColor = true;
            this.buttonExitGraph.Click += new System.EventHandler(this.buttonExitGraph_Click);
            // 
            // buttonStartGraph
            // 
            this.buttonStartGraph.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonStartGraph.Location = new System.Drawing.Point(18, 366);
            this.buttonStartGraph.Name = "buttonStartGraph";
            this.buttonStartGraph.Size = new System.Drawing.Size(75, 23);
            this.buttonStartGraph.TabIndex = 15;
            this.buttonStartGraph.Text = "Старт";
            this.buttonStartGraph.UseVisualStyleBackColor = true;
            this.buttonStartGraph.Visible = false;
            this.buttonStartGraph.Click += new System.EventHandler(this.buttonStartGraph_Click);
            // 
            // labelSelectData
            // 
            this.labelSelectData.AutoSize = true;
            this.labelSelectData.Location = new System.Drawing.Point(98, 258);
            this.labelSelectData.Name = "labelSelectData";
            this.labelSelectData.Size = new System.Drawing.Size(152, 13);
            this.labelSelectData.TabIndex = 11;
            this.labelSelectData.Text = "Выберите дату начала работ";
            this.labelSelectData.Visible = false;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.radioButAmountDays);
            this.groupBox2.Controls.Add(this.radioButAmountPeople);
            this.groupBox2.Location = new System.Drawing.Point(34, 139);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(564, 50);
            this.groupBox2.TabIndex = 9;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Выберите количество человек в бригаде или продолжительность работ в рабочих днях";
            this.groupBox2.Visible = false;
            // 
            // radioButAmountDays
            // 
            this.radioButAmountDays.AutoSize = true;
            this.radioButAmountDays.Location = new System.Drawing.Point(348, 27);
            this.radioButAmountDays.Name = "radioButAmountDays";
            this.radioButAmountDays.Size = new System.Drawing.Size(153, 17);
            this.radioButAmountDays.TabIndex = 1;
            this.radioButAmountDays.TabStop = true;
            this.radioButAmountDays.Text = "количество рабочих дней";
            this.radioButAmountDays.UseVisualStyleBackColor = true;
            this.radioButAmountDays.Click += new System.EventHandler(this.radioButAmountDays_Click);
            // 
            // radioButAmountPeople
            // 
            this.radioButAmountPeople.AutoSize = true;
            this.radioButAmountPeople.Location = new System.Drawing.Point(38, 27);
            this.radioButAmountPeople.Name = "radioButAmountPeople";
            this.radioButAmountPeople.Size = new System.Drawing.Size(127, 17);
            this.radioButAmountPeople.TabIndex = 0;
            this.radioButAmountPeople.TabStop = true;
            this.radioButAmountPeople.Text = "количество человек";
            this.radioButAmountPeople.UseVisualStyleBackColor = true;
            this.radioButAmountPeople.Click += new System.EventHandler(this.radioButAmountPeople_Click);
            // 
            // butSaveGraph
            // 
            this.butSaveGraph.Location = new System.Drawing.Point(241, 19);
            this.butSaveGraph.Name = "butSaveGraph";
            this.butSaveGraph.Size = new System.Drawing.Size(160, 52);
            this.butSaveGraph.TabIndex = 6;
            this.butSaveGraph.Text = "Выбрать папку для сохранения графиков";
            this.butSaveGraph.UseVisualStyleBackColor = true;
            this.butSaveGraph.Click += new System.EventHandler(this.butSaveGraph_Click);
            // 
            // butSelectOneSmeta
            // 
            this.butSelectOneSmeta.Location = new System.Drawing.Point(34, 19);
            this.butSelectOneSmeta.Name = "butSelectOneSmeta";
            this.butSelectOneSmeta.Size = new System.Drawing.Size(160, 52);
            this.butSelectOneSmeta.TabIndex = 3;
            this.butSelectOneSmeta.Text = "Выбрать одну смету";
            this.butSelectOneSmeta.UseVisualStyleBackColor = true;
            this.butSelectOneSmeta.Click += new System.EventHandler(this.butSelectOneSmeta_Click);
            // 
            // errorProvider1
            // 
            this.errorProvider1.ContainerControl = this;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(791, 451);
            this.Controls.Add(this.tabControl);
            this.MinimumSize = new System.Drawing.Size(600, 480);
            this.Name = "Form1";
            this.Text = "Form1";
            this.tabControl.ResumeLayout(false);
            this.tabPageSmeta.ResumeLayout(false);
            this.tabPageSmeta.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericFront)).EndInit();
            this.groupBoxSmeta.ResumeLayout(false);
            this.groupBoxSmeta.PerformLayout();
            this.tabPageGraph.ResumeLayout(false);
            this.tabPageGraph.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericDays)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericPeople)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabPageSmeta;
        private System.Windows.Forms.TabPage tabPageGraph;
        private System.Windows.Forms.GroupBox groupBoxSmeta;
        private System.Windows.Forms.RadioButton radioButExpert;
        private System.Windows.Forms.RadioButton radioButTehnadzor;
        private System.Windows.Forms.Button butExitWorkSmeta;
        private System.Windows.Forms.Button butStartWorkSmet;
        private System.Windows.Forms.Button butSelectSave;
        private System.Windows.Forms.Button butSelectFolderAct;
        private System.Windows.Forms.Button butSelectFolderSmet;
        private System.Windows.Forms.Button buttonExitGraph;
        private System.Windows.Forms.Button buttonStartGraph;
        private System.Windows.Forms.Label labelSelectData;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton radioButAmountDays;
        private System.Windows.Forms.RadioButton radioButAmountPeople;
        private System.Windows.Forms.Button butSaveGraph;
        private System.Windows.Forms.Button butSelectOneSmeta;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown numericFront;
        private System.Windows.Forms.NumericUpDown numericDays;
        private System.Windows.Forms.NumericUpDown numericPeople;
        private System.Windows.Forms.ErrorProvider errorProvider1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.FontDialog fontDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button butStartChoise;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Button butStartGraph;
        private System.Windows.Forms.Label labeAmountDays;
        private System.Windows.Forms.Label labelAmountPeople;
        private System.Windows.Forms.DomainUpDown changeColor;
        private System.Windows.Forms.Label labelColor;
    }
}

