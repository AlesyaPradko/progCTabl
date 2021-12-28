using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SmetaAndGraphs
{
    public interface IForm1
    {
        string SmetaAdres { get; }
        string SmetaAktKS { get; }
        string WhereSave { get; }
        string OneSmetaAdres { get; }
        string WhereSaveGraph { get; }
        int SizeFront { get; }
        int AmountPeople { get; }
        int AmountDays { get; }
        int MinAmountPeople { get; set; }
        int MinAmountDays { get; set; }
        int MaxAmountPeople { get; set; }
        int MaxAmountDays { get; set; }
        string ColorGraph { get; }
        bool Flag { get; set; }
        int Days { get; }
        int Month { get; }
        int Year { get; }
        void StartFormGraf();
        void GetAllError(string error);
        void GetAllErrorGr(string error);
        event EventHandler FileStartClik;
        event EventHandler SelectModE;
        event EventHandler SelectModT;
        event EventHandler SelectDays;
        event EventHandler SelectPeople;
        event EventHandler ChangeFront;
        event EventHandler GraphStartClik;
        event EventHandler GraphFirstStartClik;
        event EventHandler ChangePeople;
        event EventHandler ChangeDays;
        event EventHandler ChangeData;
        event EventHandler SelectColor;
        event EventHandler ExitError;
    }
    public partial class Form1 : Form, IForm1
    {
        public string _smetaAdres;
        public string _oneSmetaAdres;
        public string _aktKSAdres;
        public string _saveAdres;
        public string _saveGraphAdres;
        public string _selectColor;
        public int _minDays;
        public int _maxDays;
        public int _minPeople;
        public int _maxPeople;
        public int _day;
        public int _month;
        public int _year;
        bool _flag;
        public Form1()
        {
            InitializeComponent();
        }
        #region probros_events
        private void changeColor_SelectedItemChanged(object sender, EventArgs e)
        {
            if (SelectColor != null) SelectColor(this, EventArgs.Empty);
        }

        private void radioButSelect_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton rb = sender as RadioButton;

            if (rb != null)
            {
                string colorName = rb.Name;
                switch (colorName)
                {
                    case "radioButExpert":
                        if (SelectModE != null) SelectModE(this, EventArgs.Empty);
                        break;
                    case "radioButTehnadzor":
                        if (SelectModT != null) SelectModT(this, EventArgs.Empty);
                        break;
                }
            }
        }
        private void numericFront_ValueChanged(object sender, EventArgs e)
        {
            if (ChangeFront != null) ChangeFront(this, EventArgs.Empty);
        }
        private void radioButAmountPeople_Click(object sender, EventArgs e)
        {
            if (SelectPeople != null) SelectPeople(this, EventArgs.Empty);
            numericPeople.Minimum = MinAmountPeople;
            numericPeople.Maximum = MaxAmountPeople;
            numericPeople.Visible = true;
            labelAmountPeople.Visible = true;
            numericDays.Visible = false;
            labeAmountDays.Visible = false;
            dateTimePicker1.Visible = true;
            labelSelectData.Visible = true;
            changeColor.Visible = true;
            labelColor.Visible = true;
        }

        private void radioButAmountDays_Click(object sender, EventArgs e)
        {
            if (SelectDays != null) SelectDays(this, EventArgs.Empty);
            numericDays.Minimum = MinAmountDays;
            numericDays.Maximum = MaxAmountDays;
            numericDays.Visible = true;
            labeAmountDays.Visible = true;
            numericPeople.Visible = false;
            labelAmountPeople.Visible = false;
            dateTimePicker1.Visible = true;
            labelSelectData.Visible = true;
            changeColor.Visible = true;
            labelColor.Visible = true;
        }

        private void buttonStartGraph_Click(object sender, EventArgs e)
        {

            if (GraphStartClik != null) GraphStartClik(this, EventArgs.Empty);
        }

        private void numericPeople_ValueChanged(object sender, EventArgs e)
        {

            if (ChangePeople != null) ChangePeople(this, EventArgs.Empty);

        }
        private void numericDays_ValueChanged(object sender, EventArgs e)
        {
            if (ChangeDays != null) ChangeDays(this, EventArgs.Empty);
        }
        private void dateTimePicker1_CloseUp(object sender, EventArgs e)
        {
            if (ChangeData != null) ChangeData(this, EventArgs.Empty);

        }

        private void butStartGraph_Click(object sender, EventArgs e)
        {
            if (Flag == true)
            {
                if (_oneSmetaAdres != null && _saveGraphAdres != null)
                {
                    if (GraphFirstStartClik != null) GraphFirstStartClik(this, EventArgs.Empty);
                    groupBox2.Visible = true;
                    buttonStartGraph.Visible = true;
                }
                else
                {
                    MessageBox.Show("Вы пытаетесь выполнить работу, но не выбрали cмету, и папку для сохранения!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else MessageBox.Show("Вы уже запустили процесс обработки. Подождите!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
        private void butStartWorkSmet_Click(object sender, EventArgs e)
        {
            if (_smetaAdres != null && _aktKSAdres != null && _saveAdres != null)
            { 
                if (FileStartClik != null) FileStartClik(this, EventArgs.Empty);
            }
            else MessageBox.Show("Вы пытаетесь выполнить работу, но не выбрали папку со сметами, актами КС и для сохранения!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

        }
        private void buttonExitGraph_Click(object sender, EventArgs e)
        {
            if (ExitError != null) ExitError(this, EventArgs.Empty);
            if (Flag) Application.Exit();
        }

        private void butExitWorkSmeta_Click(object sender, EventArgs e)
        {
            if (ExitError != null) ExitError(this, EventArgs.Empty);
            if (Flag) Application.Exit();
        }

        #endregion

        #region IForm1
        public string SmetaAdres { get { return _smetaAdres; } }
        public string SmetaAktKS { get { return _aktKSAdres; } }
        public string WhereSave { get { return _saveAdres; } }
        public string OneSmetaAdres { get { return _oneSmetaAdres; } }
        public string WhereSaveGraph { get { return _saveGraphAdres; } }
        public int AmountPeople { get { return (int)numericPeople.Value; } }
        public int AmountDays { get { return (int)numericDays.Value; } }
        public int SizeFront { get { return (int)numericFront.Value; } }
        public int Days { get { return dateTimePicker1.Value.Day; } }
        public int Month { get { return dateTimePicker1.Value.Month; } }
        public int Year { get { return dateTimePicker1.Value.Year; } }
        public int MinAmountPeople { get { return _minPeople; } set { _minPeople = value; } }
        public int MinAmountDays { get { return _minDays; } set { _minDays = value; } }
        public int MaxAmountPeople { get { return _maxPeople; } set { _maxPeople = value; } }
        public int MaxAmountDays { get { return _maxDays; } set { _maxDays = value; } }
        public string ColorGraph { get { return changeColor.Text; } }
        public bool Flag { get { return _flag; } set { _flag = value; } }
        public event EventHandler FileStartClik;
        public event EventHandler SelectModE;
        public event EventHandler SelectModT;
        public event EventHandler ChangeFront;
        public event EventHandler SelectDays;
        public event EventHandler SelectPeople;
        public event EventHandler GraphStartClik;
        public event EventHandler ChangePeople;
        public event EventHandler ChangeDays;
        public event EventHandler ChangeData;
        public event EventHandler GraphFirstStartClik;
        public event EventHandler SelectColor;
        public event EventHandler ExitError;
        public void StartFormGraf()
        {
            numericDays.Visible = false;
            labeAmountDays.Visible = false;
            numericPeople.Visible = false;
            labelAmountPeople.Visible = false;
            dateTimePicker1.Visible = false;
            labelSelectData.Visible = false;
            changeColor.Visible = false;
            labelColor.Visible = false;
            groupBox2.Visible = false;
            buttonStartGraph.Visible = false;
        }
        public void GetAllError(string error)
        {
            textBoxStatus.Text = error;
        }
        public void GetAllErrorGr(string error)
        {
            textBoxStatusGr.Text = error;
        }
       
       
       
        #endregion

        private void butSelectFolderAct_Click(object sender, EventArgs e)
           {
               DialogResult result = folderBrowserDialog1.ShowDialog();
               if (result == DialogResult.OK)
               {
                   _aktKSAdres = folderBrowserDialog1.SelectedPath;
               }
           }

           private void butSelectSave_Click(object sender, EventArgs e)
           {
               DialogResult result = folderBrowserDialog1.ShowDialog();
               if (result == DialogResult.OK)
               {
                   _saveAdres = folderBrowserDialog1.SelectedPath;                
               }
           }

      

        private void butSelectFolderSmet_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                _smetaAdres = folderBrowserDialog1.SelectedPath;

            }
        }
    
        private void butSelectOneSmeta_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Files Excels(*.xlsx;*.csv)|*.xlsx;*.csv";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                _oneSmetaAdres = dlg.FileName;          
            }
        }


        private void butSaveGraph_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                _saveGraphAdres = folderBrowserDialog1.SelectedPath;

            }
        }

     
    }
}
