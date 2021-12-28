using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.IO;
using System.Windows;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace WpfAppSmetaGraf.Model
{
    public class NullValueException : Exception
    {
        public string parName;
        public NullValueException(string s)
        {
            parName = s;
        }
    }
    public class DontHaveExcelException : Exception
    {
        public string parName;
        public DontHaveExcelException(string s)
        {
            parName = s;
        }
    }
    public interface IModelWork
    {
        void InitializationFormTE(string adressSmeta, string adressAktKS, string adressWhereSave);
        bool SelectExpert();
        bool SelectTehnadzor();
        void StartProcessT();
        void StartProcessE();
        bool SelectDays();
        void InputDays(int amountDay);
        bool SelectPeople();
        void InputPeople(int amountPeople);
        void GetInputValueData(int day, int month, int year);
        GraphWork StartChoice();
        void StartGraphPeople(int color);
        void StartGraphDays(int color);
        void InitializationFormGr(string adressOneSmeta, string adressWhereSaveGr);
        void ExitError();
        string TextError { get; set; }
        int FrontSize { get; set; }
        int MinDays { get; set; }
        int MaxDays { get; set; }
        int MinPeople { get; set; }
        int MaxPeople { get; set; }

    }
    public class ModelWork : IModelWork
    {
        private string _userSmeta;
        private string _userOneSmeta;
        private string _userKS;
        private string _userWhereSave;
        private string _userWhereSaveGraph;
        private Excel.Application _excelApp;
        private RangeFile _processingArea;
        private string _textError = null;
        private int _size;
        private DataInput _dataStart = new DataInput();
        private int _amountDays;
        private int _amountPeople;
        private int _minDays;
        private int _minPeople;
        private int _maxDays;
        private int _maxPeople;
        public List<int> AmountWorker;
        public List<int> AmountWorkDays;
        public int FrontSize { get { return _size; } set { _size = value; } }
        public string TextError { get { return _textError; } set { _textError = value; } }
        public string AdressSmeta { get { return _userSmeta; } set { _userSmeta = value; } }
        public string AdressAktKS { get { return _userKS; } set { _userKS = value; } }
        public string AdressSaveSmeta { get { return _userWhereSave; } set { _userWhereSave = value; } }
        public string AdressOneSmeta { get { return _userOneSmeta; } set { _userOneSmeta = value; } }
        public string AdressSaveGraph { get { return _userWhereSaveGraph; } set { _userWhereSaveGraph = value; } }
        public int MinDays { get { return _minDays; } set { _minDays = value; } }
        public int MaxDays { get { return _maxDays; } set { _maxDays = value; } }
        public int MinPeople { get { return _minPeople; } set { _minPeople = value; } }
        public int MaxPeople { get { return _maxPeople; } set { _maxPeople = value; } }
        public void ExitError()
        {
            try
            {
                if (_excelApp != null)
                {
                    if (_excelApp.Workbooks.Count != 0)
                    {
                        _excelApp.Workbooks.Close();
                    }
                    _excelApp.Quit();
                }
            }
            catch (COMException ex)
            {
                _textError += ex.Message;
            }
        }

        public void InitializationFormTE(string adressSmeta, string adressAktKS, string adressWhereSave)
        {
            _userSmeta = adressSmeta;
            _userKS = adressAktKS;
            _userWhereSave = adressWhereSave;
        }

        public List<int> GetAllWorkers()
        {
            AmountWorker = new List<int>();
            int amount = _maxPeople - _minPeople + 1;
            for (int i = 0; i < amount; i++)
            {
                AmountWorker.Add(_minPeople + i);
            }
            return AmountWorker;
        }
        public List<int> GetAllWorkDays()
        {
            AmountWorkDays = new List<int>();
            int amount = _maxDays - _minDays + 1;
            for (int i = 0; i < amount; i++)
            {
                AmountWorkDays.Add(_minDays + i);
            }
            return AmountWorkDays;
        }
        public void InitializationFormGr(string adressOneSmeta, string adressWhereSaveGr)
        {
            _userOneSmeta = adressOneSmeta;
            _userWhereSaveGraph = adressWhereSaveGr;
        }
        public bool SelectExpert()
        {
            return true;
        }

        public bool SelectTehnadzor()
        {
            return true;
        }
        public bool SelectDays()
        {
            return true;
        }
        public bool SelectPeople()
        {
            return true;
        }
        public void InputDays(int amountDay)
        {
            _amountDays = amountDay;

        }
        public void InputPeople(int amountPeople)
        {
            _amountPeople = amountPeople;

        }
        public void GetInputValueData(int day, int month, int year)
        {
            _dataStart.DayStart = day;
            _dataStart.MonthStart = month;
            _dataStart.YearStart = year;
        }
        public void StartProcessE()
        {
            try
            {
                _excelApp = CheckIt.Instance;
                _processingArea = new RangeFile();
                _processingArea.FirstCell = "A1";
                _processingArea.LastCell = "AD2200";
                Expert ob = new Expert();
                ob.Initialization(AdressSmeta, AdressAktKS, AdressSaveSmeta);
                ob.ProccessAll(_processingArea, _excelApp, FrontSize, ref _textError);
                _excelApp.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (DirectoryNotFoundException exc)
            {
                _textError += exc.Message;
            }
            catch (NullValueException exc)
            {
                _textError += exc.parName;
            }
            catch (DontHaveExcelException ex)
            {
                _textError += ex.parName;
            }
            catch (COMException ex)
            {
                _textError += $"{ex.Message} Вы открыли копию сметы, над которой проводится работа программы";
            }
         
        }
        public void StartProcessT()
        {
            try
            {
                _excelApp = CheckIt.Instance;
                _processingArea = new RangeFile();
                _processingArea.FirstCell = "A1";
                _processingArea.LastCell = "AD2200";
                Tehnadzor ob = new Tehnadzor();
                ob.Initialization(AdressSmeta, AdressAktKS, AdressSaveSmeta);
                ob.ProccessAll(_processingArea, _excelApp, FrontSize, ref _textError);
                _excelApp.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (DirectoryNotFoundException exc)
            {
                _textError += exc.Message;
            }
            catch (NullValueException exc)
            {
                _textError += exc.parName;
            }
            catch (DontHaveExcelException ex)
            {
                _textError += ex.parName;
            }
            catch (COMException ex)
            {
                _textError += $"{ex.Message} Вы открыли копию сметы, над которой проводится работа программы";
            }
        }
        public GraphWork StartChoice()
        {
            _excelApp = CheckIt.Instance;
            _processingArea = new RangeFile();
            _processingArea.FirstCell = "A1";
            _processingArea.LastCell = "AD2200";
            GraphWork ob = new GraphWork();
            try
            {
                ob.InitializationGraph(AdressOneSmeta, AdressSaveGraph);
                ob.ProccessGraphFirst(_processingArea, _excelApp, ref _textError);
                _minDays = ob.GetMinDays();
                _maxDays = ob.GetMaxDays();
                _minPeople = ob.GetMinPeople();
                _maxPeople = ob.GetMaxPeople();

            }
            catch (NullValueException exc)
            {
                _textError += exc.parName;
                _excelApp.Quit();
            }
            return ob;
        }
        public void StartGraphDays(int color)
        {
            try
            {
                GraphWork ob = StartChoice();
                ob.ProccessGraph(_processingArea, _excelApp, ref _textError);
                _amountPeople = 0;
                ob.InputDays(ref _amountDays, ref _amountPeople);
                ob.RecordGraph(_excelApp, _dataStart, _amountDays, _amountPeople, color, ref _textError);
                _excelApp.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (NullValueException exc)
            {
                _textError += exc.parName;

            }
        }
        public void StartGraphPeople(int color)
        {
            try
            {
                GraphWork ob = StartChoice();
                ob.ProccessGraph(_processingArea, _excelApp, ref _textError);
                _amountDays = 0;
                ob.InputWorkers(_amountPeople, ref _amountDays);
                ob.RecordGraph(_excelApp, _dataStart, _amountDays, _amountPeople, color, ref _textError);
                _excelApp.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (NullValueException exc)
            {
                _textError += exc.parName;

            }
        }
    }
}
