using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Runtime.InteropServices;
using System.Timers;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
public enum ChangeMod { expert = 49, tehnadzor = 50, grafic = 51 };

namespace ExcelEditor.bl
{

    public class ZapredelException : Exception
    {
        public string parName;
        public ZapredelException(string s)
        {
            parName = s;
        }
    }
    public class NullvalueException : Exception
    {
        public string parName;
        public NullvalueException(string s)
        {
            parName = s;
        }
    }
    public class DonthaveExcelException : Exception
    {
        public string parName;
        public DonthaveExcelException(string s)
        {
            parName = s;
        }
    }
    public interface IFileManager
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
    public class FileManager : IFileManager
    {
        private string _userSmeta;
        private string _userOneSmeta;
        private string _userKS;
        private string _userWhereSave;
        private Excel.Application _excelApp;
        private RangeFile _processingArea;
        private string _textError=null;
        private int _size;
        private DataInput _dataStart=new DataInput();
        private int _amountDays;
        private int _amountPeople;
        private int _minDays;
        private int _minPeople;
        private int _maxDays;
        private int _maxPeople;
        public int FrontSize { get { return _size; } set { _size = value; } }
        public string TextError { get { return _textError; } set { _textError = value; } }
        public int MinDays { get { return _minDays; } set { _minDays = value; } }
        public int MaxDays { get { return _maxDays; } set { _maxDays = value; } }
        public int MinPeople { get { return _minPeople; } set { _minPeople = value; } }
        public int MaxPeople { get { return _maxPeople; } set { _maxPeople = value; } }
        public  void ExitError()
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

        public void InitializationFormGr(string adressOneSmeta, string adressWhereSaveGr)
        {
            _userOneSmeta = adressOneSmeta;
            _userWhereSave = adressWhereSaveGr;
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
        public void GetInputValueData(int day,int month,int year)
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
                ob.Initialization(_userSmeta, _userKS, _userWhereSave);
                ob.ProccessAll(_processingArea, _excelApp,_size, ref _textError);
                _excelApp.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (DirectoryNotFoundException exc)
            {
                _textError+=exc.Message;
            }
            catch (NullvalueException exc)
            {
                _textError += exc.parName;
                _excelApp.Quit();
            }
            catch (ZapredelException exc)
            {
                _textError += exc.parName;
            }
            catch (DonthaveExcelException ex)
            {
                _textError += ex.parName;
            }
            finally { }
        }
        public void StartProcessT()
        {
            try
            {
                _excelApp = CheckIt.Instance;
                _processingArea = new RangeFile();
                _processingArea.FirstCell = "A1";
                _processingArea.LastCell = "Z1200";
                Tehnadzor ob = new Tehnadzor();
                ob.Initialization(_userSmeta, _userKS, _userWhereSave);
                ob.ProccessAll(_processingArea, _excelApp, _size, ref _textError);
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
        }
        public GraphWork StartChoice()
        {
            _excelApp = CheckIt.Instance;
            _processingArea = new RangeFile();
            _processingArea.FirstCell = "A1";
            _processingArea.LastCell = "Z1200";
            GraphWork ob = new GraphWork();
            try
            {              
                ob.InitializationGrafik(_userOneSmeta, _userWhereSave);
                ob.ProccessGrafikFirst(_processingArea, _excelApp, ref _textError);
                _minDays = ob.GetMinDays();
                _maxDays = ob.GetMaxDays();
                _minPeople = ob.GetMinPeople();
                _maxPeople = ob.GetMaxPeople();
               
            }catch (NullvalueException exc)
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
            ob.ProccessGrafik(_processingArea, _excelApp,ref _textError);
            _amountPeople = 0;
            ob.InputDays(ref _amountDays, ref _amountPeople);
            ob.RecordGraph(_excelApp, _dataStart, _amountDays, _amountPeople, color,ref _textError);
            _excelApp.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            }
            catch (NullvalueException exc)
            {
                _textError += exc.parName;
               
            }
        }
        public void StartGraphPeople(int color)
        {
            try
            {
                GraphWork ob = StartChoice();
                ob.ProccessGrafik(_processingArea, _excelApp, ref _textError);
                 _amountDays = 0;
                 ob.InputWorkers(_amountPeople, ref _amountDays);
                ob.RecordGraph(_excelApp, _dataStart, _amountDays, _amountPeople, color,ref _textError);
                _excelApp.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (NullvalueException exc)
            {
                _textError += exc.parName;

            }
        }
    }
}
