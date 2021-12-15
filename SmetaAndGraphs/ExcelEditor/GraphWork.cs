using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Data.Entity;
using Excel = Microsoft.Office.Interop.Excel;
public enum ChangeSelect { DaysforWork = 49, NumberofWorker = 50 };

namespace ExcelEditor.bl
{
    public class GraphWork
    {
        private string _adresSmeta;
        private string _userSmeta;
       // private string _userData;
        private Excel.Workbook _oneSmeta;
        private List<int> _startChapter;
        private Dictionary<Excel.Range, double> _poRazdelyTrudozatrat;
        private List<Excel.Range> _cellsAllRazdel;
        private Dictionary<int, double> _chelChasForEachWork;
        private Dictionary<int, string> _nameForEachWorkinSmeta;
        private Dictionary<string, List<int>> _dayOnEachWork;
        private List<Dictionary<int, int>> _allRazdelInOrder;
        private double _trudozatratTotal;
        public int _monthsForWork;
        public string _graphAdress;
        private Dictionary<int, int> _amounWorkInChapter;

        public GraphWork()
        { }

        //метод инициализирует листы и словари хранящие в себе сметы (адреса и книги)
        public void InitializationGrafik(string userOneSmeta/*, string userData*/, string userWhereSave)
        {
            //Console.WriteLine("InitializationGrafik");
            _userSmeta = userOneSmeta;
            //_userData = userData;        
            _graphAdress = userWhereSave;  
            
        }
        //метод для работы над папкой со сметами в режиме график
        //работа надсметой для записи графика

        public void ProccessGrafikFirst(RangeFile processingArea, Excel.Application excelApp, ref string _textError)
        {
                _oneSmeta = excelApp.Workbooks.Open(_userSmeta);
                _adresSmeta = _oneSmeta.FullName;
                Excel.Worksheet workSheetSmeta = _oneSmeta.Sheets[1];
                Excel.Range rangeSmeta = workSheetSmeta.get_Range(processingArea.FirstCell, processingArea.LastCell);
                Excel.Range cellWithTrudozatrat;
                cellWithTrudozatrat = rangeSmeta.Find("Сметная трудоемкость");
                if (cellWithTrudozatrat != null)
                {
                    string trudozatrata = cellWithTrudozatrat.Value.ToString();
                    _trudozatratTotal = ParserExc.NumeralFromCell(trudozatrata, ref _textError);
                if (_trudozatratTotal !=0)
                {
                    Marshal.FinalReleaseComObject(rangeSmeta);
                    Marshal.FinalReleaseComObject(workSheetSmeta);
                }
                else throw new NullvalueException($"Проверьте чтобы в {_adresSmeta} было верно значение Сметной трудоемкости]\n");

            }
                else
                {
                    object misValue = System.Reflection.Missing.Value;
                    Marshal.FinalReleaseComObject(rangeSmeta);
                    Marshal.FinalReleaseComObject(workSheetSmeta);
                    _oneSmeta.Close(false, misValue, misValue);
                    Marshal.FinalReleaseComObject(_oneSmeta);
                    throw new NullvalueException($"Проверьте чтобы в {_adresSmeta} было верно записано устойчивое выражение  [Сметная трудоемкость]\n");               
                }         
        }
           

        public void ProccessGrafik(RangeFile processingArea, Excel.Application excelApp, ref string _textError)
        {
            try
            {             
                Excel.Worksheet workSheetSmeta = _oneSmeta.Sheets[1];
                Excel.Range rangeSmeta = workSheetSmeta.get_Range(processingArea.FirstCell, processingArea.LastCell);
                Excel.Range keyCellNumberPozSmeta = rangeSmeta.Find("№ пп");
                Excel.Range keyCellColumnTrudozatrat = rangeSmeta.Find("Т/з осн. раб. Всего");
                if (keyCellNumberPozSmeta != null && keyCellColumnTrudozatrat != null )
                { 
                    _poRazdelyTrudozatrat = ParserExc.FindForChapter(workSheetSmeta, rangeSmeta, keyCellColumnTrudozatrat, keyCellNumberPozSmeta);
                    _cellsAllRazdel = ParserExc.FindChapter(workSheetSmeta, rangeSmeta, keyCellNumberPozSmeta);
                    _startChapter = GetFirstPosChapter(workSheetSmeta,ref _textError);
                    _amounWorkInChapter = new Dictionary<int, int>();
                    List<int> deleteChapter = new List<int>();
                    _chelChasForEachWork = ChelChasForWorks(workSheetSmeta, rangeSmeta, keyCellNumberPozSmeta, keyCellColumnTrudozatrat, _adresSmeta, ref deleteChapter,ref _textError);
                    if (_cellsAllRazdel.Count > deleteChapter.Count)
                    {
                        GetTakeNewAllRazdel(deleteChapter, ref _startChapter, ref _cellsAllRazdel);
                    }
                    _nameForEachWorkinSmeta = NameWorkInPozSmeta(workSheetSmeta, rangeSmeta, keyCellNumberPozSmeta, _adresSmeta, ref _amounWorkInChapter,ref _textError);
                    int[] startChapterWithWork = _amounWorkInChapter.Keys.ToArray();
                    if (_poRazdelyTrudozatrat.Count < _cellsAllRazdel.Count)
                    {
                        throw new NullvalueException("Проверьте написание Итого по разделу");
                    }
                    string nameSmeta;
                    int smetaId=0;
                    List<string> RazdelAll = new List<string>();
                    List<string> forRazdelAll = new List<string>();
                    using (AllContext db = new AllContext())
                    {
                        var orderDetails =
                        from details in db.Tables
                        where _adresSmeta.Contains(details.NameSmeta)
                        select details;

                        foreach (var detail in orderDetails)
                        {
                            nameSmeta = detail.NameSmeta;
                            smetaId = detail.Id;
                            break;                          
                        }
                        var orderChap =
                        from details in db.Graphs
                        where details.TableId==smetaId
                        select details;
                        foreach (var detail in orderChap)
                        {
                            RazdelAll.Add(detail.NameChapter);
                            if (detail.NameWork != null)
                            { forRazdelAll.Add(detail.NameWork); }
                            else continue;
                        }
                    }
                    if (forRazdelAll.Count != 0)
                    {
                        List<List<string>> forRazdelAllForRegex = GetListForRegex(forRazdelAll);
                        List<Regex> FORRazdelAll = new List<Regex>();

                        for (int k = 0; k < forRazdelAllForRegex.Count; k++)
                        {
                            Regex forezd = GetListRegex(forRazdelAllForRegex[k]);
                            FORRazdelAll.Add(forezd);
                        }
                        _allRazdelInOrder = new List<Dictionary<int, int>>();
                        for (int i = 0; i < RazdelAll.Count; i++)
                        {
                            RankingAllWorksInOrder(RazdelAll[i], FORRazdelAll[i], ref _allRazdelInOrder);
                        }
                    }
                    else
                    {
                        _allRazdelInOrder = new List<Dictionary<int, int>>();
                        for (int i = 0; i < RazdelAll.Count; i++)
                        {
                            RankingAllWorksInOrder(RazdelAll[i], ref _allRazdelInOrder);
                        }
                    }
                    Marshal.FinalReleaseComObject(rangeSmeta);
                    Marshal.FinalReleaseComObject(workSheetSmeta);               
                }    
                else
                {
                    _textError += $"Проверьте чтобы в {_adresSmeta} было верно записано устойчивое выражение [№ пп] или [Кол.] или [Т / з осн.раб.Всего]\n";
                    object misValue = System.Reflection.Missing.Value;
                    Marshal.FinalReleaseComObject(rangeSmeta);
                    Marshal.FinalReleaseComObject(workSheetSmeta);
                    _oneSmeta.Close(false, misValue, misValue);
                    Marshal.FinalReleaseComObject(_oneSmeta);
                    return;
                }
            }
            catch (NullReferenceException ex)
            {
                _textError += $"{ex.Message} Проверьте чтобы в {_adresSmeta} было верно записано устойчивое выражение [№ пп] или [Кол.] или [Т / з осн.раб.Всего]\n";
            }
            catch (InvalidComObjectException ex)
            {
                _textError += $"{ex.Message} Проверьте чтобы в {_adresSmeta} было верно записано устойчивое выражение [№ пп] или [Кол.] или [Т / з осн.раб.Всего]\n";
            }
            catch (COMException ex)
            {
                _textError += ex.Message;
            }
        }
        //меняет листы с первыми позициями раздела и листы с ячейками разделов, остаются только те разделы, где есть работа
        private void GetTakeNewAllRazdel(List<int> deleteChapter, ref List<int> _startChapter, ref List<Excel.Range> _cellsAllRazdel)
        {
            for (int i = _startChapter.Count - 1; i >= 0; i--)
            {
                int countchap = 0;
                for (int j = deleteChapter.Count - 1; j >= 0; j--)
                {

                    if (i == deleteChapter[j]) countchap++;
                    if (countchap == 1) break;
                }
                if (countchap == 1) continue;
                else
                {
                    _startChapter.RemoveAt(i);
                    _cellsAllRazdel.RemoveAt(i);
                }

            }
        }
        private List<List<string>> GetListForRegex(List<string> forRazdelAll)
        {
            List<List<string>> forRazdelAllForRegex = new List<List<string>>();
           
            for (int i=0;i< forRazdelAll.Count;i++)
            {
                string test = forRazdelAll[i];
                List<string> Test = new List<string>();
                string word = null;
                for (int j = 0; j < test.Length; j++)
                {
                    if (test[j] != ',')
                    {
                        word += test[j];
                    }
                    else 
                    {
                        Test.Add(word);
                        word = null;
                    }
                }
                Test.Add(word);
                forRazdelAllForRegex.Add(Test);
            }
            return forRazdelAllForRegex;
        }
        private List<int> GetFirstPosChapter(Excel.Worksheet workSheetoneSmeta, ref string _textError)
        {
            List<int> startChapter = new List<int>();
            try
            {
                for (int j = 0; j < _cellsAllRazdel.Count; j++)
                {
                    Excel.Range startChapt = workSheetoneSmeta.Cells[_cellsAllRazdel[j].Row + 1, _cellsAllRazdel[j].Column];
                    if (startChapt != null && startChapt.Value2 != null && !startChapt.MergeCells && startChapt.Value2.ToString() != "" && startChapt != null)
                    {
                        startChapter.Add(Convert.ToInt32(startChapt.Value2));
                    }
                    else
                    {
                        startChapt = workSheetoneSmeta.Cells[_cellsAllRazdel[j].Row + 2, _cellsAllRazdel[j].Column];
                        if (startChapt != null && startChapt.Value2 != null && !startChapt.MergeCells && startChapt.Value2.ToString() != "" && startChapt != null)
                        {
                            startChapter.Add(Convert.ToInt32(startChapt.Value2));
                        }
                    }
                }
            }
            catch (FormatException exc)
            {
                _textError+=$"{exc.Message} Проверьте первый столбец и первые строки после разделов\n";
            }
            return startChapter;
        }
   
        public int GetMinDays()
        {
            int inputDaysMin = (int)(0.025 * _trudozatratTotal / 8);
            if (inputDaysMin == 0) inputDaysMin += 1;
            return inputDaysMin;
        }
        public int GetMaxDays()
        {
            int inputDaysMax = (int)(0.067 * _trudozatratTotal / 8);
            return inputDaysMax;
        }
        public int GetMinPeople()
        {
            int inputWorkersMin = (int)(0.016 * _trudozatratTotal / 8);
            if (inputWorkersMin == 0) inputWorkersMin += 1;
            return inputWorkersMin;
        }
        public int GetMaxPeople()
        {
            int inputWorkersMax = (int)(0.045 * _trudozatratTotal / 8);
            return inputWorkersMax;
        }

        public void InputDays(ref int daysForWork, ref int amountWorkers)
        {
            double deltaLessThenOne;
            amountWorkers = (int)(_trudozatratTotal / (daysForWork * 8));
            deltaLessThenOne = (_trudozatratTotal / (daysForWork * 8)) - amountWorkers;
            if (deltaLessThenOne >= 0.5)
            {
                amountWorkers += 1;
                daysForWork += 2;
            }
            else
            {
                daysForWork += 3;
            }
        }
        public void InputWorkers(int amountWorkers, ref int daysForWork)
        {
            double deltaLessThenOne;
            daysForWork = (int)(_trudozatratTotal / (amountWorkers * 8));
            deltaLessThenOne = (_trudozatratTotal / (amountWorkers * 8)) - daysForWork;
            if (deltaLessThenOne > 0.05) daysForWork += 3;
            else daysForWork += 2;
        }
                   


        //возвращает строку из файла содержащего все выходные дни с 1999 по 2025 г, строку искомого месяца
        private  string FindDayMonths(DataInput dataStartWork)
        {
            string findYearString=null;
            using (AllContext db = new AllContext())
            {
                var orderDetails =
                from details in db.Days
                where details.Year== dataStartWork.YearStart.ToString()
                select details;
                foreach (var detail in orderDetails)
                {
                    findYearString = detail.DaysOfYear;
                }
            }
            int numFirstQuotes = dataStartWork.MonthStart * 2 - 1;
            int numLastQuotes = dataStartWork.MonthStart * 2;
            int countQuotes = 0;
            string freeDaysinMonthPropis = null;
            for (int i = 0; i < findYearString.Length; i++)
            {
                if (findYearString[i] == '+')
                {
                    countQuotes++;
                }
                if (countQuotes == numFirstQuotes && countQuotes < numLastQuotes && findYearString[i] != '+')
                {
                    freeDaysinMonthPropis += findYearString[i];
                }
            }
            return freeDaysinMonthPropis;
        }

        //возвращает все рабочие дни определенного месяца с цчетом того с какого дня начались работы
        private List<int> GetWorksDays(string freeDaysInMonthLetter, int dayStart, int amountDaysInMonth, int daysForWork)
        {
            List<int> freeDaysinMonthInt = new List<int>();
            int freeDayInt;
            for (int i = 0; i < freeDaysInMonthLetter.Length - 1; i++)
            {
                if (i == 0)
                {
                    if (freeDaysInMonthLetter[i] >= '0' && freeDaysInMonthLetter[i] <= '9' && (freeDaysInMonthLetter[i + 1] < '0' || freeDaysInMonthLetter[i + 1] > '9'))
                    {
                        freeDayInt = freeDaysInMonthLetter[i] - '0';
                        freeDaysinMonthInt.Add(freeDayInt);
                    }
                }
                else
                {
                    if ((freeDaysInMonthLetter[i - 1] < '0' || freeDaysInMonthLetter[i - 1] > '9') && freeDaysInMonthLetter[i] >= '0' && freeDaysInMonthLetter[i] <= '9' && (freeDaysInMonthLetter[i + 1] < '0' || freeDaysInMonthLetter[i + 1] > '9'))
                    {
                        freeDayInt = freeDaysInMonthLetter[i] - '0';
                        freeDaysinMonthInt.Add(freeDayInt);
                    }
                }
                if (freeDaysInMonthLetter[i] >= '0' && freeDaysInMonthLetter[i] <= '9' && freeDaysInMonthLetter[i + 1] >= '0' && freeDaysInMonthLetter[i + 1] <= '9')
                {
                    freeDayInt = (freeDaysInMonthLetter[i] - '0') * 10 + (freeDaysInMonthLetter[i + 1] - '0');
                    freeDaysinMonthInt.Add(freeDayInt);
                }

            }
            List<int> workDaysinMonth = new List<int>();
            for (int i = dayStart; i <= amountDaysInMonth; i++)
            {
                int countFreeDayPodrad = 0;
                for (int j = 0; j < freeDaysinMonthInt.Count; j++)
                {
                    if (i == freeDaysinMonthInt[j])
                    {
                        countFreeDayPodrad++;
                    }
                }
                if (countFreeDayPodrad == 0)
                {
                    workDaysinMonth.Add(i);
                }
                if (workDaysinMonth.Count == daysForWork) break;
            }
            return workDaysinMonth;
        }
        //возвращает количество дней в каждом месяце
        private int GetDaysInMonth(DataInput dataStartWork)
        {
            int amountDaysinMonth = 0;
            switch (dataStartWork.MonthStart)
            {
                case 1: amountDaysinMonth = 31; break;
                case 2:
                    {
                        if (dataStartWork.YearStart % 4 == 0)
                        {
                            amountDaysinMonth = 29;
                        }
                        else
                        {
                            amountDaysinMonth = 28;
                        }
                        break;
                    }
                case 3: amountDaysinMonth = 31; break;
                case 4: amountDaysinMonth = 30; break;
                case 5: amountDaysinMonth = 31; break;
                case 6: amountDaysinMonth = 30; break;
                case 7: amountDaysinMonth = 31; break;
                case 8: amountDaysinMonth = 31; break;
                case 9: amountDaysinMonth = 30; break;
                case 10: amountDaysinMonth = 31; break;
                case 11: amountDaysinMonth = 30; break;
                case 12: amountDaysinMonth = 31; break;
            }
            return amountDaysinMonth;
        }
        //меняет по ссылке словарь, состоящий из строки формата месяц, год и листа из рабочих дней
        public  Dictionary<string, List<int>> GetDaysForWork(Excel.Application excelApp, DataInput dataStartWork, int daysForWork)
        {

            List<int> workDaysinMonth;
            Dictionary<string, List<int>> dayOnEachWork = new Dictionary<string, List<int>>();
            string freeDaysInMonthLetter;
            int amountDaysInMonth;
            string monthAndYearForGrafik = null;
            for (int v = 0; v < _monthsForWork + 1; v++)
            {
                if (v == 0)
                {
                    freeDaysInMonthLetter = FindDayMonths(dataStartWork);
                    amountDaysInMonth = GetDaysInMonth(dataStartWork);
                    workDaysinMonth = GetWorksDays(freeDaysInMonthLetter, dataStartWork.DayStart, amountDaysInMonth, daysForWork);
                    monthAndYearForGrafik = $"{ParserExc.MonthLetterInt(dataStartWork.MonthStart)}.{ dataStartWork.YearStart.ToString()}";                    //Console.WriteLine(rez);
                }
                else
                {
                    if (dataStartWork.MonthStart < 12)
                    {
                        dataStartWork.MonthStart += 1;
                        amountDaysInMonth = GetDaysInMonth(dataStartWork);
                        freeDaysInMonthLetter = FindDayMonths( dataStartWork);
                    }
                    else
                    {
                        dataStartWork.MonthStart = 1;
                        dataStartWork.YearStart += 1;
                        amountDaysInMonth = GetDaysInMonth(dataStartWork);
                        freeDaysInMonthLetter = FindDayMonths( dataStartWork);
                    }
                    workDaysinMonth = GetWorksDays(freeDaysInMonthLetter, 1, amountDaysInMonth, daysForWork);
                    monthAndYearForGrafik = $"{ParserExc.MonthLetterInt(dataStartWork.MonthStart)}.{ dataStartWork.YearStart.ToString()}";
                }
                dayOnEachWork.Add(monthAndYearForGrafik, workDaysinMonth);
                daysForWork -= workDaysinMonth.Count;
                if (daysForWork <= 0) break;
            }       

            return dayOnEachWork;
        }




        private Regex GetListRegex(List<string> forChapter)
        {
            Regex forChapterReg = null;
            switch (forChapter.Count)
            {
                case 1:
                    forChapterReg = new Regex($@"{forChapter[0]}", RegexOptions.IgnoreCase);
                    break;
                case 2:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]})", RegexOptions.IgnoreCase);
                    break;
                case 3:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]})", RegexOptions.IgnoreCase);
                    break;
                case 4:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]})", RegexOptions.IgnoreCase);
                    break;
                case 5:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]})", RegexOptions.IgnoreCase);
                    break;
                case 6:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]})", RegexOptions.IgnoreCase);
                    break;
                case 7:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]})", RegexOptions.IgnoreCase);
                    break;
                case 8:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]})", RegexOptions.IgnoreCase);
                    break;
                case 9:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]})", RegexOptions.IgnoreCase);
                    break;
                case 10:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]})", RegexOptions.IgnoreCase);
                    break;
                case 11:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]})", RegexOptions.IgnoreCase);
                    break;
                case 12:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]})", RegexOptions.IgnoreCase);
                    break;
                case 13:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]})", RegexOptions.IgnoreCase);
                    break;
                case 14:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]})", RegexOptions.IgnoreCase);
                    break;
                case 15:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]})", RegexOptions.IgnoreCase);
                    break;
                case 16:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]})", RegexOptions.IgnoreCase);
                    break;
                case 17:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]})", RegexOptions.IgnoreCase);
                    break;
                case 18:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]}|{forChapter[17]})", RegexOptions.IgnoreCase);
                    break;
                case 19:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]}|{forChapter[17]}|{forChapter[18]})", RegexOptions.IgnoreCase);
                    break;
                case 20:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]}|{forChapter[17]}|{forChapter[18]}|{forChapter[19]})", RegexOptions.IgnoreCase);
                    break;
                case 21:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]}|{forChapter[17]}|{forChapter[18]}|{forChapter[19]}|{forChapter[20]})", RegexOptions.IgnoreCase);
                    break;
                case 22:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]}|{forChapter[17]}|{forChapter[18]}|{forChapter[19]}|{forChapter[20]}|{forChapter[21]})", RegexOptions.IgnoreCase);
                    break;
                case 23:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]}|{forChapter[17]}|{forChapter[18]}|{forChapter[19]}|{forChapter[20]}|{forChapter[21]}|{forChapter[22]})", RegexOptions.IgnoreCase);
                    break;
                case 24:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]}|{forChapter[17]}|{forChapter[18]}|{forChapter[19]}|{forChapter[20]}|{forChapter[21]}|{forChapter[22]}|{forChapter[23]})", RegexOptions.IgnoreCase);
                    break;
            }
            return forChapterReg;
        }


        //возвращает  словарь, где ключ - номер по смете, значение - трудозатраты на данную работу
        private Dictionary<int, double> ChelChasForWorks(Excel.Worksheet workSheetSmeta, Excel.Range rangeSmeta, Excel.Range keyCellNumberPosSmeta, Excel.Range keyCellColumnTopTrudozatrat, string AdresSmeta, ref List<int> deleteChapter,ref string _textError)
        {
            Dictionary<int, double> chelChasforEachWork = new Dictionary<int, double>();
            double trudozatratOfWork;
            int numPosSmeta; 
            for (int j = keyCellNumberPosSmeta.Row + 4; j <= rangeSmeta.Rows.Count; j++)
            {
                Excel.Range cellsNumberPosColumnTabl = workSheetSmeta.Cells[j, keyCellNumberPosSmeta.Column];
                Excel.Range cellsColumnTrudozatrat = workSheetSmeta.Cells[j, keyCellColumnTopTrudozatrat.Column];
                if (cellsNumberPosColumnTabl != null && cellsNumberPosColumnTabl.Value2 != null && !cellsNumberPosColumnTabl.MergeCells && cellsNumberPosColumnTabl.Value2.ToString() != "" && cellsColumnTrudozatrat != null && cellsColumnTrudozatrat.Value2 != null && !cellsColumnTrudozatrat.MergeCells && cellsColumnTrudozatrat.Value2.ToString() != "")
                {
                    try
                    {
                        int numCellsForNumPosSmeta = cellsNumberPosColumnTabl.Row;
                        numPosSmeta = Convert.ToInt32(cellsNumberPosColumnTabl.Value2);
                        trudozatratOfWork = Convert.ToDouble(cellsColumnTrudozatrat.Value2);
                        chelChasforEachWork.Add(numPosSmeta, trudozatratOfWork);
                        for (int i = 0; i < _startChapter.Count; i++)
                        {                          
                            if (_startChapter[i] == numPosSmeta)
                            {
                                deleteChapter.Add(i);
                            }

                        }
                    }
                    catch (NullReferenceException ex)
                    {
                        _textError+=$"{ex.Message} Проверьте чтобы в {AdresSmeta} было верно записано устойчивое выражение [Наименование]\n";
                    }
                    catch (ArgumentException ex)
                    {
                        _textError += $"{ex.Message} Проверьте чтобы в {AdresSmeta} не повторялись значения позиций по смете в строке {cellsNumberPosColumnTabl.Row}\n";
                    }
                    catch (FormatException ex)
                    {
                        _textError+=$"{ex.Message} Вы ввели неверный формат для {AdresSmeta} в строке {cellsNumberPosColumnTabl.Row} в столбце {cellsNumberPosColumnTabl.Column}(не должно быть [., букв], только целые числа,или в столбце {cellsColumnTrudozatrat.Column} только числа дробные, не должно быть [.букв]  )\n";
                    }
                }
            }
            return chelChasforEachWork;
        }
        //возвращает  словарь, где ключ - номер по смете, значение - строковое наименование данных работ
        private Dictionary<int, string> NameWorkInPozSmeta(Excel.Worksheet workSheetSmeta, Excel.Range rangeSmeta, Excel.Range keyCellNumberPosSmeta, string AdresSmeta, ref Dictionary<int, int> _amounWorkInChapter, ref string _textError)
        {
            int[] keyTrudozatratEachWork = _chelChasForEachWork.Keys.ToArray();
            Dictionary<int, string> nameForEachWorkInSmeta = new Dictionary<int, string>();
            int numPosSmeta;
            int countAmount = 0;
            string nameWorkInPosSmeta;
            Excel.Range rangeSmetaForCell1 = workSheetSmeta.Cells[keyCellNumberPosSmeta.Row, 1];
            Excel.Range rangeSmetaForCell2 = workSheetSmeta.Cells[keyCellNumberPosSmeta.Row + 3, 30];
            Excel.Range rangeSmetaForCell = workSheetSmeta.get_Range(rangeSmetaForCell1, rangeSmetaForCell2);
            Excel.Range keyCellNameWork = rangeSmetaForCell.Find("Наименование");
            if (keyCellNameWork != null)
            {
                for (int j = keyCellNumberPosSmeta.Row + 4; j <= rangeSmeta.Rows.Count; j++)
                {

                    Excel.Range cellsNumberPosColumnTabl = workSheetSmeta.Cells[j, keyCellNumberPosSmeta.Column];
                    Excel.Range cellsNameWorkColumnTabl = workSheetSmeta.Cells[j, keyCellNameWork.Column];
                    if (cellsNumberPosColumnTabl != null && cellsNumberPosColumnTabl.Value2 != null && !cellsNumberPosColumnTabl.MergeCells && cellsNumberPosColumnTabl.Value2.ToString() != "" && cellsNameWorkColumnTabl != null && cellsNameWorkColumnTabl.Value2 != null && !cellsNameWorkColumnTabl.MergeCells && cellsNameWorkColumnTabl.Value2.ToString() != "")
                    {
                        try
                        {
                            for (int i = 0; i < keyTrudozatratEachWork.Length; i++)
                            {

                                numPosSmeta = Convert.ToInt32(cellsNumberPosColumnTabl.Value2);
                                if (numPosSmeta == keyTrudozatratEachWork[i])
                                {
                                    nameWorkInPosSmeta = cellsNameWorkColumnTabl.Value.ToString();
                                    nameForEachWorkInSmeta.Add(numPosSmeta, nameWorkInPosSmeta);
                                    for (int k = 0; k < _cellsAllRazdel.Count; k++)
                                    {
                                        int fff = _startChapter[k];
                                        if (numPosSmeta == _startChapter[k])
                                        {

                                            if (i == 0)
                                            {
                                                countAmount = 0;
                                            }
                                            else
                                            {
                                                _amounWorkInChapter.Add(_startChapter[k - 1], countAmount);
                                                countAmount = 0;
                                            }
                                        }
                                    }
                                    countAmount++;
                                }
                            }
                        }

                        catch (ArgumentException ex)
                        {
                            _textError += $"{ex.Message} Проверьте чтобы в {AdresSmeta} не повторялись значения позиций по смете в строке {cellsNumberPosColumnTabl.Row}\n";
                        }
                        catch (FormatException ex)
                        {
                            _textError += $"{ex.Message} Вы ввели неверный формат для {AdresSmeta} в строке {cellsNumberPosColumnTabl.Row} в столбце {cellsNumberPosColumnTabl.Column}(не должно быть [., букв], только целые числа.\n";
                        }
                    }
                }
            }
            else throw new NullvalueException($"Проверьте чтобы в смете {AdresSmeta}  верно было написано [Наименование] в шапке");
            if (countAmount != 0) { _amounWorkInChapter.Add(_startChapter[_startChapter.Count - 1], countAmount); }
            return nameForEachWorkInSmeta;
        }

        //меняет по ссылке лист, состоящий из словарей,где ключ - номер по смете, значение - номер строки в массиве наименования данных работ для всех разделов
        private void RankingAllWorksInOrder(string regulNameOfRazdel, Regex regulNameWorkOfRazdel, ref List<Dictionary<int, int>> _allRazdelInOrder)
        {
            int[] keyNumTrudozatratEachWork = _chelChasForEachWork.Keys.ToArray();
            string[] valueNameofEachWork = _nameForEachWorkinSmeta.Values.ToArray();
            Dictionary<int, int> inChapterNumPosAndNumWorkInArr;

            for (int i = 0; i < _cellsAllRazdel.Count; i++)
            {
                string stringPoRazdelyforPoisk = _cellsAllRazdel[i].Value.ToString();
                if (stringPoRazdelyforPoisk.Contains(regulNameOfRazdel))
                {
                    if (i < _cellsAllRazdel.Count - 1)
                    {
                        inChapterNumPosAndNumWorkInArr = ParserExc.InOrderChapter(regulNameWorkOfRazdel, valueNameofEachWork, keyNumTrudozatratEachWork, _startChapter[i], _startChapter[i + 1]);
                        if (inChapterNumPosAndNumWorkInArr.Count > 0)
                        {
                            _allRazdelInOrder.Add(inChapterNumPosAndNumWorkInArr);
                        }
                    }
                    else
                    {
                        inChapterNumPosAndNumWorkInArr = ParserExc.InOrderChapter(regulNameWorkOfRazdel, valueNameofEachWork, keyNumTrudozatratEachWork, _startChapter[i]);
                        if (inChapterNumPosAndNumWorkInArr.Count > 0)
                        {
                            _allRazdelInOrder.Add(inChapterNumPosAndNumWorkInArr);
                        }
                    }
                }
            }
        }

        private void RankingAllWorksInOrder(string regulNameOfRazdel, ref List<Dictionary<int, int>> _allRazdelInOrder)
        {
            int[] keyNumTrudozatratEachWork = _chelChasForEachWork.Keys.ToArray();
            string[] valueNameofEachWork = _nameForEachWorkinSmeta.Values.ToArray();
            Dictionary<int, int> inChapterNumPosAndNumWorkInArr;

            for (int i = 0; i < _cellsAllRazdel.Count; i++)
            {
                string stringPoRazdelyforPoisk = _cellsAllRazdel[i].Value.ToString();
                int countChapter = 0;
                if (stringPoRazdelyforPoisk.Contains(regulNameOfRazdel))
                {
                    if (i < _cellsAllRazdel.Count - 1)
                    {
                        inChapterNumPosAndNumWorkInArr = ParserExc.InOrderChapter(valueNameofEachWork, keyNumTrudozatratEachWork, _startChapter[i], _startChapter[i + 1]);
                        if (inChapterNumPosAndNumWorkInArr.Count > 0)
                        {
                            _allRazdelInOrder.Add(inChapterNumPosAndNumWorkInArr);
                            countChapter++;
                        }
                    }
                    else
                    {
                        inChapterNumPosAndNumWorkInArr = ParserExc.InOrderChapter(valueNameofEachWork, keyNumTrudozatratEachWork, _startChapter[i]);
                        if (inChapterNumPosAndNumWorkInArr.Count > 0)
                        {
                            _allRazdelInOrder.Add(inChapterNumPosAndNumWorkInArr);
                            countChapter++;
                        }
                    }
                }

               // if (countChapter > 0) break;
            }
        }

        //закрашивает график в соответствие с данными
        public void RecordGraph(Excel.Application excelApp, DataInput dataStartWork, int daysForWork, int numberofWorkers, int color, ref string _textError)
        {
           try { 
                string nameFailSmeta;
            ParserExc.GetNameSmeta(_adresSmeta, out nameFailSmeta);
            _graphAdress += $"\\График производства работ - {nameFailSmeta}";
            Excel.Worksheet workSheetSmeta = _oneSmeta.Sheets[1];
            Excel.Workbook workBookGraph = excelApp.Workbooks.Add();
            Excel.Worksheet workSheetGraph = (Excel.Worksheet)workBookGraph.Worksheets.get_Item(1);
            Excel.Range FirstCellGraph = workSheetGraph.Range["B4"];
            Excel.Range GraphNext = workSheetGraph.get_Range("B4", "B5");
            GraphNext.Merge();
            GraphNext.Value = "№";
            GraphNext = workSheetGraph.get_Range("C4", "C5");
            GraphNext.Merge();
            GraphNext.Value = "Наименование работ";
            GraphNext = workSheetGraph.get_Range("D4", "D5");
            GraphNext.Merge();
            GraphNext.Value = "Всего чел/час";
            GraphNext = workSheetGraph.get_Range("E4", "E5");
            GraphNext.Merge();
            GraphNext.Value = "Кол. чел.  бр";
            GraphNext = workSheetGraph.get_Range("F4", "F5");
            GraphNext.Merge();
            GraphNext.Value = "Кол-во рабоч. дней";
            Excel.Range firstMonth, lastMonth = null;
            double delta = (daysForWork / 21.0) - (int)(daysForWork / 21);
            if (delta < 0.04) _monthsForWork = daysForWork / 21;
            else _monthsForWork = 1 + daysForWork / 21;
            _dayOnEachWork = GetDaysForWork(excelApp, dataStartWork, daysForWork);
            List<int>[] valueAllWorkDaysForMonth = _dayOnEachWork.Values.ToArray();
            string[] keyNameDataWork = _dayOnEachWork.Keys.ToArray();
            for (int i = 0; i < valueAllWorkDaysForMonth.Length; i++)
            {
                firstMonth = workSheetGraph.Cells[GraphNext.Row, GraphNext.Column + 1];
                lastMonth = workSheetGraph.Cells[GraphNext.Row, GraphNext.Column + valueAllWorkDaysForMonth[i].Count];
                for (int j = 0; j < valueAllWorkDaysForMonth[i].Count; j++)
                {
                    workSheetGraph.Cells[firstMonth.Row + 1, firstMonth.Column + j] = valueAllWorkDaysForMonth[i][j];
                }
                GraphNext = workSheetGraph.get_Range(firstMonth, lastMonth);
                GraphNext.Merge();
                GraphNext.Value = keyNameDataWork[i];
                GraphNext = lastMonth;
            }
            
            int amountOfWorkInChapter = 0;
            int[] numChapterTablExcelGraph = new int[_allRazdelInOrder.Count];
                int[] amountOfWorkersInChapter= new int[_allRazdelInOrder.Count]; ;
                RecordAllString(workSheetSmeta, numberofWorkers, workSheetGraph, ref amountOfWorkInChapter, ref numChapterTablExcelGraph,ref amountOfWorkersInChapter);
            Excel.Range LastCellGraph = workSheetGraph.Cells[FirstCellGraph.Row + amountOfWorkInChapter + 1, lastMonth.Column];
            Excel.Range rangeGraph = workSheetGraph.get_Range(FirstCellGraph, LastCellGraph);
            rangeGraph.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
            rangeGraph.EntireColumn.Font.Size = 10;
            rangeGraph.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
            rangeGraph.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
            rangeGraph.EntireColumn.AutoFit();
            Excel.Range cellforDaysSimilarSize = workSheetGraph.get_Range("G5", LastCellGraph);
            cellforDaysSimilarSize.ColumnWidth = 4;
            Excel.Range rangeForColour = workSheetGraph.get_Range("E6", LastCellGraph);
            int amountOfDaysOnEachChapter, amountofWorkerOnEachChapter;
            double deltaWorker = 0;
            double allWorker, sumAllWork = 0;
            int amountOfDaysOnAllChapter = 0, summaAmountofDaysEachWork = 0, summaAmountofWorkerEachWork = 0, indexofChapter = 0;
            for (int j = rangeForColour.Row; j < rangeForColour.Rows.Count + rangeForColour.Row; j++)
            {
                if (indexofChapter < numChapterTablExcelGraph.Length)
                {
                    if (j == numChapterTablExcelGraph[indexofChapter])
                    {
                        indexofChapter++;
                        Excel.Range amountofDaysEachRazdelTabl = workSheetGraph.Cells[numChapterTablExcelGraph[indexofChapter - 1], 6];
                        amountOfDaysOnEachChapter = (int)(amountofDaysEachRazdelTabl.Value2);
                        amountOfDaysOnAllChapter += amountOfDaysOnEachChapter;
                    }
                }
                if (indexofChapter > 0)
                {
                    if (j >= numChapterTablExcelGraph[indexofChapter - 1] + 1)
                    {
                        Excel.Range amountofWorkerEachRazdelTabl = workSheetGraph.Cells[numChapterTablExcelGraph[indexofChapter - 1], 5];
                        amountofWorkerOnEachChapter = (int)(amountofWorkerEachRazdelTabl.Value2);
                        Excel.Range trudWork = workSheetGraph.Cells[j, 4];
                        Excel.Range numberofWorkerEachWorkTabl = workSheetGraph.Cells[j, 5];
                        Excel.Range numberofDaysEachWorkTabl = workSheetGraph.Cells[j, 6];
                        allWorker = (int)(numberofWorkerEachWorkTabl.Value2) + deltaWorker;
                        sumAllWork += allWorker;
                        if (sumAllWork - (int)sumAllWork < 0.1)
                        summaAmountofWorkerEachWork = (int)sumAllWork;
                        else summaAmountofWorkerEachWork = (int)sumAllWork + 1;
                        deltaWorker = (trudWork.Value2 - 8 * (int)(numberofWorkerEachWorkTabl.Value2) * (int)numberofDaysEachWorkTabl.Value2) / 8;
                        if (summaAmountofWorkerEachWork < amountofWorkerOnEachChapter)
                        {
                            Excel.Range firstFillColour = workSheetGraph.Cells[j, 7 + summaAmountofDaysEachWork];
                            Excel.Range lastFillColour = workSheetGraph.Cells[j, 7 + summaAmountofDaysEachWork + (int)(numberofDaysEachWorkTabl.Value2) - 1];
                            if (lastFillColour.Column > LastCellGraph.Column) lastFillColour = LastCellGraph;
                            Excel.Range rangeFillColour = workSheetGraph.get_Range(firstFillColour, lastFillColour);
                            rangeFillColour.Interior.ColorIndex = color;
                        }
                        else
                        {
                            Excel.Range firstFillColour = workSheetGraph.Cells[j, 7 + summaAmountofDaysEachWork];
                            Excel.Range lastFillColour = workSheetGraph.Cells[j, 7 + summaAmountofDaysEachWork + (int)(numberofDaysEachWorkTabl.Value2) - 1];
                            if (lastFillColour.Column > LastCellGraph.Column) lastFillColour = LastCellGraph;
                            Excel.Range rangeFillColour = workSheetGraph.get_Range(firstFillColour, lastFillColour);
                            rangeFillColour.Interior.ColorIndex = color;
                            summaAmountofDaysEachWork += (int)(numberofDaysEachWorkTabl.Value2);
                            if (summaAmountofDaysEachWork > amountOfDaysOnAllChapter) summaAmountofDaysEachWork -= 1; //бригада переходит на следующие работы в тот же день                          
                        }
                        if (sumAllWork > amountofWorkerOnEachChapter)
                            sumAllWork -= amountofWorkerOnEachChapter;
                    }
                }
            }
            FirstCellGraph = workSheetGraph.Cells[FirstCellGraph.Row + 2, FirstCellGraph.Column + 1];
            LastCellGraph = workSheetGraph.Cells[FirstCellGraph.Row + amountOfWorkInChapter + 1, FirstCellGraph.Column + 1];
            Excel.Range rangeCellsGrafik = workSheetGraph.get_Range(FirstCellGraph, LastCellGraph);
            rangeCellsGrafik.EntireColumn.HorizontalAlignment = Excel.Constants.xlLeft;
            workBookGraph.SaveAs(_graphAdress);
            object misValue = System.Reflection.Missing.Value;
            Marshal.FinalReleaseComObject(rangeCellsGrafik);
            Marshal.FinalReleaseComObject(workSheetGraph);
            workBookGraph.Close(true, misValue, misValue);
            Marshal.FinalReleaseComObject(workSheetSmeta);
            _oneSmeta.Close(true, misValue, misValue);
            Marshal.FinalReleaseComObject(_oneSmeta);
        }
            catch (COMException exc)
            {
                _textError+=$"{exc.Message} Закройте файл графика и повторите снова";
                return;
            }
            catch (NullReferenceException exc)
            {
                _textError += $"{exc.Message} Проверьте правильность написания Сметной трудоемкости";
                return;
            }
            catch (InvalidComObjectException exc)
            {
                _textError += $"{exc.Message} Закройте файл графика и повторите снова";
                return;
            }
        }
        //записывает в график строки, номер, наименование работ, трудозатраты, кол-во рабочих и кол-во дней
        public void RecordAllString(Excel.Worksheet workSheetSmeta, int amountOfWorkers, Excel.Worksheet workSheetGrafik, ref int amountOfWorkInChapter, ref int[] numChapterTablExcelGrafik, ref int[] amountOfWorkersInChapter)
        {
            Excel.Range firstCellAfterContent = workSheetGrafik.Range["B6"];
            int indexAmountWorkInChapter = 0, AmountofWorkerinEachWork = 0, numPosGrafik = 0,  indexChapt = 0;
            double reservPartOfDayAfterWork = 0;
            double[] trudozatratForChapter = _poRazdelyTrudozatrat.Values.ToArray();
            string[] valueNameOfEachWork = _nameForEachWorkinSmeta.Values.ToArray();
            double[] valueTrudozatratEachWork = _chelChasForEachWork.Values.ToArray();
            for (int i = 0; i < _allRazdelInOrder.Count; i++)
            {
                int indexAmountOfRowEachWorkinChapter=0;
                int[] keyNumPosSmetaChapterInOrder = _allRazdelInOrder[i].Keys.ToArray();
                int[] valueNumPosWorkChapterInOrder = _allRazdelInOrder[i].Values.ToArray();
                int daysOfEachWork;
                for (int r = 0; r < _cellsAllRazdel.Count; r++)
                {
                    indexAmountOfRowEachWorkinChapter = 0;

                    if (keyNumPosSmetaChapterInOrder[indexAmountOfRowEachWorkinChapter] == _startChapter[r] && _amounWorkInChapter[_startChapter[r]] == keyNumPosSmetaChapterInOrder.Length)
                    {
                       indexAmountOfRowEachWorkinChapter = 0;
                        string nameOfRazdel = _cellsAllRazdel[r].Value.ToString();
                        Console.WriteLine("valueNumPozWorkinRazdelpoPoradky.Length " + valueNumPosWorkChapterInOrder.Length);
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column] = ++numPosGrafik;
                        numChapterTablExcelGrafik[indexChapt++] = firstCellAfterContent.Row + amountOfWorkInChapter;
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 1] = nameOfRazdel;
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 2] = trudozatratForChapter[r];
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 3] = amountOfWorkers;
                        daysOfEachWork = (int)(trudozatratForChapter[r] / (amountOfWorkers * 8));
                        if (trudozatratForChapter[r] / (amountOfWorkers * 8) - daysOfEachWork > 0.05)
                        {
                            daysOfEachWork += 1;
                        }
                        if (daysOfEachWork == 0)
                        {
                            daysOfEachWork += 1;
                        }
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 4] = daysOfEachWork;

                    }
                    else if ((keyNumPosSmetaChapterInOrder[indexAmountOfRowEachWorkinChapter] == _startChapter[r] && _amounWorkInChapter[_startChapter[r]] > keyNumPosSmetaChapterInOrder.Length)||((r == _cellsAllRazdel.Count - 1 || keyNumPosSmetaChapterInOrder[indexAmountOfRowEachWorkinChapter] < _startChapter[r + 1]) && keyNumPosSmetaChapterInOrder[indexAmountOfRowEachWorkinChapter] > _startChapter[r] && _amounWorkInChapter[_startChapter[r]] > keyNumPosSmetaChapterInOrder.Length))
                    {
                        indexAmountOfRowEachWorkinChapter = 0;
                        string nameOfRazdel = _cellsAllRazdel[r].Value.ToString();
                        
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column] = ++numPosGrafik;
                        numChapterTablExcelGrafik[indexChapt++] = firstCellAfterContent.Row + amountOfWorkInChapter;
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 1] = nameOfRazdel;
                        double trudPartOfChapter = 0;
                        for (int q = 0; q < keyNumPosSmetaChapterInOrder.Length; q++)
                        {
                            trudPartOfChapter += _chelChasForEachWork[keyNumPosSmetaChapterInOrder[q]];
                        }

                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 2] = trudPartOfChapter;
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 3] = amountOfWorkers;
                        daysOfEachWork = (int)(trudozatratForChapter[r] / (amountOfWorkers * 8));
                        if (trudozatratForChapter[r] / (amountOfWorkers * 8) - daysOfEachWork > 0.05)
                        {
                            daysOfEachWork += 1;
                        }
         
                        if (daysOfEachWork == 0)
                        {
                            daysOfEachWork += 1;
                        }
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 4] = daysOfEachWork;

                    }
                 
                    else continue;
                    do
                        {
                            if (r < _cellsAllRazdel.Count - 1 && keyNumPosSmetaChapterInOrder[indexAmountWorkInChapter] >= _startChapter[r + 1])
                            {
                                break;
                            }
                            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter + indexAmountOfRowEachWorkinChapter + 1, firstCellAfterContent.Column] = ++numPosGrafik;
                            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter + indexAmountOfRowEachWorkinChapter + 1, firstCellAfterContent.Column + 1] = valueNameOfEachWork[valueNumPosWorkChapterInOrder[indexAmountWorkInChapter]];
                            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter + indexAmountOfRowEachWorkinChapter + 1, firstCellAfterContent.Column + 2] = valueTrudozatratEachWork[valueNumPosWorkChapterInOrder[indexAmountWorkInChapter]];
                            int amountOfWorkersInOneTime = 0;
                            do
                            {
                                amountOfWorkersInOneTime++;
                                if (valueTrudozatratEachWork[valueNumPosWorkChapterInOrder[indexAmountWorkInChapter]] > 8 * amountOfWorkers)
                                {
                                    AmountofWorkerinEachWork = amountOfWorkers;
                                    break;
                                }
                                if (valueTrudozatratEachWork[valueNumPosWorkChapterInOrder[indexAmountWorkInChapter]] <= 8 * amountOfWorkersInOneTime)
                                {
                                    AmountofWorkerinEachWork = amountOfWorkersInOneTime;
                                    break;
                                }
                            } while (amountOfWorkersInOneTime <= amountOfWorkers);
                            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter + indexAmountOfRowEachWorkinChapter + 1, firstCellAfterContent.Column + 3] = AmountofWorkerinEachWork;
                            daysOfEachWork = (int)(valueTrudozatratEachWork[valueNumPosWorkChapterInOrder[indexAmountWorkInChapter]] / (AmountofWorkerinEachWork * 8) );
                            reservPartOfDayAfterWork += valueTrudozatratEachWork[valueNumPosWorkChapterInOrder[indexAmountWorkInChapter]] / (AmountofWorkerinEachWork * 8) - daysOfEachWork;
                            if (reservPartOfDayAfterWork >= 1)
                            {
                               daysOfEachWork += 1;
                               reservPartOfDayAfterWork -= 1;
                            }
                            if (daysOfEachWork == 0)
                            {
                               daysOfEachWork += 1;
                             }                    
                            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter + indexAmountOfRowEachWorkinChapter + 1, firstCellAfterContent.Column + 4] = daysOfEachWork;
                            indexAmountWorkInChapter++;
                            indexAmountOfRowEachWorkinChapter++;
                            if (indexAmountWorkInChapter == valueNumPosWorkChapterInOrder.Length)
                            {
                                indexAmountWorkInChapter = 0;
                                break;
                            }
                        } while (indexAmountWorkInChapter > 0);
                        amountOfWorkInChapter += indexAmountOfRowEachWorkinChapter + 1;
                }
            }

        }
    }
}
