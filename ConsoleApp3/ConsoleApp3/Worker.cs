using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp3
{
    public abstract class Worker 
    {
       
        protected List<string> _adresSmeta;
        protected List<string> _adresAktKS;
        protected string _userSmeta;
        protected string _userKS;
        protected string _userWhereSave;
        protected Dictionary<string, List<string>> _aktAllKSforOneSmeta;
        public Task[] _taskobrabotka;
        //private Mutex _mutex = new Mutex();

        public Worker()
        { }

        public List<string> AdresSmeta
        {
            get { return _adresSmeta; }
            private set { }
        }
        public List<string> AdresAktKS
        {
            get { return _adresAktKS; }
            private set { }
        }
 

        //метод инициализирует листы и словари хранящие в себе сметы, Акты КС-2 (адреса и книги)
        public void Initialization()
        {
            // выбор в проводнике, пользователь сам выбирает папку
            //Console.WriteLine("Initialization");      
                _userSmeta = @"D:\иксу";             
                _adresSmeta = ParserExc.GetstringAdresa(_userSmeta);            
                // свыбор в проводнике, пользователь сам выбирает папку
                _userKS = @"D:\икси";
                _adresAktKS = ParserExc.GetstringAdresa(_userKS);
                // выбор в проводнике, пользователь сам выбирает папку
                _userWhereSave = @"D:\икси 2";
                _userWhereSave += "\\Копия";
        }
        //копирует последовательно все сметы в выбранную папку
        private List<Excel.Workbook> MadeCopyAllCopySmet(Excel.Application excelApp, List<Excel.Workbook> containFolderSmeta)
        {
            //Console.WriteLine("MadeCopyExcbook");
            List<Excel.Workbook> containCopySmeta = new List<Excel.Workbook>();
            for (int u = 0; u < containFolderSmeta.Count; u++)
            {
                string testuserwheresave = _userWhereSave;
                testuserwheresave += $"{ _adresSmeta[u].Remove(0, _userSmeta.Length + 1)}";//оставляет имя сметы(без пути)
                Console.WriteLine("testuserwheresave " +testuserwheresave);
                Excel.Workbook excelBookcopySmet = ParserExc.CopyExcelSmetaOne(_adresSmeta[u], testuserwheresave, excelApp);
                containCopySmeta.Add(excelBookcopySmet);
                Console.WriteLine("excelBookcopySmet.FullName " + excelBookcopySmet.FullName);
            }
            return containCopySmeta;
        }

        //метод для работы над папкой со сметами в разных режимах
        public void ProccessAll(RangeFile processingArea, Excel.Application excelApp)
        {
            //лист мсеты получение
            try
            {
                List<Excel.Workbook> containFolderSmeta = ParserExc.GetBookAllAktKSandSmeta(_userSmeta, excelApp);
                if (containFolderSmeta.Count == 0 || _adresSmeta.Count == 0)
                {
                    throw new DonthaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку");
                }
                List<Excel.Workbook> containCopySmeta = MadeCopyAllCopySmet(excelApp, containFolderSmeta);
                List<Excel.Workbook> containFolderKS = ParserExc.GetBookAllAktKSandSmeta(_userKS, excelApp);
                if (containFolderKS.Count == 0 || _adresAktKS.Count == 0)
                {
                    throw new DonthaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку");
                }
                _aktAllKSforOneSmeta = ParserExc.GetContainAktKSinOneSmeta(containFolderKS, AdresSmeta, AdresAktKS);
                _taskobrabotka = new Task[containCopySmeta.Count];
                for (int numSmeta = 0; numSmeta < containCopySmeta.Count; numSmeta++)
                {
                    _taskobrabotka[numSmeta] = Task.Factory.StartNew(() =>
                    {
                       
                        Console.WriteLine(Task.CurrentId + "начал  работу");
                        //_mutex.WaitOne();
                        Excel.Worksheet sheetCopySmeta = containCopySmeta[(int)Task.CurrentId-1].Sheets[1];
                        //получение листа кс к смете
                        List<Excel.Workbook> listAktKStoOneSmeta = GetAllAktKStoOneSmeta(containFolderKS, (int)Task.CurrentId - 1);
                        //_mutex.ReleaseMutex();
                        ProcessSmeta(listAktKStoOneSmeta, sheetCopySmeta, processingArea,_adresSmeta[(int)Task.CurrentId - 1]);
                        Zakrutie(listAktKStoOneSmeta, containCopySmeta[(int)Task.CurrentId - 1], sheetCopySmeta);
                        Console.WriteLine(Task.CurrentId + "завершил  работу");
                       
                    });
                }
                Task.WaitAll(_taskobrabotka);
            }
            catch (DonthaveExcelException ex)
            {
                Console.WriteLine(ex.parName);
            }

}
        private List<Excel.Workbook> GetAllAktKStoOneSmeta(List<Excel.Workbook> containFolderKS, int numSmeta)
        {
            List<Excel.Workbook> listAktKStoOneSmeta = new List<Excel.Workbook>();
            for (int v = 0; v < _aktAllKSforOneSmeta[_adresSmeta[numSmeta]].Count; v++)
            {
                for (int numKS = 0; numKS < containFolderKS.Count; numKS++)
                {
                    if (_adresAktKS[numKS] != _aktAllKSforOneSmeta[_adresSmeta[numSmeta]][v]) continue;
                    else
                    {
                        listAktKStoOneSmeta.Add(containFolderKS[numKS]);
                        Console.WriteLine("containFolderKS[numKS] " + containFolderKS[numKS].FullName);
                    }
                }
            }
            return listAktKStoOneSmeta;
        }
        //метод переопределяется в классах-наследниках для работы над сметой в разных режимах
        protected abstract void ProcessSmeta(List<Excel.Workbook> listAktKStoOneSmeta, Excel.Worksheet sheetCopySmeta, RangeFile processingArea,string adresSmeta);
        //метод возвращает ячейку в которой хранится название Акта КС-2 и его дата составления
        protected Excel.Range FindCellforNameKS(Excel.Worksheet workSheetAktKS, Excel.Range findNomerorDataKS)
        {
            //Console.WriteLine("FindCellforNameKS");
            if (!findNomerorDataKS.MergeCells)
            {
                findNomerorDataKS = workSheetAktKS.Cells[findNomerorDataKS.Row + 1, findNomerorDataKS.Column];
                Console.WriteLine(findNomerorDataKS.Value.ToString());
            }
            else
            {
                findNomerorDataKS = workSheetAktKS.Cells[findNomerorDataKS.Row + 2, findNomerorDataKS.Column];
                Console.WriteLine(findNomerorDataKS.Value.ToString());
            }
            return findNomerorDataKS;
        }

        //метод кругляет числа меньше 10 в -5 до значения 0
        protected void ObnulenieMinValue(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, int nextInsertColumn)
        {
            //Console.WriteLine("ObnulenieMinValue");
            for (int j = rangeSmetaOne.Row + 4; j < rangeSmetaOne.Rows.Count + rangeSmetaOne.Row; j++)
            {
                Excel.Range ostatocFormula = SheetcopySmetaOne.Cells[j, nextInsertColumn];
                if (ostatocFormula != null && ostatocFormula.Value2 != null && ostatocFormula.Value2.ToString() != "" && !ostatocFormula.MergeCells)
                {
                    double d = Convert.ToDouble(ostatocFormula.Value2);
                    if (d < 0.00001)
                    {
                        ostatocFormula.Value2 = 0;
                    }
                }
            }
        }

        protected void FormatZapisinCopySmeta(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, string adresSmeta)
        {
            //Console.WriteLine("FormatZapisinCopySmeta");
            try
            {
                int widthTabl = 0, testEmptyCells = 0;
                for (int x = rangeSmetaOne.Column; x < rangeSmetaOne.Columns.Count + rangeSmetaOne.Column; x++)
                {
                    Excel.Range cellsFirstRowTabl = SheetcopySmetaOne.Cells[rangeSmetaOne.Row, x];
                    if (cellsFirstRowTabl != null && cellsFirstRowTabl.Value2 != null && cellsFirstRowTabl.Value2.ToString() != "")
                    {
                        widthTabl++;
                    }
                    else
                    {
                        testEmptyCells++;
                        if (cellsFirstRowTabl != null && cellsFirstRowTabl.Value2 != null && cellsFirstRowTabl.Value2.ToString() != "")
                        {
                            widthTabl += testEmptyCells;
                            testEmptyCells = 0;
                        }
                    }
                    if (testEmptyCells > 5) break;
                }
                Excel.Range lastCellFormat = SheetcopySmetaOne.Cells[rangeSmetaOne.Rows.Count + rangeSmetaOne.Row - 1, rangeSmetaOne.Column + widthTabl - 1];
                Console.WriteLine(" Column last " + lastCellFormat.Column);
                if (lastCellFormat.Column >= rangeSmetaOne.Columns.Count)
                {
                    throw new ZapredelException($"Вы задали слишком малую ширину для {adresSmeta}");
                    //return;
                }
                Excel.Range firstCellFormat = SheetcopySmetaOne.Cells[rangeSmetaOne.Row, rangeSmetaOne.Column];
                Excel.Range formarRange = SheetcopySmetaOne.get_Range(firstCellFormat, lastCellFormat);
                formarRange.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                //formarRange.EntireColumn.Font.Size = 11;
                formarRange.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
                formarRange.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
                formarRange.EntireColumn.AutoFit();
                Excel.Range lastCellwithAnotherWidth = SheetcopySmetaOne.Cells[lastCellFormat.Row, rangeSmetaOne.Column];
                Excel.Range rangewithAnotherWidth = SheetcopySmetaOne.get_Range(firstCellFormat, lastCellwithAnotherWidth);
                rangewithAnotherWidth.ColumnWidth = 12;
            }
            catch (ZapredelException exc)
            {
                Console.WriteLine(exc.parName);
            }
        }
        //метод закрывает открытые файлы КС
        protected void Zakrutie(List<Excel.Workbook>  listAktKStoOneSmeta, Excel.Workbook copySmeta, Excel.Worksheet SheetcopySmetaOne)
        {
            Console.WriteLine("Zakrutie");
            object misValue = System.Reflection.Missing.Value;
            for (int i = 0; i < listAktKStoOneSmeta.Count; i++)
            {

                listAktKStoOneSmeta[i].Close(false, misValue, misValue);
   
            } 
            Marshal.FinalReleaseComObject(SheetcopySmetaOne);
            copySmeta.Close(true, misValue, misValue);
            Marshal.FinalReleaseComObject(copySmeta);
        }
    }
}