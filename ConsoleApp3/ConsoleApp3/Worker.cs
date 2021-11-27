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
        protected Dictionary<int, double> totalScopeWorkAktKSone;
        protected List<string> adresSmeta;
        protected List<string> adresAktKS;
        protected List<Excel.Workbook> containPapkaKS;
        protected List<Excel.Workbook> containCopySmeta;
        protected Dictionary<string, List<string>> aktAllKSforOneSmeta;

        private List<Excel.Workbook> _containPapkaSmeta;
        private Mutex _mutex = new Mutex();

        public Task[] taskobrabotka;

        public Worker()
        { }


        public List<Excel.Workbook> ContainPapkaSmeta
        {
            get { return _containPapkaSmeta; }
            private set { _containPapkaSmeta = value; } //!!! здесь дописал set
        }
        public List<string> AdresSmeta
        {
            get { return adresSmeta; }
            private set { adresSmeta = value; }//!!! здесь дописал set
        }
        public List<string> AdresAktKS
        {
            get { return adresAktKS; }
            private set { adresAktKS = value; }//!!! здесь дописал set
        }
        public List<Excel.Workbook> ContainPapkaKS
        {
            get { return containPapkaKS; }
            private set { containPapkaKS = value; }//!!! здесь дописал set
        }
        //public List<Excel.Workbook> CopySmet { get; set; }

        //public Dictionary<string, List<string>> KskSmete { get; set; }

        //метод инициализирует листы и словари хранящие в себе сметы, Акты КС-2 (адреса и книги)
        public void Initialization(Excel.Application excelApp)
        {
            // выбор в проводнике, пользователь сам выбирает папку
            //Console.WriteLine("Initialization");
            try
            {
                string usersmeta = @"D:\иксу";
                _containPapkaSmeta = ParserExc.GetBookAllAktKSandSmeta(usersmeta, excelApp);
                adresSmeta = ParserExc.GetstringAdresa(usersmeta);
                if (_containPapkaSmeta.Count == 0 || adresSmeta.Count == 0)
                {
                    throw new DonthaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку");
                }
                // свыбор в проводнике, пользователь сам выбирает папку
                string userKS = @"D:\икси";
                adresAktKS = ParserExc.GetstringAdresa(userKS);
                // выбор в проводнике, пользователь сам выбирает папку
                string userwheresave = @"D:\икси 2";
                userwheresave += "\\Копия";
                containCopySmeta = MadeCopyAllCopySmet(userwheresave, usersmeta, excelApp);
                containPapkaKS = ParserExc.GetBookAllAktKSandSmeta(userKS, excelApp);
                if (containPapkaKS.Count == 0 || adresAktKS.Count == 0)
                {
                    throw new DonthaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку");
                }
                aktAllKSforOneSmeta = ParserExc.GetContainAktKSinOneSmeta(ContainPapkaKS, AdresSmeta, AdresAktKS);
            }
            catch (DonthaveExcelException ex)
            {
                Console.WriteLine(ex.parName);
            }
        }
        //копирует последовательно все сметы в выбранную папку
        private List<Excel.Workbook> MadeCopyAllCopySmet(string userwheresave, string usersmeta, Excel.Application excelApp)
        {
            //Console.WriteLine("MadeCopyExcbook");
            List<Excel.Workbook> containCopySmeta = new List<Excel.Workbook>();
            for (int u = 0; u < _containPapkaSmeta.Count; u++)
            {
                string testuserwheresave = userwheresave;
                testuserwheresave += $"{ adresSmeta[u].Remove(0, usersmeta.Length + 1)}";//оставляет имя сметы(без пути)
                Excel.Workbook excelBookcopySmet = ParserExc.CopyExcelSmetaOne(adresSmeta[u], testuserwheresave, excelApp);
                containCopySmeta.Add(excelBookcopySmet);
            }
            return containCopySmeta;
        }

        //метод для работы над папкой со сметами в разных режимах
        public void ProccessAll(RangeFile oblastobrabotki)
        {
            taskobrabotka = new Task[containCopySmeta.Count];
            for (int num = 0; num < containCopySmeta.Count; num++)
            {

                taskobrabotka[num] = Task.Factory.StartNew(() =>
                {
                    _mutex.WaitOne();
                    ProcessSmeta(oblastobrabotki);
                    _mutex.ReleaseMutex();
                    //ProcessSmeta(oblastobrabotki);
                    // Долго соображал почему здесь вызывается один метод с одним аргументом. Как я понял в самом методе используется индекс таски и от него уже выбираются определенные данные
                    // Нужно в ProcessSmeta передавть объект,содержищий только одну смету, а не весь файл, а потом уже внутри из него что-то выбирать с учетом индексов.
                    //Нужен некоторый метод, который будет выбирать из целого файла данные для одной сметы, например, GetEstimateRawData(oblastobrabotki,num) //ПолучитьСырыеДанныеСметы
                    //тогда код может быть
                    //_mutex.WaitOne();
                    //RangeFile estimateData = GetEstimateRawData(oblastobrabotki, num);
                    //_mutex.ReleaseMutex();
                    //ProcessSmeta(estimateData); //Эту штуку можно будет запускать без мьютекса т.к. она назевисима. Если у Вас вся обработка будет завернута в мьютекс, смысла в потоках
                    //не будет т.к. все равно все будет последовательно.
                    //Мьютекс остается один и только здесь. Поудалял работу с массивом мьютексов, если еще где-то осталось - все удалем
                    //Если подготовить данные не получится таким образом, то смысла в параллельности опять же не будет, т.к. исходник на всех один и все будет под мьютексом параллельно


                });
            }
            Task.WaitAll(taskobrabotka);
        }

        //метод переопределяется в классах-наследниках для работы над сметой в разных режимах
        protected abstract void ProcessSmeta(RangeFile oblastobrabotki);
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

        protected void FormatZapisinCopySmeta(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, int numSmeta)
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
                    throw new ZapredelException($"Вы задали слишком малую ширину для {adresSmeta[numSmeta]}");
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
        protected void Zakrutie(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, int numSmeta)
        {
            Console.WriteLine("Zakrutie");
            object misValue = System.Reflection.Missing.Value;
            for (int v = 0; v < aktAllKSforOneSmeta[adresSmeta[numSmeta]].Count; v++)
            {
                for (int i = 0; i < containPapkaKS.Count; i++)
                {
                    if (adresAktKS[i] == aktAllKSforOneSmeta[adresSmeta[numSmeta]][v])
                    {
                        containPapkaKS[i].Close(false, misValue, misValue);
                        Marshal.FinalReleaseComObject(containPapkaKS[i]);
                    }
                    else continue;
                }
            }
            Marshal.FinalReleaseComObject(rangeSmetaOne);
            Marshal.FinalReleaseComObject(SheetcopySmetaOne);
            containCopySmeta[numSmeta].Close(true, misValue, misValue);
            Marshal.FinalReleaseComObject(containCopySmeta[numSmeta]);
        }
    }
}