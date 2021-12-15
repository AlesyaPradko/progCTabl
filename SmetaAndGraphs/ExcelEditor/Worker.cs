using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditor.bl
{
    public abstract class Worker
    {
        protected List<string> _adresSmeta;
        protected List<string> _adresAktKS;
        public string _userAdresSmeta;
        protected string _userAdresKS;
        protected string _userAdresSave;
        protected Dictionary<string, List<string>> _aktAllKSforOneSmeta;

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
        public void Initialization(string userSmeta, string userKS, string userWhereSave)
        {
            _userAdresSmeta = userSmeta;
            _adresSmeta = ParserExc.GetstringAdres(_userAdresSmeta);
            _userAdresKS = userKS;
            _adresAktKS = ParserExc.GetstringAdres(_userAdresKS);
            _userAdresSave = userWhereSave;
            _userAdresSave += "\\Копия";
        }
        //копирует последовательно все сметы в выбранную папку
        private List<Excel.Workbook> MadeCopySmet(Excel.Application excelApp, List<Excel.Workbook> containFolderSmeta)
        {
            List<Excel.Workbook> containCopySmeta = new List<Excel.Workbook>();
            for (int u = 0; u < containFolderSmeta.Count; u++)
            {
                string testuserwheresave = _userAdresSave;
                testuserwheresave += $"{ _adresSmeta[u].Remove(0, _userAdresSmeta.Length + 1)}";//оставляет имя сметы(без пути)
                Excel.Workbook excelBookcopySmet = ParserExc.CopyExcelSmetaOne(_adresSmeta[u], testuserwheresave, excelApp);
                containCopySmeta.Add(excelBookcopySmet);
            }
            return containCopySmeta;
        }

        //метод для работы над папкой со сметами в разных режимах
        public void ProccessAll(RangeFile processingArea, Excel.Application excelApp, int size, ref string _textError)
        {
            try
            { 
            List<Excel.Workbook> containFolderSmeta = ParserExc.GetBookAllAktandSmeta(_userAdresSmeta, excelApp);
            if (containFolderSmeta.Count == 0 || _adresSmeta.Count == 0)
            {
                throw new DonthaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку");
            }
            int countDelete = 0;
            for (int i = 0; i < _adresSmeta.Count; i++)
            {
                if (!_adresSmeta[i].Contains("№"))
                {
                    object misValue = System.Reflection.Missing.Value;
                    containFolderSmeta[i].Close(false, misValue, misValue);
                    countDelete++;
                    throw new NullvalueException($"В названии сметы {_adresSmeta[i]} отсутствует символ № перед номером сметы\n");
                }
            }
            if (countDelete == _adresSmeta.Count) excelApp.Quit();
            List<Excel.Workbook> containFolderKS = ParserExc.GetBookAllAktandSmeta(_userAdresKS, excelApp);
            if (containFolderKS.Count == 0 || _adresAktKS.Count == 0)
            {
                throw new DonthaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку\n");
            }
            _aktAllKSforOneSmeta = ParserExc.GetContainAktKSinOneSmeta(containFolderKS, AdresSmeta, AdresAktKS);
            if (_aktAllKSforOneSmeta.Count != 0)
            {
                for (int numSmeta = 0; numSmeta < containFolderSmeta.Count; numSmeta++)
                {
                    List<Excel.Workbook> containCopySmeta = MadeCopySmet(excelApp, containFolderSmeta);
                    List<Excel.Workbook> listAktKStoOneSmeta = GetAllAktToOneSmeta(containFolderKS, numSmeta, ref _textError);
                    ProcessSmeta(listAktKStoOneSmeta, containCopySmeta[numSmeta], processingArea, containCopySmeta[numSmeta].FullName, size, ref _textError);
                }
            }
            else
            {
                object misValue = System.Reflection.Missing.Value;
                for (int i = 0; i < containFolderSmeta.Count; i++)
                {
                    containFolderSmeta[i].Close(false, misValue, misValue);
                }
                for (int i = 0; i < containFolderKS.Count; i++)
                {
                    containFolderKS[i].Close(false, misValue, misValue);
                }
                excelApp.Quit();
                throw new DonthaveExcelException("В актах КС отсутствует номер сметы или неверно записан\n");
            } 
        }
              catch (DonthaveExcelException ex)
            {
                _textError += ex.parName;
            }
            catch (NullvalueException exc)
            {
                _textError += exc.parName;
            }
        }
        private List<Excel.Workbook> GetAllAktToOneSmeta(List<Excel.Workbook> containFolderKS, int numSmeta,ref string _textError)
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
                        }
                    }
                }
                return listAktKStoOneSmeta;           
        }
        //метод переопределяется в классах-наследниках для работы над сметой в разных режимах
        protected abstract void ProcessSmeta(List<Excel.Workbook> listAktKStoOneSmeta, Excel.Workbook CopySmeta, RangeFile processingArea, string adresSmeta,int size,ref string _textError);
        //метод возвращает ячейку в которой хранится название Акта КС-2 и его дата составления
        protected Excel.Range FindCellforNameKS(Excel.Worksheet workSheetAktKS, Excel.Range findNumberOrDataKS)
        {
            //Console.WriteLine("FindCellforNameKS");
            if (!findNumberOrDataKS.MergeCells)
            {
                findNumberOrDataKS = workSheetAktKS.Cells[findNumberOrDataKS.Row + 1, findNumberOrDataKS.Column];
            }
            else
            {
                findNumberOrDataKS = workSheetAktKS.Cells[findNumberOrDataKS.Row + 2, findNumberOrDataKS.Column];
            }
            return findNumberOrDataKS;
        }

        //метод кругляет числа меньше 10 в -5 до значения 0
        protected void ZeroMinValue(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, int nextInsertColumn)
        {
            //Console.WriteLine("ObnulenieMinValue");
            for (int j = rangeSmetaOne.Row + 4; j < rangeSmetaOne.Rows.Count + rangeSmetaOne.Row; j++)
            {
                Excel.Range restFormula = SheetcopySmetaOne.Cells[j, nextInsertColumn];
                if (restFormula != null && restFormula.Value2 != null && restFormula.Value2.ToString() != "" && !restFormula.MergeCells)
                {
                    double d = Convert.ToDouble(restFormula.Value2);
                    if (d < 0.00001)
                    {
                        restFormula.Value2 = 0;
                    }
                }
            }
        }

        protected void FormatRecordCopySmeta(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, string adresSmeta,int size,ref string _textError)
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
             if (lastCellFormat.Column >= rangeSmetaOne.Columns.Count)
             {
                 throw new ZapredelException($"Вы задали слишком малую ширину для {adresSmeta}\n");                    //return;
             }
             Excel.Range firstCellFormat = SheetcopySmetaOne.Cells[rangeSmetaOne.Row, rangeSmetaOne.Column];
             Excel.Range formarRange = SheetcopySmetaOne.get_Range(firstCellFormat, lastCellFormat);
             formarRange.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
             formarRange.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
             formarRange.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
             formarRange.EntireColumn.Font.Size = size;
             formarRange.EntireColumn.Font.FontStyle = "Times New Roman";
             formarRange.EntireColumn.AutoFit();
             Excel.Range lastCellwithAnotherWidth = SheetcopySmetaOne.Cells[lastCellFormat.Row, rangeSmetaOne.Column];
             Excel.Range rangewithAnotherWidth = SheetcopySmetaOne.get_Range(firstCellFormat, lastCellwithAnotherWidth);
             rangewithAnotherWidth.ColumnWidth = 12;
        }       
    }
}
