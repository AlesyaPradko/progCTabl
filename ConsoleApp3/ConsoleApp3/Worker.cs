using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp3
{
    public abstract class Worker
    {
        protected Dictionary<int, double> getVupoln;
        protected List<string> adresSmeta;
        protected List<string> adresKS;
        protected List<Excel.Workbook> containPapkaKS;
        private List<Excel.Workbook> containPapkaSmeta;
        protected List<Excel.Workbook> copySmet;
        protected Dictionary<string, List<string>> kskSmete;
        public Worker()
        { }

  
        public List<Excel.Workbook> ContainPapkaSmeta 
        {
            get { return containPapkaSmeta; }
            private set { } 
        }
        public List<string> AdresSmeta
        {
            get { return adresSmeta; }
            private set { }
        }
        public List<string> AdresKS
        {
            get { return adresKS; }
            private set { }
        }
        public List<Excel.Workbook> ContainPapkaKS
        {
            get { return containPapkaKS; }
            private set { }
        }
        //public List<Excel.Workbook> CopySmet { get; set; }

        //public Dictionary<string, List<string>> KskSmete { get; set; }

        public void Initialization(Excel.Application excelApp)
        {
            //возможно стоит отправить аргументов, пользователь сам выбирает папку
            string usersmetu = @"D:\иксу";
            containPapkaSmeta = ParserExc.GetListKS(usersmetu, excelApp);
            adresSmeta = ParserExc.Getstring(usersmetu);
            if (containPapkaSmeta.Count == 0 || adresSmeta.Count == 0)
            {
                throw new DonthaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку");
            }
            //возможно стоит отправить аргументов, пользователь сам выбирает папку
            string userKS = @"D:\икси";
            adresKS = ParserExc.Getstring(userKS);
            //возможно стоит отправить аргументов, пользователь сам выбирает папку
            string userwheresave = @"D:\икси 2";
            userwheresave += "\\Копия";
            copySmet = ParserExc.MadeCopyExcbook(userwheresave, usersmetu, excelApp, ContainPapkaSmeta, AdresSmeta);
            containPapkaKS = ParserExc.GetListKS(userKS, excelApp);
            if (containPapkaKS.Count == 0 || adresKS.Count == 0)
            {
                throw new DonthaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку");
            }
            kskSmete = ParserExc.GetContainSM(ContainPapkaKS, AdresSmeta, AdresKS);
        }
        public void ProccessAll(RangeFile obl)
        {
            for (int num = 0; num < copySmet.Count; num++)
            {
                ProcessSmeta(num, obl);
            }
        }
        
        protected abstract void ProcessSmeta(int num, RangeFile obl);

        protected Excel.Range FindCellforNameKS(Excel.Range f, Excel.Worksheet workSheet)
        {
            if (!f.MergeCells)
            {
                f = workSheet.Cells[f.Row + 1, f.Column];
            }
            else
            {
                f = workSheet.Cells[f.Row + 2, f.Column];
            }
            return f;
        }

        //метод задает формат записи в файл иксель

        //Console.WriteLine("FormatZapis");
        public void FormatZapis(int n, int i,Excel.Range firkeysm, Excel.Worksheet Sheetcopy, Excel.Range rangecopy)
        {
            int shir = 0, vus = 0, test = 0;
            for (int y = firkeysm.Row; y <= rangecopy.Rows.Count; y++)
            {
                 Excel.Range fors = Sheetcopy.Cells[y, firkeysm.Column];
                 if (fors != null && fors.Value2 != null) vus++;
                 else 
                 {
                     test++; 
                     if (fors != null && fors.Value2 != null)
                     {
                         vus += test; 
                         test = 0;
                     } 
                     test = 0;
                 }
                 if (test > 5) break;
             }
             test = 0;
             for (int x = firkeysm.Column; x <= rangecopy.Columns.Count; x++)
             {
                 Excel.Range fors = Sheetcopy.Cells[firkeysm.Row, x];
                 if (fors != null && fors.Value2 != null) shir++;
                 else 
                 { 
                     test++;
                     if (fors != null && fors.Value2 != null)
                     { 
                         shir += test; 
                         test = 0;
                     } 
                 }
                 if (test > 5) break;
             }
             Excel.Range flast = Sheetcopy.Cells[firkeysm.Row + vus + 1, firkeysm.Column + shir];
             if (flast.Column == rangecopy.Columns.Count || flast.Row == rangecopy.Rows.Count)
             { 
                 throw new ZapredelException($"Вы задали слишком малую область для {adresSmeta[n]} и для {adresKS[i]}");
                 return;
             }
             Excel.Range forIs = Sheetcopy.get_Range(firkeysm, flast);
             forIs.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
             forIs.EntireColumn.Font.Size = 13;
             forIs.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
             forIs.EntireColumn.AutoFit();
        }
     
        public void Zakrutie(int n, Excel.Worksheet Sheetcopy, Excel.Range rangecopy)
        {
            Console.WriteLine("Zakrutie");
            object misValue = System.Reflection.Missing.Value;     
            for (int v = 0; v < kskSmete[adresSmeta[n]].Count; v++)
            {
                for (int i = 0; i < containPapkaKS.Count; i++)
                {
                    if (adresKS[i] == kskSmete[adresSmeta[n]][v])
                    {
                        containPapkaKS[i].Close(false, misValue, misValue);
                        Marshal.FinalReleaseComObject(containPapkaKS[i]);
                    }
                    else continue;
                }
            }       
            Marshal.FinalReleaseComObject(rangecopy);
            Marshal.FinalReleaseComObject(Sheetcopy);
            copySmet[n].Close(true, misValue, misValue);
            Marshal.FinalReleaseComObject(copySmet[n]);
        }
    }
}