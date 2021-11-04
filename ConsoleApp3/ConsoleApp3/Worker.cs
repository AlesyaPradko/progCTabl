using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp3
{
    public class Worker
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
        public List<Excel.Workbook> ContainPapkaSmeta { get; set; }
        public List<string> AdresSmeta { get; set; }
        public List<string> AdresKS { get; set; }    
        public List<Excel.Workbook> CopySmet { get; set; }
        public List<Excel.Workbook> ContainPapkaKS { get; set; }
        public Dictionary<string, List<string>> KskSmete { get; set; }
      
        public  void Zakrutie(List<string> adrSm, Dictionary<string, List<string>> ksSm, List<Excel.Workbook> contPapKS, List<string> adrKS, List<Excel.Workbook> copS,int n, Excel.Worksheet exc, Excel.Range r)
        {
            Console.WriteLine("Zakrutie");
            object misValue = System.Reflection.Missing.Value;     
            for (int v = 0; v < ksSm[adrSm[n]].Count; v++)
            {
                for (int i = 0; i < contPapKS.Count; i++)
                {
                    if (adrKS[i] == ksSm[adrSm[n]][v])
                    {
                        contPapKS[i].Close(false, misValue, misValue);
                        Marshal.FinalReleaseComObject(contPapKS[i]);
                    }
                    else continue;
                }
            }       
            Marshal.FinalReleaseComObject(r);
            Marshal.FinalReleaseComObject(exc);
            copS[n].Close(true, misValue, misValue);
            Marshal.FinalReleaseComObject(copS[n]);
        }
    }
}