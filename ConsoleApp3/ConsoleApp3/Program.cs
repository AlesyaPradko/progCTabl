using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
public enum ChangeMod { expert=49, tehnadzor=50 };

namespace ConsoleApp3
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Workbook excelBookcopy = ParserExc.CopyExcel(@"D:\Книга 6.xlsx", @"D:\икси 2\Копия сметы.xlsx");
            Excel.Worksheet Sheetcopy;
            Sheetcopy = excelBookcopy.Sheets[1];
            Excel.Range rangecopy;
            rangecopy = Sheetcopy.get_Range("A1", "L40");
            Excel.Range firkeysm = rangecopy.Find("Номер") as Excel.Range;
            Excel.Range firvalsm = rangecopy.Find("Количество") as Excel.Range;
            List<Excel.Workbook> ListKS = ParserExc.GetListKS(@"D:\икси");
            Console.WriteLine("Выберите режим эксперт(нажмите 1) или техназор(нажмите 2)");
            ChangeMod chan;
            int changeregim = (int)(Console.ReadKey().Key); 
            chan = (ChangeMod)changeregim;
            switch (chan)
            {
                case ChangeMod.expert: ParserExc.Worklikeexpert(ListKS,Sheetcopy, rangecopy, firkeysm, firvalsm);break;
                case ChangeMod.tehnadzor: ParserExc.Workliketehnadzor(ListKS, Sheetcopy, rangecopy,firkeysm, firvalsm); break;
                default:
                    Console.WriteLine("Вы ввели неверный символ ");
                    break;
            }
           
            for (int i = 1; i <= rangecopy.Rows.Count; i++)
            {       
                Console.Write("\r\n");
                for (int j = 1; j <= rangecopy.Columns.Count; j++)
                {
                    Excel.Range forYach = Sheetcopy.Cells[i, j] as Excel.Range;           
                    if (forYach != null && forYach.Value2 != null)
                        Console.Write(forYach.Value2.ToString() + "\t");
                }
            }
            object misValue = System.Reflection.Missing.Value;
            for (int i = 0; i < ListKS.Count; i++)
            {
                ListKS[i].Close(false, misValue, misValue);
                Marshal.FinalReleaseComObject(ListKS[i]);
            }
            Marshal.FinalReleaseComObject(rangecopy);
            Marshal.FinalReleaseComObject(Sheetcopy);
            excelBookcopy.Close(true, misValue, misValue);
            Marshal.FinalReleaseComObject(excelBookcopy);
            ParserExc.Proverka().Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Console.ReadLine();
        }
    }
}