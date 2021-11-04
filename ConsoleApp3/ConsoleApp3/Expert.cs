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

namespace ConsoleApp3
{
    public class Expert:Worker
    {
        private Dictionary<int, double> summSmeta;
        private Dictionary<int, string> periodSmeta;
        public Expert():base()
        { } 
        public  void Worklikeexpert(int num, List<Excel.Workbook> copyS, List<string> adresSm, List<string> adrKS, List<Excel.Workbook> containPapKS, Dictionary<string, List<string>> kskSm,string s1, string s2)
        {
            //Console.WriteLine("Worklikeexpert");
            try
            {
                Excel.Worksheet Sheetc;
                Sheetc = copyS[num].Sheets[1];
                Excel.Range rangec;
                rangec = Sheetc.get_Range(s1, s2);
                Excel.Range firkeys = rangec.Find("Номер") as Excel.Range;
                Excel.Range firvalsm = rangec.Find("Количество") as Excel.Range;
                int a = firvalsm.Column + 1;
                Excel.Range forYs = Sheetc.Cells[firkeys.Row, a] as Excel.Range;
                forYs.Insert(XlInsertShiftDirection.xlShiftToRight);
                summSmeta = ParserExc.Getkeysm<double>(Sheetc, rangec, firkeys);
                periodSmeta = ParserExc.Getkeysm<string>(Sheetc, rangec, firkeys);
                Sheetc.Cells[firkeys.Row, a] = "Выполнение по смете";
                int dprimst = ParserExc.GetColumzapis(Sheetc, rangec, firkeys);
                if (dprimst == -1) { throw new ZapredelException("Вы задали слишком малую область по ширине таблицы, задайте большую"); return; }
                int[] keysm = summSmeta.Keys.ToArray(); 
                double[] valsm = summSmeta.Values.ToArray();
                string[] valstr = periodSmeta.Values.ToArray();
                string[] sk = new string[containPapKS.Count];
                int numKS = 0;   
                for (int v = 0; v < kskSm[adresSm[num]].Count; v++)
                {
                   for (int i = 0; i < containPapKS.Count; i++)
                   {
                       if (adrKS[i] != kskSm[adresSm[num]][v]) continue;
                       else
                       {
                           sk[i] = "Акт КС-2 №";
                           Excel.Worksheet workSheet;
                           workSheet = containPapKS[i].Sheets[1];
                           Excel.Range range;
                           range = workSheet.get_Range(s1, s2);
                           RegexReg regul = new RegexReg();
                           Excel.Range firstkey = range.Find("по смете") as Excel.Range;
                           Excel.Range otregexval = ParserExc.GetCell(workSheet, range, regul.regexval);
                           Excel.Range findnum = range.Find("Номер документа") as Excel.Range;
                           if (!findnum.MergeCells)
                            { findnum = workSheet.Cells[findnum.Row + 1, findnum.Column] as Excel.Range; }
                           else 
                            { findnum = workSheet.Cells[findnum.Row + 2, findnum.Column] as Excel.Range; }
                           Excel.Range finddata = range.Find("Дата составления") as Excel.Range;
                           if (!finddata.MergeCells)
                            { finddata = workSheet.Cells[finddata.Row + 1, finddata.Column] as Excel.Range; }
                           else 
                            { finddata = workSheet.Cells[finddata.Row + 2, finddata.Column] as Excel.Range; }
                           sk[i] += findnum.Value.ToString();
                           string sgod = ParserExc.Finddate(regul.regexgod, finddata);
                           string smes = ParserExc.Finddate(regul.regexmes, finddata);
                           string havemes = ParserExc.Mespropis(smes);
                           havemes += sgod;
                           sk[i] += " ";
                           sk[i] += havemes;
                           sk[i] += "\n";
                           getVupoln = ParserExc.Getvupoln(workSheet, range, firstkey, otregexval);
                           ICollection keyColl = getVupoln.Keys;
                           ICollection valColl = getVupoln.Values;
                           bool eqva;
                           int ob1;
                           for (int j = firkeys.Row + 2; j <= rangec.Rows.Count; j++)
                           {
                               eqva = false;
                               int ind = 0;
                               Excel.Range forY4 = Sheetc.Cells[j, firkeys.Column] as Excel.Range;
                               if (forY4 != null && forY4.Value2 != null && forY4.Value2.ToString() != "" && !forY4.MergeCells)
                               {
                                   ob1 = Convert.ToInt32(forY4.Value2);
                                   foreach (int ob in keyColl)
                                   {
                                       if (ob1 == ob)
                                       {
                                           eqva = true;
                                           ind = Array.IndexOf(keysm, ob);
                                           valsm[ind] += getVupoln[ob];
                                           summSmeta[ob] = valsm[ind];
                                           Sheetc.Cells[j, a] = summSmeta[ob];
                                       }
                                   }
                                   if (eqva)
                                   {
                                        valstr[ind] += sk[i];
                                        valstr[ind] += " ";
                                        Sheetc.Cells[j, dprimst] = valstr[ind];
                                   }
                               }
                           }
                           Marshal.FinalReleaseComObject(range);
                           Marshal.FinalReleaseComObject(workSheet);
                       }
                   }
                }
                //вставка столбца "Остток" с формулой разности
                Excel.Range Ost = Sheetc.Cells[firkeys.Row, a + 1] as Excel.Range;
                Ost.Insert(XlInsertShiftDirection.xlShiftToRight);
                Sheetc.Cells[firkeys.Row, a+1] = "Остаток";
                for (int j = firvalsm.Row + 1; j <= rangec.Rows.Count; j++)
                {
                    Excel.Range vst = Sheetc.Cells[j, a+1] as Excel.Range;
                    vst.Insert(XlInsertShiftDirection.xlShiftToRight);
                    if (j > firvalsm.Row + 1)
                    {
                        Excel.Range act = Sheetc.Cells[j, firvalsm.Column] as Excel.Range;
                        if (act != null && act.Value2 != null && act.Value2.ToString() != "" && !act.MergeCells)
                        {
                            Excel.Range activ = Sheetc.Cells[j, a+1] as Excel.Range;
                            activ.FormulaR1C1 = "=RC[-2]-RC[-1]";
                        }
                    }
                }
                ParserExc.FormatZapis(num, adresSm, numKS, adrKS, firkeys, firvalsm, Sheetc, rangec);
                ////вывод в консоль до окончательной отладки
                //for (int i = 1; i <= rangec.Rows.Count; i++)
                //{
                //    Console.Write("\r\n");
                //    for (int j = 1; j <= rangec.Columns.Count; j++)
                //    {
                //        Excel.Range forYach = Sheetc.Cells[i, j] as Excel.Range;
                //        if (forYach != null && forYach.Value2 != null)
                //            Console.Write(forYach.Value2.ToString() + "\t");
                //    }
                //}
                Zakrutie(adresSm, kskSm, containPapKS,adrKS, copyS,num, Sheetc,rangec);
            }
            catch (ZapredelException exc)
            { Console.WriteLine(exc.parName); }
        }
    }
}