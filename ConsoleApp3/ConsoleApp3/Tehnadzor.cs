using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp3
{
    public class Tehnadzor: Worker
    {
        public Tehnadzor():base()
        { }
        protected override void ProcessSmeta(int num, string sx1, string sx2)
        {
            //Console.WriteLine("Workliketehnadzor");
            Excel.Worksheet Sheetcopy;
            Sheetcopy = CopySmet[num].Sheets[1];
            Excel.Range rangecopy;
            rangecopy = Sheetcopy.get_Range(sx1, sx2);
            Excel.Range firkeysm = rangecopy.Find("Номер") as Excel.Range;
            Excel.Range firvalsm = rangecopy.Find("Количество") as Excel.Range;
            string sk;
            int a = firvalsm.Column + 1;
            for (int v = 0; v < KskSmete[AdresSmeta[num]].Count; v++)
            {
                for (int i = 0; i < ContainPapkaKS.Count; i++)
                {
                    if (AdresKS[i] != KskSmete[AdresSmeta[num]][v]) continue;
                    else
                    {
                        sk="Акт КС-2 №";
                        Excel.Worksheet workSheet;
                        workSheet = ContainPapkaKS[i].Sheets[1];
                        Excel.Range range;
                        range = workSheet.get_Range(sx1, sx2);
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
                        sk += findnum.Value.ToString();
                        string sgod = ParserExc.Finddate(regul.regexgod, finddata);
                        string smes = ParserExc.Finddate(regul.regexmes, finddata);
                        string havemes = ParserExc.Mespropis(smes);
                        havemes += sgod;
                        sk += " ";
                        sk += havemes;
                        sk += " ";
                        getVupoln = ParserExc.Getvupoln(workSheet, range, firstkey, otregexval);
                        ParserExc.Zapisinfile(getVupoln, firkeysm, firvalsm, Sheetcopy, rangecopy, sk, a);
                        ParserExc.FormatZapis(num,AdresSmeta,i,AdresKS,firkeysm, firvalsm, Sheetcopy, rangecopy);
                        Marshal.FinalReleaseComObject(range);
                        Marshal.FinalReleaseComObject(workSheet);
                        a += 1;
                    }
                }
            }
            Sheetcopy.Cells[firkeysm.Row, a]="Остаток";
            int colon = a - firvalsm.Column;
            if (colon > 1)
            {
                for (int j = firvalsm.Row + 2; j <= rangecopy.Rows.Count; j++)
                {
                    Excel.Range act = Sheetcopy.Cells[j, firvalsm.Column] as Excel.Range;
                    if (act != null && act.Value2 != null && act.Value2.ToString() != "" && !act.MergeCells)
                    {
                        Excel.Range activ = Sheetcopy.Cells[j, a] as Excel.Range;
                        switch (colon)
                        {
                            case 2:
                                activ.FormulaR1C1 = "=RC[-2]-RC[-1]";break;
                            case 3:
                                activ.FormulaR1C1 = "=RC[-3]-RC[-2]-RC[-1]"; break;
                            case 4:
                                activ.FormulaR1C1 = "=RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 5:
                                activ.FormulaR1C1 = "=RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 6:
                                activ.FormulaR1C1 = "=RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 7:
                                activ.FormulaR1C1 = "=RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 8:
                                activ.FormulaR1C1 = "=RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 9:
                                activ.FormulaR1C1 = "=RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 10:
                                activ.FormulaR1C1 = "=RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 11:
                                activ.FormulaR1C1 = "=RC[-11]-RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 12:
                                activ.FormulaR1C1 = "=RC[-12]-RC[-11]-RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 13:
                                activ.FormulaR1C1 = "=RC[-12]-RC[-12]-RC[-11]-RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            default: Console.WriteLine("Сводная таблица ведется до года, начните новую");break;
                        }
                    }
                }
            }
                    //вывод в консоль до окончательной отладки
                    //for (int i = 1; i <= rangecopy.Rows.Count; i++)
                    //{
                    //    Console.Write("\r\n");
                    //    for (int j = 1; j <= rangecopy.Columns.Count; j++)
                    //    {
                    //        Excel.Range forYach = Sheetcopy.Cells[i, j] as Excel.Range;
                    //        if (forYach != null && forYach.Value2 != null)
                    //            Console.Write(forYach.Value2.ToString() + "\t");
                    //    }
                    //}
             Zakrutie(AdresSmeta, KskSmete, ContainPapkaKS, AdresKS, CopySmet, num, Sheetcopy, rangecopy);
        }

    }
}