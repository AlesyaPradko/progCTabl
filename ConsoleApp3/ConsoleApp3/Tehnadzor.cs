using System;
using System.Collections;
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
        protected override void ProcessSmeta(int num, RangeFile obl)
        {
            //Console.WriteLine("Workliketehnadzor");
            Excel.Worksheet Sheetcopy = copySmet[num].Sheets[1];
            Excel.Range rangecopy = Sheetcopy.get_Range(obl.FirstCell,obl.LastCell);
            Excel.Range firkeysm = rangecopy.Find("Номер");
            Excel.Range firvalsm = rangecopy.Find("Количество");
            int a = firvalsm.Column + 1;
            for (int v = 0; v < kskSmete[adresSmeta[num]].Count; v++)
            {
                for (int i = 0; i < containPapkaKS.Count; i++)
                {
                    if (adresKS[i] != kskSmete[adresSmeta[num]][v]) continue;
                    else
                    {
                        Excel.Worksheet workSheet = containPapkaKS[i].Sheets[1];
                        Excel.Range range = workSheet.get_Range(obl.FirstCell, obl.LastCell);
                        string sk=WorkKST(workSheet, range);
                        ZapisinfileT(firkeysm, Sheetcopy, rangecopy, sk, a);
                        FormatZapis(num, i, firkeysm, Sheetcopy, rangecopy);
                        Marshal.FinalReleaseComObject(range);
                        Marshal.FinalReleaseComObject(workSheet);
                        a += 1;
                    }
                }
            }
            ZapisFormulaT(a,Sheetcopy, rangecopy, firvalsm); 
            Zakrutie( num, Sheetcopy, rangecopy);
        }

        private string WorkKST(Excel.Worksheet workSheet, Excel.Range range)
        {
            string sk = "Акт КС-2 №";
            RegexReg regul = new RegexReg();
            Excel.Range firstkey = range.Find("по смете");
            Excel.Range otregexval = ParserExc.GetCell(workSheet, range, regul.regexval);
            Excel.Range findnum = range.Find("Номер документа");
            findnum= FindCellforNameKS(findnum, workSheet);
            Excel.Range finddata = range.Find("Дата составления");
            finddata= FindCellforNameKS(finddata, workSheet);
            sk += findnum.Value.ToString();
            string sgod = ParserExc.Finddate(regul.regexgod, finddata);
            string smes = ParserExc.Finddate(regul.regexmes, finddata);
            string havemes = ParserExc.Mespropis(smes);
            havemes += sgod;
            sk += $" {havemes} ";
            getVupoln = ParserExc.Getvupoln(workSheet, range, firstkey, otregexval);
            return sk;
        }
        //метод записывает в файл копии сметы объемы из Актов КС-2, каждый месяц в новый столбец,
        //вставка столбцов идет за столбцом объемы по смете  
        private void ZapisinfileT(Excel.Range firkeysm, Excel.Worksheet Sheetcopy, Excel.Range rangecopy, string sk, int a)
        {
            //Console.WriteLine(" Zapisinfile");
            ICollection keyColl = getVupoln.Keys;
            int ob1 = 0;
            for (int j = firkeysm.Row; j <= rangecopy.Rows.Count; j++)
            {
                Excel.Range forYs = Sheetcopy.Cells[j, a] as Excel.Range;
                forYs.Insert(XlInsertShiftDirection.xlShiftToRight);
                if (j > firkeysm.Row + 1)
                {
                    Excel.Range forY4 = Sheetcopy.Cells[j, firkeysm.Column] as Excel.Range;
                    if (forY4 != null && forY4.Value2 != null && forY4.Value2.ToString() != "" && !forY4.MergeCells)
                    {
                        ob1 = Convert.ToInt32(forY4.Value2);
                        foreach (int ob in keyColl)
                        {
                            if (ob1 == ob) Sheetcopy.Cells[j, a] = getVupoln[ob];
                        }
                    }
                }
            }
            Sheetcopy.Cells[firkeysm.Row, a] = sk;
        }
        private void ZapisFormulaT(int a,Excel.Worksheet Sheetcopy, Excel.Range rangecopy, Excel.Range firvalsm)
        {
            Sheetcopy.Cells[firvalsm.Row, a] = "Остаток";
            int colon = a - firvalsm.Column;
            if (colon > 1)
            {
                for (int j = firvalsm.Row + 2; j <= rangecopy.Rows.Count; j++)
                {
                    Excel.Range act = Sheetcopy.Cells[j, firvalsm.Column];
                    if (act != null && act.Value2 != null && act.Value2.ToString() != "" && !act.MergeCells)
                    {
                        Excel.Range activ = Sheetcopy.Cells[j, a];
                        switch (colon)
                        {
                            case 2:
                                activ.FormulaR1C1 = "=RC[-2]-RC[-1]"; break;
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
                                activ.FormulaR1C1 = "=RC[-13]-RC[-12]-RC[-11]-RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            default: Console.WriteLine("Сводная таблица ведется до года, начните новую"); break;
                        }
                        activ.EntireColumn.AutoFit();
                    }
                }
            }
        }
    }
}