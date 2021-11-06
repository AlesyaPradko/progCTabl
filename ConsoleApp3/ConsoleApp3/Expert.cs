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
        protected override void ProcessSmeta(int num, RangeFile obl)
        {
            //Console.WriteLine("Worklikeexpert");
             Excel.Worksheet Sheetcopy = copySmet[num].Sheets[1];
             Excel.Range rangecopy = Sheetcopy.get_Range(obl.FirstCell, obl.LastCell);
             Excel.Range firkeysm = rangecopy.Find("Номер");
             Excel.Range firvalsm = rangecopy.Find("Количество");
             int a = firvalsm.Column + 1;
             Excel.Range forYs = Sheetcopy.Cells[firkeysm.Row, a];
             forYs.Insert(XlInsertShiftDirection.xlShiftToRight);
             summSmeta = ParserExc.Getkeysm<double>(Sheetcopy, rangecopy, firkeysm);
             periodSmeta = ParserExc.Getkeysm<string>(Sheetcopy, rangecopy, firkeysm);
             Sheetcopy.Cells[firkeysm.Row, a] = "Выполнение по смете";
             int dprimst = ParserExc.GetColumzapis(Sheetcopy, rangecopy, firkeysm);
             if (dprimst == -1) { throw new ZapredelException("Вы задали слишком малую область по ширине таблицы, задайте большую"); return; }
             string[] sk = new string[containPapkaKS.Count];
             string[] valstr = periodSmeta.Values.ToArray();
             int numKS = 0;   
             for (int v = 0; v < kskSmete[adresSmeta[num]].Count; v++)
             {
                for (int i = 0; i < containPapkaKS.Count; i++)
                {
                    if (adresKS[i] != kskSmete[adresSmeta[num]][v]) continue;
                    else
                    {
                        Excel.Worksheet workSheet= containPapkaKS[i].Sheets[1];
                        Excel.Range range = workSheet.get_Range(obl.FirstCell, obl.LastCell);
                        WorKSE(workSheet, range,i, ref sk);
                        ZapisinfileE(firkeysm, Sheetcopy, rangecopy, sk, a, dprimst,i, valstr);
                        numKS = i;
                        Marshal.FinalReleaseComObject(range);
                        Marshal.FinalReleaseComObject(workSheet);
                    }
                }
             }
             ZapisFormulaE(a, Sheetcopy, rangecopy, firvalsm);
             FormatZapis(num, numKS, firkeysm, Sheetcopy, rangecopy);
             Zakrutie(num, Sheetcopy, rangecopy);
        }

        //вставка столбца "Осaток" с формулой разности
        private void ZapisFormulaE(int a, Excel.Worksheet Sheetcopy, Excel.Range rangecopy, Excel.Range firvalsm)
        {
            Excel.Range Ost = Sheetcopy.Cells[firvalsm.Row, a + 1] as Excel.Range;
            Ost.Insert(XlInsertShiftDirection.xlShiftToRight);
            Sheetcopy.Cells[firvalsm.Row, a + 1] = "Остаток";
            for (int j = firvalsm.Row + 1; j <= rangecopy.Rows.Count; j++)
            {
                Excel.Range vst = Sheetcopy.Cells[j, a + 1] as Excel.Range;
                vst.Insert(XlInsertShiftDirection.xlShiftToRight);
                if (j > firvalsm.Row + 1)
                {
                    Excel.Range act = Sheetcopy.Cells[j, firvalsm.Column] as Excel.Range;
                    if (act != null && act.Value2 != null && act.Value2.ToString() != "" && !act.MergeCells)
                    {
                        Excel.Range activ = Sheetcopy.Cells[j, a + 1] as Excel.Range;
                        activ.FormulaR1C1 = "=RC[-2]-RC[-1]";
                    }
                }
            }
        }

        private void WorKSE(Excel.Worksheet workSheet, Excel.Range range,int i,ref string[]sk)
        {
            sk[i] = "Акт КС-2 №";
            RegexReg regul = new RegexReg();
            Excel.Range firstkey = range.Find("по смете");
            Excel.Range otregexval = ParserExc.GetCell(workSheet, range, regul.regexval);
            Excel.Range findnum = range.Find("Номер документа");
            findnum = FindCellforNameKS(findnum, workSheet);
            Excel.Range finddata = range.Find("Дата составления") as Excel.Range;
            finddata = FindCellforNameKS(finddata, workSheet);
            sk[i] += findnum.Value.ToString();
            string sgod = ParserExc.Finddate(regul.regexgod, finddata);
            string smes = ParserExc.Finddate(regul.regexmes, finddata);
            string havemes = ParserExc.Mespropis(smes);
            havemes += sgod;
            sk[i] += $" {havemes}\n";
            getVupoln = ParserExc.Getvupoln(workSheet, range, firstkey, otregexval);
        }

        private void ZapisinfileE(Excel.Range firkeysm, Excel.Worksheet Sheetcopy, Excel.Range rangecopy, string[] sk, int a, int dprimst,int i, string[] valstr)
        {
            int[] keysm = summSmeta.Keys.ToArray();
            double[] valsm = summSmeta.Values.ToArray();
            ICollection keyColl = getVupoln.Keys;
            bool eqva;
            int ob1;
            for (int j = firkeysm.Row + 2; j <= rangecopy.Rows.Count; j++)
            {
                eqva = false;
                int ind = 0;
                Excel.Range forY4 = Sheetcopy.Cells[j, firkeysm.Column] as Excel.Range;
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
                            Sheetcopy.Cells[j, a] = summSmeta[ob];
                        }
                    }
                    if (eqva)
                    {
                        valstr[ind] += $"{sk[i]} ";
                        Sheetcopy.Cells[j, dprimst] = valstr[ind];
                    }
                }
            }
        }
    }
}