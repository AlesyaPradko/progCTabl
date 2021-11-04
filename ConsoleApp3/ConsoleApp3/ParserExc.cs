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
public enum XlInsertShiftDirection { xlShiftDown, xlShiftToRight };

namespace ConsoleApp3
{
    public static class ParserExc
    {
        //метод копирует одну книгу иксель по заданоому адресу
        public static Excel.Workbook CopyExcel(string s, string d, Excel.Application e)
        {
            //Console.WriteLine("CopyExcel");
            Excel.Workbook excelBooksm = e.Workbooks.Open(s);
            if (!File.Exists(d))
            {
                excelBooksm.SaveCopyAs(d);
            }
            excelBooksm.Close(false, Type.Missing, Type.Missing);
            Excel.Workbook excelBookco = e.Workbooks.Open(d);
            return excelBookco;
        }
        //метод возвращает лист строк с адресами смет и кс
        public static List<string> Getstring(string s)
        {
            //Console.WriteLine("Getstring");
            string[] gt = Directory.GetFiles(s);
            List<string> normal = new List<string>();
            foreach (string oc in gt)
            {
                if (oc.Contains("~$")) continue;
                if (!oc.Contains(".xlsx")) continue;
                else
                {
                    normal.Add(oc);
                }
            }
            return normal;
        }
        //запись файлов с Актами КС-2 в лист книг Excel
        public static List<Excel.Workbook> GetListKS(string s, Excel.Application e)
        {
            //Console.WriteLine("GetListKS");
            string[] gt = Directory.GetFiles(s);
            List<Excel.Workbook> listK = new List<Excel.Workbook>();
            foreach (string oc in gt)
            {
                if (oc.Contains("~$")) continue;
                if (!oc.Contains(".xlsx")) continue;
                else
                {
                    Excel.Workbook listKSone = e.Workbooks.Open(oc);
                    listK.Add(listKSone);
                }
            }
            return listK;
        }
        //копирует последовательно все сметы в выбранную папку
        public static List<Excel.Workbook> MadeCopyExcbook(string hran, string s, Excel.Application ex, List<Excel.Workbook> dd, List<string> adr)
        {
            //Console.WriteLine("MadeCopyExcbook");
            List<Excel.Workbook> excelBookcopy = new List<Excel.Workbook>();
            string test, t2;
            for (int u = 0; u < dd.Count; u++)
            {
                t2 = hran;
                test = adr[u];
                string d = test.Remove(0, s.Length + 1);
                t2 += " ";
                t2 += d;
                Excel.Workbook excelBook = ParserExc.CopyExcel(test, t2, ex);
                excelBookcopy.Add(excelBook);
            }
            return excelBookcopy;
        }
        //метод возвращает словарь, ключ - адрес сметы, 
        public static Dictionary<string, List<string>> GetContainSM(List<Excel.Workbook> bb, List<string> aa, List<string> cc)
        {
           // Console.WriteLine("GetContainSM");
            Dictionary<string, List<string>> forSM = new Dictionary<string, List<string>>();
            RegexReg reg = new RegexReg();
            for (int u = 0; u < aa.Count; u++)
            {
                string re = null;
                MatchCollection mathes = reg.namesmet.Matches(aa[u]);
                if (mathes.Count > 0)
                {
                    foreach (Match math in mathes)
                        re = math.Value;
                }
                List<string> forsm = new List<string>();
                for (int c = 0; c < bb.Count; c++)
                {
                    Excel.Worksheet workShet;
                    workShet = bb[c].Sheets[1];
                    Excel.Range rang;
                    rang = workShet.get_Range("A1", "I20");
                    if (rang.Find(re) == null) continue;
                    else forsm.Add(cc[c]);
                    Marshal.FinalReleaseComObject(rang);
                    Marshal.FinalReleaseComObject(workShet);
                }
                forSM.Add(aa[u], forsm);
            }
            return forSM;
        }
        //метод возвращает необходимую ячейку
        public static Excel.Range GetCell(Excel.Worksheet x, Excel.Range y, Regex regul)
        {
            //Console.WriteLine("GetCell");
            MatchCollection mathes;
            Excel.Range result = null;
            for (int u = 1; u <= y.Rows.Count; u++)
            {
                for (int j = 1; j <= y.Columns.Count; j++)
                {
                    Excel.Range s1 = x.Cells[u, j] as Excel.Range;
                    if (s1 != null && s1.Value != null)
                    {
                        mathes = regul.Matches(s1.Value.ToString());
                    }
                    else continue;
                    if (mathes.Count > 0) { result = x.Cells[u, j] as Excel.Range; break; }
                }
            }
            return result;
        }
        //метод возвращает строку с записанной датой составления акта для последующей обработки
        public static string Finddate(Regex a, Excel.Range dat)
        {
            //Console.WriteLine(" Finddate");
            string sg = null;
            MatchCollection god = a.Matches(dat.Value.ToString());
            if (god.Count > 0)
            {
                foreach (Match math in god)
                {
                    sg = math.Value;
                }
            }
            return sg;
        }
        //метод возвращает строку где месяц записан прописью
        public static string Mespropis(string s)
        {
            //Console.WriteLine(" Mespropis");
            string sm = null;
            int mesac = 0;
            int f = 10;
            for (int j = 0; j < s.Length; j++)
            {
                if (s[j] >= '0' && s[j] <= '9')
                {
                    mesac += (s[j] - '0') * f;
                    f /= 10;
                }
            }
            switch (mesac)
            {
                case 1: sm = "январь"; break;
                case 2: sm = "февраль"; break;
                case 3: sm = "март"; break;
                case 4: sm = "апрель"; break;
                case 5: sm = "май"; break;
                case 6: sm = "июнь"; break;
                case 7: sm = "июль"; break;
                case 8: sm = "август"; break;
                case 9: sm = "сентябрь"; break;
                case 10: sm = "октябрь"; break;
                case 11: sm = "ноябрь"; break;
                case 12: sm = "декабрь"; break;
            }
            return sm;
        }
        //метод возвращает словарь где в ключ записан номер позиции по смете из Акта КС-2, а значение - объем работ по этой позиции
        public static Dictionary<int, double> Getvupoln(Excel.Worksheet x, Excel.Range y, Excel.Range f, Excel.Range e)
        {
            //Console.WriteLine("Getvupoln");
            Dictionary<int, double> W = new Dictionary<int, double>();
            int n1;
            double n2;
            for (int j = f.Row + 2; j <= y.Rows.Count; j++)
            {
                Excel.Range forY2 = x.Cells[j, f.Column] as Excel.Range;
                Excel.Range forY3 = x.Cells[j, e.Column] as Excel.Range;
                if (forY2 != null && forY2.Value2 != null && forY3 != null && forY3.Value2 != null&& forY3.Value2.ToString() != ""&&forY2.Value2.ToString() != ""&& !forY2.MergeCells&& !forY3.MergeCells)
                {
                    n1 = Convert.ToInt32(forY2.Value2);
                    n2 = Convert.ToDouble(forY3.Value2);  
                    W.Add(n1, n2);
                }
            }
            return W;
        }
        //метод записывает в файл копии сметы объемы из Актов КС-2, каждый месяц в новый столбец,
        //вставка столбцов идет за столбцом объемы по смете  
        public static void Zapisinfile(Dictionary<int, double> V, Excel.Range fk, Excel.Range fv, Excel.Worksheet excS, Excel.Range r, string zx, int aa)
        {
            //Console.WriteLine(" Zapisinfile");
            ICollection keyColl = V.Keys;
            ICollection valColl = V.Values;
            int ob1=0;
            for (int j = fk.Row; j <= r.Rows.Count; j++)
            {
                Excel.Range forYs = excS.Cells[j, aa] as Excel.Range;
                forYs.Insert(XlInsertShiftDirection.xlShiftToRight);
                if (j > fk.Row+1)
                {
                    Excel.Range forY4 = excS.Cells[j, fk.Column] as Excel.Range;
                    if (forY4 != null && forY4.Value2 != null && forY4.Value2.ToString() != "" && !forY4.MergeCells)
                    {
                        ob1 = Convert.ToInt32(forY4.Value2);
                        foreach (int ob in keyColl)
                        {
                            if (ob1 == ob) excS.Cells[j, aa] = V[ob];
                        }
                    }
                }
            }
            excS.Cells[fk.Row, aa] = zx;
        }
        //метод задает формат записи в файл иксель
        public static void FormatZapis(int n,List<string>  ad,int i, List<string> adr,Excel.Range fk, Excel.Range fv, Excel.Worksheet excS, Excel.Range r)
        {
            //Console.WriteLine("FormatZapis");
            try
            {
                int shir = 0, vus = 0,test=0;
                for (int y = fk.Row; y <= r.Rows.Count; y++)
                {
                    Excel.Range fors = excS.Cells[y, fk.Column] as Excel.Range;
                    if (fors != null && fors.Value2 != null) vus++;
                    else { test++; if (fors != null && fors.Value2 != null) { vus += test; test = 0; } test = 0; }
                    if (test>5) break;
                    //Console.WriteLine(vus);
                }
                test = 0;
                for (int x = fk.Column; x <= r.Columns.Count; x++)
                {
                    Excel.Range fors = excS.Cells[fk.Row, x] as Excel.Range;
                    if (fors != null && fors.Value2 != null) shir++;
                    else { test++; if (fors != null && fors.Value2 != null) { shir += test; test = 0; } }
                    if (test > 5) break;
                    
                }
                Excel.Range flast = excS.Cells[fk.Row + vus+1, fk.Column + shir] as Excel.Range;
                if (flast.Column == r.Columns.Count || flast.Row == r.Rows.Count) { throw new ZapredelException($"Вы задали слишком малую область для {ad[n]} и для {adr[i]}"); return; }
                Excel.Range forIs = excS.get_Range(fk, flast);
                forIs.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                forIs.EntireColumn.Font.Size = 13;
                forIs.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
                forIs.EntireColumn.AutoFit();
            }
            catch (ZapredelException exc)
            { Console.WriteLine(exc.parName); }
        }

        //метод возвращает словарь где в ключ записан номер позиции по смете из Акта КС-2, а значение - нулл,
        //при записи в режиме эксперт в него будут суммироваться значения из Актов КС-2 в общей графе в смете
        public static Dictionary<int, T> Getkeysm<T>(Excel.Worksheet exc, Excel.Range r1, Excel.Range f1)
        {
            //Console.WriteLine("Getkeysm<T>");
            Dictionary<int, T> s = new Dictionary<int, T>();
            int n3;
            T n4;
            for (int j = f1.Row + 2; j <= r1.Rows.Count; j++)
            {
                Excel.Range forY6 = exc.Cells[j, f1.Column] as Excel.Range;
               
                if (forY6 != null && forY6.Value2 != null && forY6.Value2.ToString() != "")
                {
                    n3 = (int)(forY6.Value2);
                    n4 = default(T);
                    s.Add(n3, n4);
                }
            }
            return s;
        } 
        //получение столбца где будет записан столбец примечание для записи в него из каких актов КС-2 взяты объемы
      
        public static int GetColumzapis(Excel.Worksheet x, Excel.Range y, Excel.Range f)
        {
            //Console.WriteLine(" GetColumzapis");
            int d = -1;
            for (int j = f.Column; j <= y.Columns.Count; j++)
            {
                Excel.Range forY2 = x.Cells[f.Row, j] as Excel.Range;
                if (forY2 != null && forY2.Value2 != null) continue;
                else { x.Cells[f.Row, j] = "Примечание"; d = j; break; }
            }
            return d;  
        }
    
    }
}