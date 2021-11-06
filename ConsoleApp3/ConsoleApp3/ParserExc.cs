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
        private static Excel.Workbook CopyExcel(string test, string t2, Excel.Application excelApp)
        {
            //Console.WriteLine("CopyExcel");
            Excel.Workbook excelBooksm = excelApp.Workbooks.Open(test);
            if (!File.Exists(t2))
            {
                excelBooksm.SaveCopyAs(t2);
            }
            excelBooksm.Close(false, Type.Missing, Type.Missing);
            Excel.Workbook excelBookco = excelApp.Workbooks.Open(t2);
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
        public static List<Excel.Workbook> GetListKS(string s, Excel.Application excelApp)
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
                    Excel.Workbook listKSone = excelApp.Workbooks.Open(oc);
                    listK.Add(listKSone);
                }
            }
            return listK;
        }
        //копирует последовательно все сметы в выбранную папку
        public static List<Excel.Workbook> MadeCopyExcbook(string userwheresave, string usersmetu, Excel.Application excelApp, List<Excel.Workbook> ContainPapkaSmeta, List<string> AdresSmeta)
        {
            //Console.WriteLine("MadeCopyExcbook");
            List<Excel.Workbook> excelBookcopy = new List<Excel.Workbook>();
            string test, t2;
            for (int u = 0; u < ContainPapkaSmeta.Count; u++)
            {
                t2 = userwheresave;
                test = AdresSmeta[u];
                string d = test.Remove(0, usersmetu.Length + 1);
                t2 += $" {d}";
                Excel.Workbook excelBook = CopyExcel(test, t2, excelApp);
                excelBookcopy.Add(excelBook);
            }
            return excelBookcopy;
        }
        //метод возвращает словарь, ключ - адрес сметы, 
        public static Dictionary<string, List<string>> GetContainSM(List<Excel.Workbook> ContainPapkaKS, List<string> AdresSmeta, List<string> AdresKS)
        {
           // Console.WriteLine("GetContainSM");
            Dictionary<string, List<string>> forSM = new Dictionary<string, List<string>>();
            RegexReg reg = new RegexReg();
            for (int u = 0; u < AdresSmeta.Count; u++)
            {
                string re = null;
                MatchCollection mathes = reg.namesmet.Matches(AdresSmeta[u]);
                if (mathes.Count > 0)
                {
                    foreach (Match math in mathes)
                        re = math.Value;
                }
                List<string> forsm = new List<string>();
                for (int c = 0; c < ContainPapkaKS.Count; c++)
                {
                    Excel.Worksheet workShet = ContainPapkaKS[c].Sheets[1];
                    Excel.Range rang = workShet.get_Range("A1", "I20");
                    if (rang.Find(re) == null) continue;
                    else forsm.Add(AdresKS[c]);
                    Marshal.FinalReleaseComObject(rang);
                    Marshal.FinalReleaseComObject(workShet);
                }
                forSM.Add(AdresSmeta[u], forsm);
            }
            return forSM;
        }
        //метод возвращает необходимую ячейку
        public static Excel.Range GetCell(Excel.Worksheet workSheet, Excel.Range range, Regex regul)
        {
            //Console.WriteLine("GetCell");
            MatchCollection mathes;
            Excel.Range result = null;
            for (int u = 1; u <= range.Rows.Count; u++)
            {
                for (int j = 1; j <= range.Columns.Count; j++)
                {
                    Excel.Range s1 = workSheet.Cells[u, j];
                    if (s1 != null && s1.Value != null)
                    {
                        mathes = regul.Matches(s1.Value.ToString());
                    }
                    else continue;
                    if (mathes.Count > 0)
                    { 
                        result = workSheet.Cells[u, j];
                        break;
                    }
                }
            }
            return result;
        }
        //метод возвращает строку с записанной датой составления акта для последующей обработки
        public static string Finddate(Regex a, Excel.Range finddata)
        {
            //Console.WriteLine(" Finddate");
            string sg = null;
            MatchCollection god = a.Matches(finddata.Value.ToString());
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
        public static string Mespropis(string smes)
        {
            //Console.WriteLine(" Mespropis");
            string sm = null;
            int mesac = 0;
            int f = 10;
            for (int j = 0; j < smes.Length; j++)
            {
                if (smes[j] >= '0' && smes[j] <= '9')
                {
                    mesac += (smes[j] - '0') * f;
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
        public static Dictionary<int, double> Getvupoln(Excel.Worksheet workSheet, Excel.Range range, Excel.Range firstkey, Excel.Range otregexval)
        {
            //Console.WriteLine("Getvupoln");
            Dictionary<int, double> W = new Dictionary<int, double>();
            int n1;
            double n2;
            for (int j = firstkey.Row + 2; j <= range.Rows.Count; j++)
            {
                Excel.Range forY2 = workSheet.Cells[j, firstkey.Column];
                Excel.Range forY3 = workSheet.Cells[j, otregexval.Column];
                if (forY2 != null && forY2.Value2 != null && forY3 != null && forY3.Value2 != null&& forY3.Value2.ToString() != ""&&forY2.Value2.ToString() != ""&& !forY2.MergeCells&& !forY3.MergeCells)
                {
                    n1 = Convert.ToInt32(forY2.Value2);
                    n2 = Convert.ToDouble(forY3.Value2);  
                    W.Add(n1, n2);
                }
            }
            return W;
        }
        

        //метод возвращает словарь где в ключ записан номер позиции по смете из Акта КС-2, а значение - нулл,
        //при записи в режиме эксперт в него будут суммироваться значения из Актов КС-2 в общей графе в смете
        public static Dictionary<int, T> Getkeysm<T>(Excel.Worksheet Sheetcopy, Excel.Range rangecopy, Excel.Range firkeysm)
        {
            //Console.WriteLine("Getkeysm<T>");
            Dictionary<int, T> s = new Dictionary<int, T>();
            int n3;
            T n4;
            for (int j = firkeysm.Row + 2; j <= rangecopy.Rows.Count; j++)
            {
                Excel.Range forY6 = Sheetcopy.Cells[j, firkeysm.Column] as Excel.Range;
               
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
      
        public static int GetColumzapis(Excel.Worksheet Sheetcopy, Excel.Range rangecopy, Excel.Range firkeysm)
        {
            //Console.WriteLine(" GetColumzapis");
            int d = -1;
            for (int j = firkeysm.Column; j <= rangecopy.Columns.Count; j++)
            {
                Excel.Range forY2 = Sheetcopy.Cells[firkeysm.Row, j];
                if (forY2 != null && forY2.Value2 != null) continue;
                else 
                {
                    Sheetcopy.Cells[firkeysm.Row, j] = "Примечание";
                    d = j;
                    break; 
                }
            }
            return d;  
        }
    
    }
}