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
        //проверка, установлен ли Excel на компьютере и зоздание элемента класса для работы с файлами
        public static Excel.Application Proverka()
        {
            Excel.Application excelA = new Excel.Application();
            if (excelA == null)
            {
                Console.WriteLine("Excel is not installed!!");
                return null;
            }
            else return excelA;
        }
        //копирование сметы в выбранную папку, если она уже создана, то работа ведется с существующим файлом
        public static Excel.Workbook CopyExcel(string s, string d)
        {

            Excel.Workbook excelBooksm = Proverka().Workbooks.Open(s);
            if (!File.Exists(d))
            {
                excelBooksm.SaveCopyAs(d);
            }
            excelBooksm.Close(false, Type.Missing, Type.Missing);
            Excel.Workbook excelBookco = Proverka().Workbooks.Open(d);
            return excelBookco;
        }
        //запись файлов с Актами КС-2 в лист книг Excel
        public static List<Excel.Workbook> GetListKS(string s)
        {
            string[] gt = Directory.GetFiles(s);
            List<Excel.Workbook> listK = new List<Excel.Workbook>();
            foreach (string oc in gt)
            {
                if (oc.Contains("~$")) continue;
                else
                {
                    Excel.Workbook listKSone = Proverka().Workbooks.Open(oc);
                    listK.Add(listKSone);
                }
            }
            return listK;
        }
        //метод возвращает необходимую ячейку
        public static Excel.Range GetCell(Excel.Worksheet x, Excel.Range y, Regex regul)
        {
            MatchCollection mathes;
            Excel.Range result = null;
            for (int u = 1; u <= y.Rows.Count; u++)
            {
                for (int j = 1; j <= y.Columns.Count; j++)
                {
                    Excel.Range s1 = x.Cells[u, j] as Excel.Range;
                    if (s1 == null && s1.Value == null)
                    {
                        continue;
                    }

                    mathes = regul.Matches(s1.Value.ToString());
                    if (mathes.Count > 0) 
                    { 
                        result = x.Cells[u, j] as Excel.Range;
                        break; 
                    }
                }
            }
            return result;
        }
        //метод возвращает строку с записанной датой составления акта для последующей обработки
        public static string Finddate(Regex a, Excel.Range dat)
        {
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
            Dictionary<int, double> W = new Dictionary<int, double>();

            int n1;
            double n2;
            for (int j = f.Row + 1; j <= y.Rows.Count; j++)
            {
                Excel.Range forY2 = x.Cells[j, f.Column] as Excel.Range;
                Excel.Range forY3 = x.Cells[j, e.Column] as Excel.Range;
                if (forY2 != null && forY2.Value2 != null && forY3 != null && forY3.Value2 != null)
                {
                    n1 = (int)forY2.Value2;
                    n2 = forY3.Value2;
                    W.Add(n1, n2);
                }
            }
            return W;
        }
        //метод записывает в файл копии сметы объемы из Актов КС-2, каждый месяц в новый столбец,
        //вставка столбцов идет за столбцом объемы по смете  
        public static void Zapisinfile(Dictionary<int, double> V, Excel.Range fk, Excel.Range fv, Excel.Worksheet excS, Excel.Range r, string zx)
        {
            ICollection keyColl = V.Keys;
            ICollection valColl = V.Values;
            int a = fv.Column + 1;
            for (int j = fk.Row; j <= r.Rows.Count; j++)
            {
                Excel.Range forYs = excS.Cells[j, a] as Excel.Range;
                forYs.Insert(XlInsertShiftDirection.xlShiftToRight);
                if (j > fk.Row)
                {
                    Excel.Range forY4 = excS.Cells[j, fk.Column] as Excel.Range;
                    foreach (int ob in keyColl) 
                    {
                        if (forY4.Value2 == ob) excS.Cells[j, a] = V[ob];
                    }
                }
            } 
            excS.Cells[fk.Row, a] = zx;
            a += 1;
        }
        //обработка сметы и Актов КС-2 в режиме технадзор
        public static void Workliketehnadzor(List<Excel.Workbook> d, Excel.Worksheet excS, Excel.Range r, Excel.Range fk, Excel.Range fv)
        {
            string sk;
            for (int i = 0; i < d.Count; i++)
            {
                Excel.Worksheet workSheet;
                workSheet = d[i].Sheets[1];
                Excel.Range range;
                range = workSheet.get_Range("A1", "L30");
                RegexReg regul = new RegexReg();
                Excel.Range firstkey = range.Find("по смете") as Excel.Range;
                Excel.Range otregexval = GetCell(workSheet, range, regul.regexval);
                Excel.Range otregexKS = GetCell(workSheet, range, regul.regexKS);
                Excel.Range otregexdat = GetCell(workSheet, range, regul.regexdat);
                sk = otregexKS.Value;
                string sgod = Finddate(regul.regexgod, otregexdat);
                string smes = Finddate(regul.regexmes, otregexdat);
                string havemes = Mespropis(smes);
                havemes += sgod;
                sk += " ";
                sk += havemes;
                sk += " ";
                Dictionary<int, double> Vupolnenie = Getvupoln(workSheet, range, firstkey, otregexval);
                Zapisinfile(Vupolnenie, fk, fv, excS, r, sk);
                Marshal.FinalReleaseComObject(range);
                Marshal.FinalReleaseComObject(workSheet);
            }

        }
        //метод возвращает словарь где в ключ записан номер позиции по смете из Акта КС-2, а значение - нулл,
        //при записи в режиме эксперт в него будут суммироваться значения из Актов КС-2 в общей графе в смете
        public static Dictionary<int, T> Getkeysm<T>(Excel.Worksheet exc, Excel.Range r1, Excel.Range f1)
        {
            Dictionary<int, T> s = new Dictionary<int, T>();
            int n3;
            T n4;
            for (int j = f1.Row + 1; j <= r1.Rows.Count; j++)
            {
                Excel.Range forY6 = exc.Cells[j, f1.Column] as Excel.Range;
                if (forY6 != null && forY6.Value2 != null)
                {
                    n3 = (int)forY6.Value2;
                    n4 = default(T);
                    s.Add(n3, n4);
                }
            }
            return s;
        }
        //получение столбца где будет записан столбец примечание для записи в него из каких актов КС-2 взяты объемы
        public static int GetColumzapis(Excel.Worksheet x, Excel.Range y, Excel.Range f)
        {
            int d = 0;
            for (int j = f.Column; j <= y.Columns.Count; j++)
            {
                Excel.Range forY2 = x.Cells[f.Row, j] as Excel.Range;
                if (forY2 != null && forY2.Value2 != null) continue;
                else { x.Cells[f.Row, j] = "Примечание"; d = j; break; }
            }
            return d;
        }
        //обработка сметы и запись в нее данных из Актов КС-2 в режиме эксперт
        public static void Worklikeexpert(List<Excel.Workbook> d, Excel.Worksheet excS, Excel.Range r, Excel.Range fk, Excel.Range fv)
        {

            int a = fv.Column + 1;
            Excel.Range forYs = excS.Cells[fk.Row, a] as Excel.Range;
            forYs.Insert(XlInsertShiftDirection.xlShiftToRight);
            Dictionary<int, double> sum = Getkeysm<double>(excS, r, fk);
            Dictionary<int, string> stroc = Getkeysm<string>(excS, r, fk);
            excS.Cells[fk.Row, a] = "Выполнение по смете";
            int dprimst = GetColumzapis(excS, r, fk);
            int[] keysm = sum.Keys.ToArray();
            double[] valsm = sum.Values.ToArray();
            string[] valstr = stroc.Values.ToArray();
            string[] sk = new string[d.Count];
            for (int i = 0; i < d.Count; i++)
            {
                Excel.Worksheet workSheet;
                workSheet = d[i].Sheets[1];
                Excel.Range range;
                range = workSheet.get_Range("A1", "L30");
                RegexReg regul = new RegexReg();
                Excel.Range firstkey = range.Find("по смете") as Excel.Range;
                Excel.Range otregexval = GetCell(workSheet, range, regul.regexval);
                Excel.Range otregexKS = GetCell(workSheet, range, regul.regexKS);
                Excel.Range otregexdat = GetCell(workSheet, range, regul.regexdat);
                sk[i] = otregexKS.Value;
                string sgod = Finddate(regul.regexgod, otregexdat);
                string smes = Finddate(regul.regexmes, otregexdat);
                string havemes = Mespropis(smes);
                havemes += sgod;
                sk[i] += " ";
                sk[i] += havemes;
                sk[i] += " ";
                Dictionary<int, double> Withdoub = Getvupoln(workSheet, range, firstkey, otregexval);
                int[]  keysu = Withdoub.Keys.ToArray();
                double[] valu = Withdoub.Values.ToArray();
                bool eqva;
                for (int j = fk.Row + 1; j <= r.Rows.Count; j++)
                {
                    eqva = false;
                    int ind = 0;
                    for (int u = 0; u < keysu.Length; u++)
                    {
                        Excel.Range forY4 = excS.Cells[j, fk.Column] as Excel.Range;
                        if (forY4.Value2 == keysu[u])
                        {
                            eqva = true;
                            ind = Array.IndexOf(keysm, keysu[u]);
                            valsm[ind] += valu[u];
                            sum[keysu[u]] = valsm[ind];
                            excS.Cells[j, a] = sum[keysu[u]];
                        }
                    }
                    if (eqva)
                    {
                        valstr[ind] += sk[i];
                        valstr[ind] += " ";
                        excS.Cells[j, dprimst] = valstr[ind];
                    }
                }
                Marshal.FinalReleaseComObject(range);
                Marshal.FinalReleaseComObject(workSheet);
            }  
        }
    }
}