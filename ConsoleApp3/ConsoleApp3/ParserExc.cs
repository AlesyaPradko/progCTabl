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
        //метод копирует одну книгу иксель по заданому адресу
        public static Excel.Workbook CopyExcelSmetaOne(string adresoneSmeta, string testuserwheresave, Excel.Application excelApp)
        {
            //Console.WriteLine("CopyExcelSmetaOne");
            Excel.Workbook excelBooksm = excelApp.Workbooks.Open(adresoneSmeta);
            if (!File.Exists(testuserwheresave))
            {
                excelBooksm.SaveCopyAs(testuserwheresave);
            }
            excelBooksm.Close(false, Type.Missing, Type.Missing);
            Excel.Workbook excelBookcopySmet = excelApp.Workbooks.Open(testuserwheresave);
            return excelBookcopySmet;
        }
        //метод возвращает лист строк с адресами смет и кс
        public static List<string> GetstringAdresa(string useradress)
        {
            //Console.WriteLine("GetstringAdresa");
            string[] allAdresContainsFolder = Directory.GetFiles(useradress);
            List<string> adresDocuments = new List<string>();
            foreach (string adresDoc in allAdresContainsFolder)
            {
                if (adresDoc.Contains("~$")) continue;
                if (!adresDoc.Contains(".xlsx")) continue;
                else
                {
                    adresDocuments.Add(adresDoc);
                }
            }
            return adresDocuments;
        }
        //запись файлов с Актами КС-2 и сметами в лист книг Excel
        public static List<Excel.Workbook> GetBookAllAktKSandSmeta(string userKS, Excel.Application excelApp)
        {
            //Console.WriteLine("GetBookAllAktKS");
            string[] nameAdresAktKS = Directory.GetFiles(userKS);
            List<Excel.Workbook> containPapkaKS = new List<Excel.Workbook>();
            foreach (string onenameKS in nameAdresAktKS)
            {
                if (onenameKS.Contains("~$")) continue;
                if (!onenameKS.Contains(".xlsx")) continue;
                else
                {
                    Excel.Workbook bookAktKSone = excelApp.Workbooks.Open(onenameKS);
                    containPapkaKS.Add(bookAktKSone);
                }
            }
            return containPapkaKS;
        }

        //метод возвращает словарь, ключ - адрес сметы, значение - адреса актов КС, относящихся к смете
        public static Dictionary<string, List<string>> GetContainAktKSinOneSmeta(List<Excel.Workbook> ContainPapkaKS, List<string> AdresSmeta, List<string> AdresAktKS)
        {
            // Console.WriteLine("GetContainAktKSinOneSmeta");
            Dictionary<string, List<string>> aktAllKSforOneSmeta = new Dictionary<string, List<string>>();
            RegexReg reg = new RegexReg();
            for (int u = 0; u < AdresSmeta.Count; u++)
            {
                string numerSmetastring = null;
                MatchCollection mathesNumerSmeta = reg.nameSsmeta.Matches(AdresSmeta[u]);
                Console.WriteLine(" smeta " + AdresSmeta[u]);
                if (mathesNumerSmeta.Count > 0)
                {
                    foreach (Match numerSmeta in mathesNumerSmeta)
                        numerSmetastring = numerSmeta.Value;
                    Console.WriteLine(numerSmetastring + " smeta " + AdresSmeta[u]);
                }
                List<string> aktKSforSmeta = new List<string>();
                for (int c = 0; c < ContainPapkaKS.Count; c++)
                {
                    Excel.Worksheet workShetAktKS = ContainPapkaKS[c].Sheets[1];
                    Excel.Range rangAktKS = workShetAktKS.get_Range("A1", "I40");
                    if (rangAktKS.Find(numerSmetastring) == null) continue;
                    else
                    {
                        aktKSforSmeta.Add(AdresAktKS[c]);
                        Console.WriteLine("AdresKS[c] " + numerSmetastring + " " + AdresAktKS[c]);
                    }
                    Marshal.FinalReleaseComObject(rangAktKS);
                    Marshal.FinalReleaseComObject(workShetAktKS);
                }
                aktAllKSforOneSmeta.Add(AdresSmeta[u], aktKSforSmeta);
            }
            return aktAllKSforOneSmeta;
        }
        //метод возвращает необходимую ячейку
        public static Excel.Range FindCellofRegul(Excel.Worksheet workSheetAktKS, Excel.Range rangeAktKS, Regex regul)
        {
            //Console.WriteLine("FindCellofRegul");
            MatchCollection mathesfindCell;
            Excel.Range findCellsColumn = null;
            for (int u = 1; u <= rangeAktKS.Rows.Count; u++)
            {
                for (int j = 1; j <= rangeAktKS.Columns.Count; j++)
                {
                    Excel.Range nextCellinAktKS = workSheetAktKS.Cells[u, j];
                    if (nextCellinAktKS != null && nextCellinAktKS.Value != null && nextCellinAktKS.ToString() != "")
                    {
                        mathesfindCell = regul.Matches(nextCellinAktKS.Value.ToString());
                    }
                    else continue;
                    if (mathesfindCell.Count > 0)
                    {
                        findCellsColumn = workSheetAktKS.Cells[u, j];
                        break;
                    }
                }
            }
            return findCellsColumn;
        }
        //метод возвращает строку с записанной датой составления акта для последующей обработки
        public static string FinddateAktKS(Regex monthorYear, Excel.Range finddata)
        {
            //Console.WriteLine("FinddateAktKS");
            string dateMonthorYear = null;
            MatchCollection yearmonth = monthorYear.Matches(finddata.Value.ToString());
            if (yearmonth.Count > 0)
            {
                foreach (Match onedate in yearmonth)
                {
                    dateMonthorYear = onedate.Value;
                }
            }
            return dateMonthorYear;
        }
        //метод возвращает строку где месяц записан прописью, входной параметр число
        public static string MonthpropisInt(int montnach)
        {
            //Console.WriteLine("MonthpropisInt");
            string monthpropis = null;
            switch (montnach)
            {
                case 1: monthpropis = "январь"; break;
                case 2: monthpropis = "февраль"; break;
                case 3: monthpropis = "март"; break;
                case 4: monthpropis = "апрель"; break;
                case 5: monthpropis = "май"; break;
                case 6: monthpropis = "июнь"; break;
                case 7: monthpropis = "июль"; break;
                case 8: monthpropis = "август"; break;
                case 9: monthpropis = "сентябрь"; break;
                case 10: monthpropis = "октябрь"; break;
                case 11: monthpropis = "ноябрь"; break;
                case 12: monthpropis = "декабрь"; break;
            }
            return monthpropis;
        }

        //метод возвращает строку где месяц записан прописью, входной параметр строкак как число
        public static string Monthpropis(string monthAktKS)
        {
            //Console.WriteLine("Monthpropis");
            string monthAktKSpropis;
            int month = 0;
            int twodigit = 10;
            for (int j = 0; j < monthAktKS.Length; j++)
            {
                if (monthAktKS[j] >= '0' && monthAktKS[j] <= '9')
                {
                    month += (monthAktKS[j] - '0') * twodigit;
                    twodigit /= 10;
                }
            }
            monthAktKSpropis = MonthpropisInt(month);
            return monthAktKSpropis;
        }

        //метод возвращает словарь где в ключ записан номер позиции по смете из Акта КС-2, а значение - объем работ по этой позиции
        public static Dictionary<int, double> GetScopeWorkAktKSone(Excel.Worksheet workSheetAktKS, Excel.Range rangeAktKS, Excel.Range keyNumPozpoSmeteinAktKS, Excel.Range keyscopeWorkinAktKS, string adresKs)
        {
            //Console.WriteLine("GetScopeWorkAktKSone");
            Dictionary<int, double> totalScopeWorkAktKSone = new Dictionary<int, double>();
            int valueNumPoz;
            double valueScopeWork;
            for (int j = keyNumPozpoSmeteinAktKS.Row + 2; j < rangeAktKS.Rows.Count + keyNumPozpoSmeteinAktKS.Row; j++)
            {
                Excel.Range cellsNumPozColumnTabl = workSheetAktKS.Cells[j, keyNumPozpoSmeteinAktKS.Column];
                Excel.Range cellsScopeColumnTabl = workSheetAktKS.Cells[j, keyscopeWorkinAktKS.Column];
                if (cellsNumPozColumnTabl != null && cellsNumPozColumnTabl.Value2 != null && cellsScopeColumnTabl != null && cellsScopeColumnTabl.Value2 != null && cellsScopeColumnTabl.Value2.ToString() != "" && cellsNumPozColumnTabl.Value2.ToString() != "" && !cellsNumPozColumnTabl.MergeCells && !cellsScopeColumnTabl.MergeCells)
                {
                    try
                    {
                        valueNumPoz = Convert.ToInt32(cellsNumPozColumnTabl.Value2);
                        valueScopeWork = Convert.ToDouble(cellsScopeColumnTabl.Value2);
                        totalScopeWorkAktKSone.Add(valueNumPoz, valueScopeWork);
                    }
                    catch (FormatException ex)
                    {
                        Console.WriteLine($"{ex.Message} Вы ввели неверный формат для {adresKs} в строке {cellsNumPozColumnTabl.Row} в столбце {cellsNumPozColumnTabl.Column}(не должно быть [.,букв], только целые числа или же в столбце {cellsScopeColumnTabl.Column} не должно быть [.букв], только дробные числа с [,] или целые числа)");
                    }
                    catch (ArgumentException ex)
                    {
                        Console.WriteLine($"{ex.Message} Проверьте чтобы в {adresKs} не повторялись значения позиций по смете в строке {cellsNumPozColumnTabl.Row}");
                    }
                }
            }
            return totalScopeWorkAktKSone;
        }


        //метод возвращает словарь где в ключ записан номер позиции по смете из сметы, а значение - нулл,
        //при записи в режиме эксперт в него будут суммироваться значения из Актов КС-2 в общей графе в смете
        // Зачем делать словарь и в значение класть нулы? Если нужен список номеров можно использовать обычный лист.
        public static Dictionary<int, T> GetkeySmetaForZapis<T>(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, string AdresSmeta)
        {
            //Console.WriteLine("GetkeySmetaForZapis<T>");
            Dictionary<int, T> resultwithNumPoz = new Dictionary<int, T>();
            int numPozSmeta;
            T zerovalue;
            for (int j = rangeSmetaOne.Row + 4; j < rangeSmetaOne.Rows.Count + rangeSmetaOne.Row; j++)
            {
                Excel.Range cellsFirstColumnTabl = SheetcopySmetaOne.Cells[j, rangeSmetaOne.Column];
                if (cellsFirstColumnTabl != null && cellsFirstColumnTabl.Value2 != null && cellsFirstColumnTabl.Value2.ToString() != "" && !cellsFirstColumnTabl.MergeCells)
                {
                    try
                    {
                        numPozSmeta = Convert.ToInt32(cellsFirstColumnTabl.Value2);
                        zerovalue = default(T);
                        resultwithNumPoz.Add(numPozSmeta, zerovalue);
                    }
                    catch (ArgumentException ex)
                    {
                        Console.WriteLine($"{ex.Message} Проверьте чтобы в {AdresSmeta} не повторялись значения позиций по смете в строке {cellsFirstColumnTabl.Row}");
                    }
                    catch (FormatException ex)
                    {
                        Console.WriteLine($"{ex.Message} Вы ввели неверный формат для {AdresSmeta} в строке {cellsFirstColumnTabl.Row} в столбце {cellsFirstColumnTabl.Column}(не должно быть [., букв], только целые числа)");
                    }
                }
            }
            return resultwithNumPoz;
        }

        //получение столбца где будет записан столбец примечание для записи в него из каких актов КС-2 взяты объемы
        public static int GetColumforZapisNote(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne)
        {
            //Console.WriteLine(" GetColumforZapis");
            int numLastColumnCellNote = -1;
            for (int j = rangeSmetaOne.Column; j <= rangeSmetaOne.Columns.Count; j++)
            {
                Excel.Range cellsFirstRowTabl = SheetcopySmetaOne.Cells[rangeSmetaOne.Row, j];
                if (cellsFirstRowTabl != null && cellsFirstRowTabl.Value2 != null || cellsFirstRowTabl.MergeCells) continue;
                else
                {
                    Excel.Range topCellmergeCellContentNote = SheetcopySmetaOne.Cells[rangeSmetaOne.Row, j];
                    Excel.Range bottomCellmergeCellContentNote = SheetcopySmetaOne.Cells[rangeSmetaOne.Row + 2, j];
                    Excel.Range mergeCellContentNote = SheetcopySmetaOne.get_Range(topCellmergeCellContentNote, bottomCellmergeCellContentNote);
                    mergeCellContentNote.Merge();
                    mergeCellContentNote.Value = "Примечание";
                    SheetcopySmetaOne.Cells[rangeSmetaOne.Row + 3, j] = j - rangeSmetaOne.Column + 2;
                    numLastColumnCellNote = j;
                    break;
                }
            }
            return numLastColumnCellNote;
        }
        //метод удаляет ненужные столбцы и строки для формирования ведомости выполненных объемов работ
        public static void DeleteColumnAndRow(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, Excel.Range keyCellNomerpozSmeta, string AdresSm, ref int lastRowCellsafterDelete)
        {
            //Console.WriteLine(" DeleteColumnandRow");
            try
            {
                List<int> deleteExcessCells = new List<int>();
                for (int u = keyCellNomerpozSmeta.Row + 6; u <= rangeSmetaOne.Rows.Count; u++)
                {
                    Excel.Range cellsFirstColumnTabl = SheetcopySmetaOne.Cells[u, keyCellNomerpozSmeta.Column];
                    if (cellsFirstColumnTabl.MergeCells && !cellsFirstColumnTabl.Value.ToString().Contains("Раздел"))
                    {
                        deleteExcessCells.Add(cellsFirstColumnTabl.Row);
                    }
                }
                Excel.Range lastcellsFirstColumnTabl = SheetcopySmetaOne.Cells[rangeSmetaOne.Rows.Count, keyCellNomerpozSmeta.Column];
                if (lastcellsFirstColumnTabl != null && lastcellsFirstColumnTabl.Value != null && lastcellsFirstColumnTabl.Value2.ToString() != "")
                {
                    throw new ZapredelException($"Вы задали слишком малую высоту для {AdresSm}");
                }
                deleteExcessCells.Reverse();
                lastRowCellsafterDelete = deleteExcessCells[0] - deleteExcessCells.Count; //тестить
                Console.WriteLine(AdresSm);
                Console.WriteLine("Work Task.CurrentId " + Task.CurrentId);
                Console.WriteLine("na udalenie " + deleteExcessCells.Count + "last Row" + lastRowCellsafterDelete);
                for (int u = rangeSmetaOne.Rows.Count; u > keyCellNomerpozSmeta.Row + 6; u--)
                {
                    Excel.Range cellsFirstColumnTabl = SheetcopySmetaOne.Cells[u, keyCellNomerpozSmeta.Column];
                    for (int v = 0; v < deleteExcessCells.Count; v++)
                    {
                        if (cellsFirstColumnTabl.Row == deleteExcessCells[v])
                        {
                            Excel.Range lastColumnOnDelet = SheetcopySmetaOne.Cells[cellsFirstColumnTabl.Row, rangeSmetaOne.Columns.Count];
                            Excel.Range rowOnDelet = SheetcopySmetaOne.get_Range(cellsFirstColumnTabl, lastColumnOnDelet);
                            //Console.WriteLine(AdresSm);
                            //Console.WriteLine (" Task.CurrentId " + Task.CurrentId + " Cell to delete " + cellsFirstColumnTabl.Row);
                            rowOnDelet.Delete();
                            break;
                        }
                    }
                }
                Regex rex = new Regex(@"(С|с)тоимость");
                MatchCollection mathesStoim = null;
                Excel.Range lastCellOnRangeForDelet = SheetcopySmetaOne.Cells[rangeSmetaOne.Rows.Count, rangeSmetaOne.Columns.Count];
                Excel.Range firstCellOnRangeForDelet = null;
                for (int u = keyCellNomerpozSmeta.Column; u <= rangeSmetaOne.Columns.Count; u++)
                {
                    Excel.Range cellsFirstRowTabl = SheetcopySmetaOne.Cells[keyCellNomerpozSmeta.Row, u];
                    if (cellsFirstRowTabl != null && cellsFirstRowTabl.Value != null)
                    {
                        mathesStoim = rex.Matches(cellsFirstRowTabl.Value.ToString());
                    }
                    if (mathesStoim.Count > 0)
                    {
                        firstCellOnRangeForDelet = SheetcopySmetaOne.Cells[keyCellNomerpozSmeta.Row, u];
                        break;
                    }
                }
                if (firstCellOnRangeForDelet != null)
                {
                    Excel.Range rangeOnDelet = SheetcopySmetaOne.get_Range(firstCellOnRangeForDelet, lastCellOnRangeForDelet);
                    rangeOnDelet.Delete();
                }
                else
                {
                    Console.WriteLine($"Проверьте верно ли в {AdresSm} записано выражение (С|с)тоимость");
                    return;
                }
            }
            catch (ZapredelException exc)
            {
                Console.WriteLine(exc.parName);
            }
        }
        //методы для Графика
        //отсекает адрес и формат от имени сметы
        public static void GetNameSmeta(string adresSmeta, out string nameFailSmeta)
        {
            //Console.WriteLine(" GetNameSmeta");
            nameFailSmeta = null;
            int numberChertuinString = 0;
            for (int i = 0; i < adresSmeta.Length; i++)
            {
                if (adresSmeta[i] == '\\') numberChertuinString = i;
            }
            for (int i = numberChertuinString + 1; i < adresSmeta.Length - 5; i++)
            {
                nameFailSmeta += adresSmeta[i];
            }
        }
        //возвращает общую трудоемкость по смете в виде цифры, входной параметр - строковый
        public static double CifrafromStringCell(string trudozatrataString)
        {
            //Console.WriteLine(" CifrafromStringCell");
            string summaString = null;
            for (int i = 0; i < trudozatrataString.Length; i++)
            {
                if ((trudozatrataString[i] >= '0' && trudozatrataString[i] <= '9') || trudozatrataString[i] == ',')
                {
                    summaString += trudozatrataString[i];
                }
            }
            double trudozatratobch = Convert.ToDouble(summaString);
            return trudozatratobch;
        }
        //возвращает словарь, где ключ ячейка "Итого по разделу", значение - трудоемкость во разделу
        public static Dictionary<Excel.Range, double> FindpoRazdely(Excel.Worksheet workSheetoneSmeta, Excel.Range rangeoneSmeta, Excel.Range keyCellColumnTopTrudozatrat, Excel.Range keyCellNumberPozSmeta)
        {
            //Console.WriteLine("  FindpoRazdely");
            Dictionary<Excel.Range, double> porazdely = new Dictionary<Excel.Range, double>();
            RegexReg p = new RegexReg();
            MatchCollection mathes1;
            for (int j = 1; j <= rangeoneSmeta.Rows.Count; j++)
            {
                Excel.Range namerazd = workSheetoneSmeta.Cells[j, keyCellNumberPozSmeta.Column];
                if (namerazd != null && namerazd.Value2 != null && namerazd.MergeCells && namerazd.Value2.ToString() != "")
                {
                    string s = namerazd.Value.ToString();
                    mathes1 = p.cellItogoPorazdely.Matches(s);
                }
                else continue;
                if (mathes1.Count > 0)
                {
                    Excel.Range c1 = workSheetoneSmeta.Cells[namerazd.Row, keyCellColumnTopTrudozatrat.Column];
                    if (c1.Value2 != null && c1 != null && c1.Value.ToString() != "")
                    {
                        double trudrazd = Convert.ToDouble(c1.Value2);
                        porazdely.Add(namerazd, trudrazd);
                    }
                }
            }
            return porazdely;
        }
        //возвращает лист из ячеек "Раздел такой-то"
        public static List<Excel.Range> FindRazdel(Excel.Worksheet workSheetoneSmeta, Excel.Range rangeoneSmeta, Excel.Range keyCellNumberPozSmeta)
        {
            //Console.WriteLine("  FindRazdel");
            List<Excel.Range> cellsAllRazdel = new List<Excel.Range>();
            RegexReg cellRazdelFind = new RegexReg();
            MatchCollection mathesRazdel;
            for (int j = 1; j <= rangeoneSmeta.Rows.Count; j++)
            {
                Excel.Range nameRazdel = workSheetoneSmeta.Cells[j, keyCellNumberPozSmeta.Column];
                if (nameRazdel != null && nameRazdel.Value2 != null && nameRazdel.MergeCells && nameRazdel.Value2.ToString() != "")
                {
                    string stringNameRazdel = nameRazdel.Value.ToString();
                    mathesRazdel = cellRazdelFind.cellOfRazdel.Matches(stringNameRazdel);
                }
                else continue;
                if (mathesRazdel.Count > 0)
                {
                    cellsAllRazdel.Add(nameRazdel);
                }
            }
            return cellsAllRazdel;
        }
        //функция меняет по ссылке лист значении номера начала каждого раздела, для ориентации по ним
        public static void OrientRazdel(List<Excel.Range> cellsAllRazdel, int numPozSmeta, int numCellspoNumPozSmeta, ref int indexAllRazdel, ref List<int> numRowStartRazdel)
        {
            //Console.WriteLine("  OrientRazdel");
            if (indexAllRazdel < cellsAllRazdel.Count)
            {
                if (numCellspoNumPozSmeta > cellsAllRazdel[indexAllRazdel].Row)
                {
                    numRowStartRazdel.Add(numPozSmeta);
                    indexAllRazdel++;
                }
            }
            else return;
        }

        //возвращает  словарь, где ключ - номер по смете, значение - номер строки в массиве наименования данных работ для одного раздела
        public static Dictionary<int, int> PoradokRazdel(Regex regulNameWorkOfRazdel, string[] valueNameofEachWork, int[] keynumerTrudozatratEachWork)
        {
            //Console.WriteLine(" PoradokRazdel");
            MatchCollection mathesNameWork;
            Dictionary<int, int> inRazdelnumerPozandnumWorkinArr = new Dictionary<int, int>();
            for (int i = 0; i < valueNameofEachWork.Length; i++)
            {
                mathesNameWork = regulNameWorkOfRazdel.Matches(valueNameofEachWork[i]);
                if (mathesNameWork.Count > 0)
                {
                    Console.WriteLine(keynumerTrudozatratEachWork[i] + " + " + i);
                    inRazdelnumerPozandnumWorkinArr.Add(keynumerTrudozatratEachWork[i], i);
                }
            }
            return inRazdelnumerPozandnumWorkinArr;
        }
        //возвращает строку из файла содержащего все выходные дни с 1999 по 2025 г, строку искомого месяца
        private static string Finddaymes(Excel.Range rangeData, DataVvod dataStartWork)
        {
            //Console.WriteLine(" Finddaymes");
            Excel.Range findYear = rangeData.Find(dataStartWork.YearStart.ToString());
            string findYearString = findYear.Value.ToString();
            int numFirstQuotes = dataStartWork.MonthStart * 2 - 1;
            int numLastQuotes = dataStartWork.MonthStart * 2;
            int countQuotes = 0;
            string freeDaysinMonthPropis = null;
            for (int i = 0; i < findYearString.Length; i++)
            {
                if (findYearString[i] == '"')
                {
                    countQuotes++;
                }
                if (countQuotes == numFirstQuotes && countQuotes < numLastQuotes && findYearString[i] != '"')
                {
                    freeDaysinMonthPropis += findYearString[i];
                }
            }
            return freeDaysinMonthPropis;
        }

        //возвращает все рабочие дни определенного месяца с цчетом того с какого дня начались работы
        private static List<int> RabDni(string freeDaysinMonthPropis, int dayStart, int amountDaysinMonth, int daysforWork)
        {
            //Console.WriteLine(" RabDni");
            List<int> freeDaysinMonthInt = new List<int>();
            int freeDayInt;
            for (int i = 0; i < freeDaysinMonthPropis.Length - 1; i++)
            {
                if (i == 0)
                {
                    if (freeDaysinMonthPropis[i] >= '0' && freeDaysinMonthPropis[i] <= '9' && (freeDaysinMonthPropis[i + 1] < '0' || freeDaysinMonthPropis[i + 1] > '9'))
                    {
                        freeDayInt = freeDaysinMonthPropis[i] - '0';
                        freeDaysinMonthInt.Add(freeDayInt);
                    }
                }
                else
                {
                    if ((freeDaysinMonthPropis[i - 1] < '0' || freeDaysinMonthPropis[i - 1] > '9') && freeDaysinMonthPropis[i] >= '0' && freeDaysinMonthPropis[i] <= '9' && (freeDaysinMonthPropis[i + 1] < '0' || freeDaysinMonthPropis[i + 1] > '9'))
                    {
                        freeDayInt = freeDaysinMonthPropis[i] - '0';
                        freeDaysinMonthInt.Add(freeDayInt);
                    }
                }
                if (freeDaysinMonthPropis[i] >= '0' && freeDaysinMonthPropis[i] <= '9' && freeDaysinMonthPropis[i + 1] >= '0' && freeDaysinMonthPropis[i + 1] <= '9')
                {
                    freeDayInt = (freeDaysinMonthPropis[i] - '0') * 10 + (freeDaysinMonthPropis[i + 1] - '0');
                    freeDaysinMonthInt.Add(freeDayInt);
                }

            }
            //for (int i = 0; i < mesdni.Count; i++)
            //{ Console.WriteLine(mesdni[i]); }
            List<int> workDaysinMonth = new List<int>();

            for (int i = dayStart; i <= amountDaysinMonth; i++)
            {
                int countFreeDayPodrad = 0;
                for (int j = 0; j < freeDaysinMonthInt.Count; j++)
                {
                    if (i == freeDaysinMonthInt[j])
                    {
                        countFreeDayPodrad++;
                    }
                }
                if (countFreeDayPodrad == 0)
                {
                    workDaysinMonth.Add(i);
                }
                if (workDaysinMonth.Count == daysforWork) break;
            }
            return workDaysinMonth;
        }
        //возвращает количество дней в каждом месяце
        private static int DayinMonth(DataVvod dataStartWork)
        {
            //Console.WriteLine(" DayinMonth");
            int amountDaysinMonth = 0;
            switch (dataStartWork.MonthStart)
            {
                case 1: amountDaysinMonth = 31; break;
                case 2:
                    {
                        if (dataStartWork.YearStart % 4 == 0)
                        {
                            amountDaysinMonth = 29;
                        }
                        else
                        {
                            amountDaysinMonth = 28;
                        }
                        break;
                    }
                case 3: amountDaysinMonth = 31; break;
                case 4: amountDaysinMonth = 30; break;
                case 5: amountDaysinMonth = 31; break;
                case 6: amountDaysinMonth = 30; break;
                case 7: amountDaysinMonth = 31; break;
                case 8: amountDaysinMonth = 31; break;
                case 9: amountDaysinMonth = 30; break;
                case 10: amountDaysinMonth = 31; break;
                case 11: amountDaysinMonth = 30; break;
                case 12: amountDaysinMonth = 31; break;
            }
            return amountDaysinMonth;
        }
        //меняет по ссылке словарь, состоящий из строки формата месяц, год и листа из рабочих дней
        public static Dictionary<string, List<int>> DninaRabotyZadan(Excel.Range rangeData, DataVvod dataStartWork, int monthsforWork, int daysforWork)
        {
            //Console.WriteLine(" DninaRabotyZadan");
            List<int> workDaysinMonth;
            Dictionary<string, List<int>> dayonEachWork = new Dictionary<string, List<int>>();
            string freeDaysinMonthPropis;
            int amountDaysinMonth;
            string monthAndYearForGrafik = null;
            for (int v = 0; v < monthsforWork + 1; v++)
            {
                if (v == 0)
                {
                    freeDaysinMonthPropis = Finddaymes(rangeData, dataStartWork);
                    amountDaysinMonth = DayinMonth(dataStartWork);
                    workDaysinMonth = RabDni(freeDaysinMonthPropis, dataStartWork.DayStart, amountDaysinMonth, daysforWork);
                    monthAndYearForGrafik = $"{MonthpropisInt(dataStartWork.MonthStart)}.{ dataStartWork.YearStart.ToString()}";                    //Console.WriteLine(rez);
                }
                else
                {
                    if (dataStartWork.MonthStart < 12)
                    {
                        dataStartWork.MonthStart += 1;
                        amountDaysinMonth = DayinMonth(dataStartWork);
                        freeDaysinMonthPropis = Finddaymes(rangeData, dataStartWork);
                    }
                    else
                    {
                        dataStartWork.MonthStart = 1;
                        dataStartWork.YearStart += 1;
                        amountDaysinMonth = DayinMonth(dataStartWork);
                        freeDaysinMonthPropis = Finddaymes(rangeData, dataStartWork);
                    }
                    workDaysinMonth = RabDni(freeDaysinMonthPropis, 1, amountDaysinMonth, daysforWork);
                    monthAndYearForGrafik = $"{MonthpropisInt(dataStartWork.MonthStart)}.{ dataStartWork.YearStart.ToString()}";
                }
                dayonEachWork.Add(monthAndYearForGrafik, workDaysinMonth);
                daysforWork -= workDaysinMonth.Count;
                if (daysforWork <= 0) break;
            }
            return dayonEachWork;
        }
    }
}
