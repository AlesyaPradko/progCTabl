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
        public static Excel.Workbook CopyExcelSmetaOne(string adresSmeta, string testUserWhereSave, Excel.Application excelApp)
        {
            //Console.WriteLine("CopyExcelSmetaOne");
            Excel.Workbook excelBooksm = excelApp.Workbooks.Open(adresSmeta);
            if (!File.Exists(testUserWhereSave))
            {
                excelBooksm.SaveCopyAs(testUserWhereSave);
            }
            excelBooksm.Close(false, Type.Missing, Type.Missing);
            Excel.Workbook excelBookcopySmet = excelApp.Workbooks.Open(testUserWhereSave);
            return excelBookcopySmet;
        }
        //метод возвращает лист строк с адресами смет и кс
        public static List<string> GetstringAdres(string userAdress)
        {
            //Console.WriteLine("GetstringAdresa");
            string[] allAdresContainsFolder = Directory.GetFiles(userAdress);
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
        public static List<Excel.Workbook> GetBookAllAktandSmeta(string userKS, Excel.Application excelApp)
        {
            //Console.WriteLine("GetBookAllAktKSandSmeta");
            string[] nameAdresAktKS = Directory.GetFiles(userKS);
            List<Excel.Workbook> containFolderKS = new List<Excel.Workbook>();
            foreach (string oneNameKS in nameAdresAktKS)
            {
                if (oneNameKS.Contains("~$")) continue;
                if (!oneNameKS.Contains(".xlsx")) continue;
                else
                {
                    Excel.Workbook bookAktKS = excelApp.Workbooks.Open(oneNameKS);
                    containFolderKS.Add(bookAktKS);
                }
            }
            return containFolderKS;
        }

        //метод возвращает словарь, ключ - адрес сметы, значение - адреса актов КС, относящихся к смете
        public static Dictionary<string, List<string>> GetContainAktKSinOneSmeta(List<Excel.Workbook> ContainFolderKS, List<string> AdresSmeta, List<string> AdresAktKS)
        {
            // Console.WriteLine("GetContainAktKSinOneSmeta");
            Dictionary<string, List<string>> aktAllKSforOneSmeta = new Dictionary<string, List<string>>();
            RegexReg reg = new RegexReg();
            for (int u = 0; u < AdresSmeta.Count; u++)
            {
                string numberSmeta = null;
                MatchCollection mathesNumerSmeta = reg.nameSmeta.Matches(AdresSmeta[u]);
                //Console.WriteLine(" smeta " + AdresSmeta[u]);
                if (mathesNumerSmeta.Count > 0)
                {
                    foreach (Match numSmeta in mathesNumerSmeta)
                    {
                        numberSmeta = numSmeta.Value;
                    }
                    int foundS1 = numberSmeta.IndexOf("№");
                    if (foundS1 != -1)
                    {
                        numberSmeta = numberSmeta.Remove(0, 1 + foundS1);
                        List<string> aktKSforSmeta = new List<string>();
                        for (int c = 0; c < ContainFolderKS.Count; c++)
                        {
                            Excel.Worksheet workShetAktKS = ContainFolderKS[c].Sheets[1];
                            Excel.Range rangAktKS = workShetAktKS.get_Range("A1", "Q40");
                            if (rangAktKS.Find(numberSmeta) == null) continue;
                            else
                            {
                                aktKSforSmeta.Add(AdresAktKS[c]);
                                //Console.WriteLine("AdresKS[c] " + numberSmeta + " " + AdresAktKS[c]);
                            }
                            Marshal.FinalReleaseComObject(rangAktKS);
                            Marshal.FinalReleaseComObject(workShetAktKS);
                        }
                        aktAllKSforOneSmeta.Add(AdresSmeta[u], aktKSforSmeta);
                    }
                }           
            }        
            return aktAllKSforOneSmeta;
        }
        //метод возвращает необходимую ячейку
        public static Excel.Range FindCellOfRegul(Excel.Worksheet workSheetAktKS, Excel.Range rangeAktKS, Regex regul)
        {
            //Console.WriteLine("FindCellofRegul");
            MatchCollection mathesFindCell;
            Excel.Range findCellsColumn = null;
            for (int u = 1; u <= rangeAktKS.Rows.Count; u++)
            {
                for (int j = 1; j <= rangeAktKS.Columns.Count; j++)
                {
                    Excel.Range nextCellinAktKS = workSheetAktKS.Cells[u, j];
                    if (nextCellinAktKS != null && nextCellinAktKS.Value != null && nextCellinAktKS.ToString() != "")
                    {
                        mathesFindCell = regul.Matches(nextCellinAktKS.Value.ToString());
                    }
                    else continue;
                    if (mathesFindCell.Count > 0)
                    {
                        findCellsColumn = workSheetAktKS.Cells[u, j];
                        break;
                    }
                    if (findCellsColumn != null) break;
                }
            }
            return findCellsColumn;
        }
        //метод возвращает строку с записанной датой составления акта для последующей обработки
        public static string FindDateAktKS(Regex monthorYear, Excel.Range finddata)
        {
            //Console.WriteLine("FindDateAktKS");
            string dateMonthOrYear = null;
            MatchCollection yearMonth = monthorYear.Matches(finddata.Value.ToString());
            if (yearMonth.Count > 0)
            {
                foreach (Match oneDate in yearMonth)
                {
                    dateMonthOrYear = oneDate.Value;
                }
            }
            return dateMonthOrYear;
        }
        //метод возвращает строку где месяц записан прописью, входной параметр число
        public static string MonthpropisInt(int montStart)
        {
            //Console.WriteLine("MonthpropisInt");
            string monthLetter = null;
            switch (montStart)
            {
                case 1: monthLetter = "январь"; break;
                case 2: monthLetter = "февраль"; break;
                case 3: monthLetter = "март"; break;
                case 4: monthLetter = "апрель"; break;
                case 5: monthLetter = "май"; break;
                case 6: monthLetter = "июнь"; break;
                case 7: monthLetter = "июль"; break;
                case 8: monthLetter = "август"; break;
                case 9: monthLetter = "сентябрь"; break;
                case 10: monthLetter = "октябрь"; break;
                case 11: monthLetter = "ноябрь"; break;
                case 12: monthLetter = "декабрь"; break;
            }
            return monthLetter;
        }

        //метод возвращает строку где месяц записан прописью, входной параметр строкак как число
        public static string MonthLetter(string monthAktKS)
        {
            //Console.WriteLine("Monthpropis");
            string monthAktKSpropis;
            int month = 0;
            int twoDigit = 10;
            for (int j = 0; j < monthAktKS.Length; j++)
            {
                if (monthAktKS[j] >= '0' && monthAktKS[j] <= '9')
                {
                    month += (monthAktKS[j] - '0') * twoDigit;
                    twoDigit /= 10;
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
            for (int j = keyNumPozpoSmeteinAktKS.Row + 3; j < rangeAktKS.Rows.Count + keyNumPozpoSmeteinAktKS.Row; j++)
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
        public static Dictionary<int, T> GetkeySmetaForRecord<T>(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, string AdresSmeta)
        {
            //Console.WriteLine("GetkeySmetaForZapis<T>");
            Dictionary<int, T> resultwithNumPoz = new Dictionary<int, T>();
            int numPozSmeta;
            T zerovalue;
            for (int j = rangeSmetaOne.Row + 4; j < rangeSmetaOne.Rows.Count + rangeSmetaOne.Row+4; j++)
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

        
        
        //метод удаляет ненужные столбцы и строки для формирования ведомости выполненных объемов работ
        public static void DeleteColumnandRow(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, Excel.Range keyCellNomerpozSmeta, string AdresSm, ref int lastRowCellsafterDelete)
        {
            //Console.WriteLine(" DeleteColumnandRow");
            try
            {
                int amountRow = 0;
                List<int> deleteExcessCells = new List<int>();
                for (int u = keyCellNomerpozSmeta.Row + 5; u <= rangeSmetaOne.Rows.Count; u++)
                {
                    
                    Excel.Range cellsFirstColumnTabl = SheetcopySmetaOne.Cells[u, keyCellNomerpozSmeta.Column];
                    if (cellsFirstColumnTabl.MergeCells && !cellsFirstColumnTabl.Value.ToString().Contains("Раздел"))
                    {
                        deleteExcessCells.Add(cellsFirstColumnTabl.Row);
                    }
                    if ( cellsFirstColumnTabl.Value!=null && cellsFirstColumnTabl.Value.ToString()!="")
                    {
                        amountRow++;
                    }
                }
                Excel.Range lastcellsFirstColumnTabl = SheetcopySmetaOne.Cells[rangeSmetaOne.Rows.Count, keyCellNomerpozSmeta.Column];
                if (lastcellsFirstColumnTabl != null && lastcellsFirstColumnTabl.Value != null && lastcellsFirstColumnTabl.Value2.ToString() != "")
                {
                    throw new ZapredelException($"Вы задали слишком малую высоту для {AdresSm}");
                }
                deleteExcessCells.Reverse();
                if (deleteExcessCells.Count != 0)
                {
                    lastRowCellsafterDelete = deleteExcessCells[0] - deleteExcessCells.Count;
                }
                else
                {
                    lastRowCellsafterDelete = keyCellNomerpozSmeta.Row + 5 + amountRow;
                }      
                for (int u = rangeSmetaOne.Rows.Count; u > keyCellNomerpozSmeta.Row + 4; u--)
                {
                    Excel.Range cellsFirstColumnTabl = SheetcopySmetaOne.Cells[u, keyCellNomerpozSmeta.Column];
                    for (int v = 0; v < deleteExcessCells.Count; v++)
                    {
                        if (cellsFirstColumnTabl.Row == deleteExcessCells[v])
                        {
                            Excel.Range lastColumnOnDelet = SheetcopySmetaOne.Cells[cellsFirstColumnTabl.Row, rangeSmetaOne.Columns.Count];
                            Excel.Range rowOnDelet = SheetcopySmetaOne.get_Range(cellsFirstColumnTabl, lastColumnOnDelet);
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
            catch (ArgumentOutOfRangeException exc)
            {
                Console.WriteLine($"{exc.Message} вы пытаетесь повторно удалить уже удаленные ячейки");
            }
        }
        //методы для Графика
        //отсекает адрес и формат от имени сметы
        public static void GetNameSmeta(string adresSmeta, out string nameFailSmeta)
        {
            //Console.WriteLine(" GetNameSmeta");
            nameFailSmeta = null;
            int numberSlash = 0;
            for (int i = 0; i < adresSmeta.Length; i++)
            {
                if (adresSmeta[i] == '\\') numberSlash = i;
            }
            for (int i = numberSlash + 1; i < adresSmeta.Length - 5; i++)
            {
                nameFailSmeta += adresSmeta[i];
            }
        }
        //возвращает общую трудоемкость по смете в виде цифры, входной параметр - строковый
        public static double NumeralFromCell(string trudozatrata)
        {
            //Console.WriteLine(" CifrafromStringCell");
            string summaString = null;
            double trudozatratTotal = 0;
            try
            {
                for (int i = 0; i < trudozatrata.Length; i++)
                {
                    if ((trudozatrata[i] >= '0' && trudozatrata[i] <= '9') || trudozatrata[i] == ',' || trudozatrata[i] == '.')
                    {
                        summaString += trudozatrata[i];
                    }
                }
                int index = summaString.IndexOf('.');
                if (index == summaString.Length - 1)
                {
                    summaString = summaString.Remove(summaString.Length - 1, 1);
                    trudozatratTotal = Convert.ToDouble(summaString);
                }
               else if (index == - 1)
                {
                    trudozatratTotal = Convert.ToDouble(summaString);
                }
                else 
                {
                    Console.WriteLine("Проверьте, чтобы значение трудоемкости не содержало в себе знака[.]");
                } 
            }
            catch(FormatException exc)
            {
                Console.WriteLine($"{exc.Message} Проверьте, чтобы значение трудоемкости не содержало в себе знака[.]");
            }
            return trudozatratTotal;
        }
        //возвращает словарь, где ключ ячейка "Итого по разделу", значение - трудоемкость во разделу
        public static Dictionary<Excel.Range, double> FindForChapter(Excel.Worksheet workSheetSmeta, Excel.Range rangeSmeta, Excel.Range keyCellColumnTopTrudozatrat, Excel.Range keyCellNumberPozSmeta)
        {
            //Console.WriteLine("  FindpoRazdely");
            Dictionary<Excel.Range, double> forChapter = new Dictionary<Excel.Range, double>();
            RegexReg p = new RegexReg();
            MatchCollection mathes1;
            for (int j = 1; j <= rangeSmeta.Rows.Count; j++)
            {
                Excel.Range nameChapter = workSheetSmeta.Cells[j, keyCellNumberPozSmeta.Column];
                if (nameChapter != null && nameChapter.Value2 != null && nameChapter.MergeCells && nameChapter.Value2.ToString() != "")
                {
                    string s = nameChapter.Value.ToString();
                    mathes1 = p.cellTotalForChapter.Matches(s);
                }
                else continue;
                if (mathes1.Count > 0)
                {
                    Excel.Range c1 = workSheetSmeta.Cells[nameChapter.Row, keyCellColumnTopTrudozatrat.Column];
                    if (c1.Value2 != null && c1 != null && c1.Value.ToString() != "")
                    {
                        double trudChapter = Convert.ToDouble(c1.Value2);
                        forChapter.Add(nameChapter, trudChapter);
                    }
                }
            }
            return forChapter;
        }
        //возвращает лист из ячеек "Раздел такой-то"
        public static List<Excel.Range> FindChapter(Excel.Worksheet workSheetSmeta, Excel.Range rangeSmeta, Excel.Range keyCellNumberPosSmeta)
        {
            //Console.WriteLine("  FindRazdel");
            List<Excel.Range> cellsAllChapter = new List<Excel.Range>();
            RegexReg cellChapterFind = new RegexReg();
            MatchCollection mathesChapter;
            for (int j = 1; j <= rangeSmeta.Rows.Count; j++)
            {
                Excel.Range nameChapter = workSheetSmeta.Cells[j, keyCellNumberPosSmeta.Column];
                if (nameChapter != null && nameChapter.Value2 != null && nameChapter.MergeCells && nameChapter.Value2.ToString() != "")
                {
                    string stringNameChapter = nameChapter.Value.ToString();
                    mathesChapter = cellChapterFind.cellOfRazdel.Matches(stringNameChapter);
                }
                else continue;
                if (mathesChapter.Count > 0)
                {
                    cellsAllChapter.Add(nameChapter);
                }
            }
            return cellsAllChapter;
        }
        //функция меняет по ссылке лист значении номера начала каждого раздела, для ориентации по ним
        public static void OrientChapter(List<Excel.Range> cellsAllChapter, int numPosSmeta, int numCellsNumPosSmeta, ref int indexAllChapter, ref List<int> numRowStartChapter)
        {
            //Console.WriteLine("  OrientRazdel");
            if (indexAllChapter < cellsAllChapter.Count)
            {
                if (numCellsNumPosSmeta > cellsAllChapter[indexAllChapter].Row)
                {
                    numRowStartChapter.Add(numPosSmeta);
                    indexAllChapter++;
                }
            }
            else return;
        }

        //возвращает  словарь, где ключ - номер по смете, значение - номер строки в массиве наименования данных работ для одного раздела
        public static Dictionary<int, int> InOrderChapter(Regex regulNameWorkOfChapter, string[] valueNameOfEachWork, int[] keyNumberTrudozatratEachWork,int startChapt)
        {
            //Console.WriteLine(" InOrderChapter");
            MatchCollection mathesNameWork;
            Dictionary<int, int> inChapterNumberPosAndNumWorkInArr = new Dictionary<int, int>();
            for (int i = 0; i < valueNameOfEachWork.Length; i++)
            {
                mathesNameWork = regulNameWorkOfChapter.Matches(valueNameOfEachWork[i]);
                if (mathesNameWork.Count > 0)
                {
                    if (keyNumberTrudozatratEachWork[i] >= startChapt)
                    {
                        Console.WriteLine(keyNumberTrudozatratEachWork[i] + " + " + i);
                        inChapterNumberPosAndNumWorkInArr.Add(keyNumberTrudozatratEachWork[i], i);
                    }
                }
            }
            return inChapterNumberPosAndNumWorkInArr;
        }
        public static Dictionary<int, int> InOrderChapter(Regex regulNameWorkOfChapter, string[] valueNameOfEachWork, int[] keyNumberTrudozatratEachWork, int startChapt, int lastChapt)
        {
            //Console.WriteLine(" InOrderChapter");
            MatchCollection mathesNameWork;
            Dictionary<int, int> inChapterNumberPosAndNumWorkInArr = new Dictionary<int, int>();
            for (int i = 0; i < valueNameOfEachWork.Length; i++)
            {
                mathesNameWork = regulNameWorkOfChapter.Matches(valueNameOfEachWork[i]);
                if (mathesNameWork.Count > 0)
                {
                    if (keyNumberTrudozatratEachWork[i] >= startChapt && keyNumberTrudozatratEachWork[i]<lastChapt)
                    {
                        Console.WriteLine(keyNumberTrudozatratEachWork[i] + " + " + i);
                        inChapterNumberPosAndNumWorkInArr.Add(keyNumberTrudozatratEachWork[i], i);
                    }
                }
            }
            return inChapterNumberPosAndNumWorkInArr;
        }
        //возвращает строку из файла содержащего все выходные дни с 1999 по 2025 г, строку искомого месяца
        private static string FindDayMonths(Excel.Range rangeData, DataInput dataStartWork)
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
        private static List<int> GetWorksDays(string freeDaysInMonthLetter, int dayStart, int amountDaysInMonth, int daysForWork)
        {
            //Console.WriteLine(" RabDni");
            List<int> freeDaysinMonthInt = new List<int>();
            int freeDayInt;
            for (int i = 0; i < freeDaysInMonthLetter.Length - 1; i++)
            {
                if (i == 0)
                {
                    if (freeDaysInMonthLetter[i] >= '0' && freeDaysInMonthLetter[i] <= '9' && (freeDaysInMonthLetter[i + 1] < '0' || freeDaysInMonthLetter[i + 1] > '9'))
                    {
                        freeDayInt = freeDaysInMonthLetter[i] - '0';
                        freeDaysinMonthInt.Add(freeDayInt);
                    }
                }
                else
                {
                    if ((freeDaysInMonthLetter[i - 1] < '0' || freeDaysInMonthLetter[i - 1] > '9') && freeDaysInMonthLetter[i] >= '0' && freeDaysInMonthLetter[i] <= '9' && (freeDaysInMonthLetter[i + 1] < '0' || freeDaysInMonthLetter[i + 1] > '9'))
                    {
                        freeDayInt = freeDaysInMonthLetter[i] - '0';
                        freeDaysinMonthInt.Add(freeDayInt);
                    }
                }
                if (freeDaysInMonthLetter[i] >= '0' && freeDaysInMonthLetter[i] <= '9' && freeDaysInMonthLetter[i + 1] >= '0' && freeDaysInMonthLetter[i + 1] <= '9')
                {
                    freeDayInt = (freeDaysInMonthLetter[i] - '0') * 10 + (freeDaysInMonthLetter[i + 1] - '0');
                    freeDaysinMonthInt.Add(freeDayInt);
                }

            }
            List<int> workDaysinMonth = new List<int>();
            for (int i = dayStart; i <= amountDaysInMonth; i++)
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
                if (workDaysinMonth.Count == daysForWork) break;
            }
            return workDaysinMonth;
        }
        //возвращает количество дней в каждом месяце
        private static int GetDaysInMonth(DataInput dataStartWork)
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
        public static Dictionary<string, List<int>> GetDaysForWork(Excel.Range rangeData, DataInput dataStartWork, int monthsForWork, int daysForWork)
        {
            //Console.WriteLine(" GetDaysForWork");
            List<int> workDaysinMonth;
            Dictionary<string, List<int>> dayOnEachWork = new Dictionary<string, List<int>>();
            string freeDaysInMonthLetter;
            int amountDaysInMonth;
            string monthAndYearForGrafik = null;
            for (int v = 0; v < monthsForWork + 1; v++)
            {
                if (v == 0)
                {
                    freeDaysInMonthLetter = FindDayMonths(rangeData, dataStartWork);
                    amountDaysInMonth = GetDaysInMonth(dataStartWork);
                    workDaysinMonth = GetWorksDays(freeDaysInMonthLetter, dataStartWork.DayStart, amountDaysInMonth, daysForWork);
                    monthAndYearForGrafik = $"{MonthpropisInt(dataStartWork.MonthStart)}.{ dataStartWork.YearStart.ToString()}";                    //Console.WriteLine(rez);
                }
                else
                {
                    if (dataStartWork.MonthStart < 12)
                    {
                        dataStartWork.MonthStart += 1;
                        amountDaysInMonth = GetDaysInMonth(dataStartWork);
                        freeDaysInMonthLetter = FindDayMonths(rangeData, dataStartWork);
                    }
                    else
                    {
                        dataStartWork.MonthStart = 1;
                        dataStartWork.YearStart += 1;
                        amountDaysInMonth = GetDaysInMonth(dataStartWork);
                        freeDaysInMonthLetter = FindDayMonths(rangeData, dataStartWork);
                    }
                    workDaysinMonth = GetWorksDays(freeDaysInMonthLetter, 1, amountDaysInMonth, daysForWork);
                    monthAndYearForGrafik = $"{MonthpropisInt(dataStartWork.MonthStart)}.{ dataStartWork.YearStart.ToString()}";
                }
                dayOnEachWork.Add(monthAndYearForGrafik, workDaysinMonth);
                daysForWork -= workDaysinMonth.Count;
                if (daysForWork <= 0) break;
            }
            return dayOnEachWork;
        }
    }
}
