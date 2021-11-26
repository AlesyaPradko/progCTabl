using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
public enum ChangeSelect { DaysforWork = 49, NumberofWorker = 50 };

namespace ConsoleApp3
{
    public class Grafik
    {
        private List<Excel.Workbook> containPapkaSmeta;
        private List<string> adresSmeta;
        private Dictionary<Excel.Range, double> poRazdelyTrudozatrat;
        private List<Excel.Range> cellsAllRazdel;
        private Dictionary<int, double> chelChasforEachWork;
        private List<int> numRowStartRazdel;
        private Dictionary<int, string> nameforEachWorkinSmeta;
        private Dictionary<string, List<int>> dayonEachWork;
        private List<Dictionary<int, int>> allRazdelpoPoradky;
        private double trudozatratTotal;
        public Grafik()
        { }


        //метод инициализирует листы и словари хранящие в себе сметы (адреса и книги)
        public void InitializationGrafik(Excel.Application excelApp)
        {
            //Console.WriteLine("InitializationGrafik");
            try
            {
                string usersmetu = @"D:\иксу";
                containPapkaSmeta = ParserExc.GetBookAllAktKSandSmeta(usersmetu, excelApp);
                adresSmeta = ParserExc.GetstringAdresa(usersmetu);
                if (containPapkaSmeta.Count == 0 || adresSmeta.Count == 0)
                {
                    throw new DonthaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку");
                }
            }
            catch (DonthaveExcelException ex)
            {
                Console.WriteLine(ex.parName);
            }
        }
        //метод для работы над папкой со сметами в режиме график
        public void ProccessGrafik(RangeFile obl, Excel.Application excelApp)
        {
            for (int numSmeta = 0; numSmeta < containPapkaSmeta.Count; numSmeta++)
            {
                WorkGrafik(numSmeta, obl, excelApp);
            }
        }
        //работа надсметой для записи графика
        private void WorkGrafik(int numSmeta, RangeFile obl, Excel.Application excelApp)
        {
            //Console.WriteLine("WorkGrafik");
            try
            {
                Excel.Worksheet workSheetoneSmeta = containPapkaSmeta[numSmeta].Sheets[1];
                Excel.Range rangeoneSmeta = workSheetoneSmeta.get_Range(obl.FirstCell, obl.LastCell);
                Excel.Range keyCellNumberPozSmeta = rangeoneSmeta.Find("№ пп");
                Excel.Range keyCellColumnTopTrudozatrat = rangeoneSmeta.Find("Т/з осн. раб. Всего");
                Excel.Range cellwithTrudozatrat = rangeoneSmeta.Find("Сметная трудоемкость");
                string trudozatrataString = cellwithTrudozatrat.Value.ToString();
                string nameFailSmeta;
                ParserExc.GetNameSmeta(adresSmeta[numSmeta], out nameFailSmeta);
                trudozatratTotal = ParserExc.CifrafromStringCell(trudozatrataString);
                poRazdelyTrudozatrat = ParserExc.FindpoRazdely(workSheetoneSmeta, rangeoneSmeta, keyCellColumnTopTrudozatrat, keyCellNumberPozSmeta);
                cellsAllRazdel = ParserExc.FindRazdel(workSheetoneSmeta, rangeoneSmeta, keyCellNumberPozSmeta);
                numRowStartRazdel = new List<int>();
                chelChasforEachWork = ChelChaspoRabotam(workSheetoneSmeta, rangeoneSmeta, keyCellNumberPozSmeta, keyCellColumnTopTrudozatrat, adresSmeta[numSmeta]);
                nameforEachWorkinSmeta = NameWorkinPozSmeta(workSheetoneSmeta, rangeoneSmeta, keyCellNumberPozSmeta, adresSmeta[numSmeta]);
                //для базы данных начало
                Regex nameSmeta = new Regex(@"Тепловые сети", RegexOptions.IgnoreCase);
                List<Regex> RazdelAll = new List<Regex>();
                Regex razdel1 = new Regex(@"Демонтаж", RegexOptions.IgnoreCase);
                Regex razdel2 = new Regex(@"Землян", RegexOptions.IgnoreCase);
                Regex razdel3 = new Regex(@"Общестроит", RegexOptions.IgnoreCase);
                Regex razdel456 = new Regex(@"(Тепловые|Конденсат|Паропров)", RegexOptions.IgnoreCase);
                Regex razdel7 = new Regex(@"Изоляц", RegexOptions.IgnoreCase);
                Regex razdel8 = new Regex(@"СОДК", RegexOptions.IgnoreCase);
                RazdelAll.Add(razdel1);
                RazdelAll.Add(razdel2);
                RazdelAll.Add(razdel3);
                RazdelAll.Add(razdel456);
                RazdelAll.Add(razdel7);
                RazdelAll.Add(razdel8);
                List<Regex> FORRazdelAll = new List<Regex>();
                Regex forrazdel1 = new Regex(@"Демонтаж", RegexOptions.IgnoreCase);
                Regex forrazdel2 = new Regex(@"(Засыпка)", RegexOptions.IgnoreCase);
                Regex forrazdel3 = new Regex(@"(Заливка)", RegexOptions.IgnoreCase);
                Regex forrazdel456 = new Regex(@"(Установка|прокладка)", RegexOptions.IgnoreCase);
                Regex forrazdel7 = new Regex(@"(Изоляция|Покрытие|Нанесение)", RegexOptions.IgnoreCase);
                Regex forrazdel8 = new Regex(@"(Протяжка)", RegexOptions.IgnoreCase);
                FORRazdelAll.Add(forrazdel1);
                FORRazdelAll.Add(forrazdel2);
                FORRazdelAll.Add(forrazdel3);
                FORRazdelAll.Add(forrazdel456);
                FORRazdelAll.Add(forrazdel7);
                FORRazdelAll.Add(forrazdel8);
                //для базы данных конец
                allRazdelpoPoradky = new List<Dictionary<int, int>>();
                for (int i = 0; i < RazdelAll.Count; i++)
                {
                    RasstanovkaAllWorkspoPoradky(RazdelAll[i], FORRazdelAll[i], ref allRazdelpoPoradky);
                }
                Excel.Workbook excelBookData = excelApp.Workbooks.Open(@"D:\даты\data-20191112T1252-structure-20191112T1247.csv");
                Excel.Worksheet workSheetData = excelBookData.Sheets[1];
                Excel.Range rangeData = workSheetData.get_Range("A1", "A30");
                int daysforWork = 0, numberofWorkers = 0;
                DataVvod dataStartWork = new DataVvod();
                VvodDataorNumberofWorkers(ref dataStartWork, ref daysforWork, ref numberofWorkers);
                int monthsforWork;
                double delta = (daysforWork / 21.0) - (int)(daysforWork / 21);
                if (delta < 0.04) monthsforWork = daysforWork / 21;
                else monthsforWork = 1 + daysforWork / 21;
                dayonEachWork = ParserExc.DninaRabotyZadan(rangeData, dataStartWork, monthsforWork, daysforWork);
                string grafikAdress = @"D:\график\результат\График производства работ - " + nameFailSmeta;
                ZapisGrafik(excelApp, workSheetoneSmeta, grafikAdress, numberofWorkers);
                object misValue = System.Reflection.Missing.Value;
                Marshal.FinalReleaseComObject(rangeData);
                Marshal.FinalReleaseComObject(workSheetData);
                excelBookData.Close(false, misValue, misValue);
                Marshal.FinalReleaseComObject(excelBookData);
                Marshal.FinalReleaseComObject(rangeoneSmeta);
                Marshal.FinalReleaseComObject(workSheetoneSmeta);
                containPapkaSmeta[numSmeta].Close(true, misValue, misValue);
                Marshal.FinalReleaseComObject(containPapkaSmeta[numSmeta]);
            }
            catch (NullReferenceException ex)
            {
                Console.WriteLine($"{ex.Message} Проверьте чтобы в {adresSmeta[numSmeta]} было верно записано устойчивое выражение [№ пп] или [Кол.] или трудоемкость");
            }
        }
        //метод для ввода исходных данных от пользователя  для построения графика
        private void VvodDataorNumberofWorkers(ref DataVvod dataStartWork, ref int daysforWork, ref int numberofWorkers)
        {
            //Console.WriteLine("VvodDataorNumberofWorkers");
            try
            {
                Console.WriteLine("Введите дату начала работ день");
                dataStartWork.DayStart = Int32.Parse(Console.ReadLine());
                Console.WriteLine("Введите дату начала работ месяц");
                dataStartWork.MonthStart = Int32.Parse(Console.ReadLine());
                Console.WriteLine("Введите дату начала работ год(1999-2022)");
                dataStartWork.YearStart = Int32.Parse(Console.ReadLine());
                Console.WriteLine("Выберите, вы хотите ввести количество месяцев на работы(1) или количество человек в бригаде(2)");
                var selectDayorWorker = (ChangeSelect)Console.ReadKey().Key;
                double deltaLessthenOne = 0;
                switch (selectDayorWorker)
                {
                    case ChangeSelect.DaysforWork:
                        {
                            Console.WriteLine("Введите количество дней, планируемых на данные виды работ");
                            daysforWork = Int32.Parse(Console.ReadLine());
                            numberofWorkers = (int)(trudozatratTotal / (daysforWork * 8));
                            deltaLessthenOne = (trudozatratTotal / (daysforWork * 8)) - numberofWorkers;
                            Console.WriteLine($"Nrab = {numberofWorkers}  dayrab { daysforWork} delta {deltaLessthenOne}");
                            if (deltaLessthenOne >= 0.5)
                            {
                                numberofWorkers += 1;
                                daysforWork += 1;
                            }
                            else
                            {
                                daysforWork += 2;
                            }
                            break;
                        }
                    case ChangeSelect.NumberofWorker:
                        {
                            Console.WriteLine("Введите количество человек в бригаде, планируемых на данные виды работ");
                            numberofWorkers = Int32.Parse(Console.ReadLine());
                            daysforWork = (int)(trudozatratTotal / (numberofWorkers * 8));
                            deltaLessthenOne = (trudozatratTotal / (numberofWorkers * 8)) - daysforWork;
                            Console.WriteLine($"Nrab = {numberofWorkers}  dayrab { daysforWork} delta {deltaLessthenOne}");
                            if (deltaLessthenOne > 0.05) daysforWork += 2;
                            else daysforWork += 1;
                            Console.WriteLine($"Nrab = {numberofWorkers}  dayrab { daysforWork} delta {deltaLessthenOne}");
                            break;
                        }
                    default:
                        Console.WriteLine("Вы ввели неверный символ");
                        break;
                }
            }
            catch (FormatException ex)
            {
                Console.WriteLine($"{ex.Message} вы ввели неверный формат данных");
            }
        }


        //возвращает  словарь, где ключ - номер по смете, значение - трудозатраты на данную работу
        private Dictionary<int, double> ChelChaspoRabotam(Excel.Worksheet workSheetoneSmeta, Excel.Range rangeoneSmeta, Excel.Range keyCellNumberPozSmeta, Excel.Range keyCellColumnTopTrudozatrat, string AdresSmeta)
        {
            //Console.WriteLine("ChelChaspoRabotam");
            Dictionary<int, double> chelChasforEachWork = new Dictionary<int, double>();
            double trudozatratofWork;
            int numPozSmeta, indexAllRazdel = 0;
            for (int j = keyCellNumberPozSmeta.Row + 4; j <= rangeoneSmeta.Rows.Count; j++)
            {
                Excel.Range cellsNumberPozColumnTabl = workSheetoneSmeta.Cells[j, keyCellNumberPozSmeta.Column];
                Excel.Range cellsColumnTrudozatrat = workSheetoneSmeta.Cells[j, keyCellColumnTopTrudozatrat.Column];
                if (cellsNumberPozColumnTabl != null && cellsNumberPozColumnTabl.Value2 != null && !cellsNumberPozColumnTabl.MergeCells && cellsNumberPozColumnTabl.Value2.ToString() != "" && cellsColumnTrudozatrat != null && cellsColumnTrudozatrat.Value2 != null && !cellsColumnTrudozatrat.MergeCells && cellsColumnTrudozatrat.Value2.ToString() != "")
                {
                    try
                    {
                        int numCellspoNumPozSmeta = cellsNumberPozColumnTabl.Row;
                        numPozSmeta = Convert.ToInt32(cellsNumberPozColumnTabl.Value2);
                        ParserExc.OrientRazdel(cellsAllRazdel, numPozSmeta, numCellspoNumPozSmeta, ref indexAllRazdel, ref numRowStartRazdel);
                        trudozatratofWork = Convert.ToDouble(cellsColumnTrudozatrat.Value2);
                        chelChasforEachWork.Add(numPozSmeta, trudozatratofWork);
                    }
                    catch (NullReferenceException ex)
                    {
                        Console.WriteLine($"{ex.Message} Проверьте чтобы в {AdresSmeta} было верно записано устойчивое выражение [Наименование]");
                    }
                    catch (ArgumentException ex)
                    {
                        Console.WriteLine($"{ex.Message} Проверьте чтобы в {AdresSmeta} не повторялись значения позиций по смете в строке {cellsNumberPozColumnTabl.Row}");
                    }
                    catch (FormatException ex)
                    {
                        Console.WriteLine($"{ex.Message} Вы ввели неверный формат для {AdresSmeta} в строке {cellsNumberPozColumnTabl.Row} в столбце {cellsNumberPozColumnTabl.Column}(не должно быть [., букв], только целые числа,или в столбце {cellsColumnTrudozatrat} только числа дробные, не должно быть [.букв]  )");
                    }
                }
            }
            return chelChasforEachWork;
        }
        //возвращает  словарь, где ключ - номер по смете, значение - строковое наименование данных работ
        private Dictionary<int, string> NameWorkinPozSmeta(Excel.Worksheet workSheetoneSmeta, Excel.Range rangeoneSmeta, Excel.Range keyCellNumberPozSmeta, string AdresSmeta)
        {
            //Console.WriteLine("NameWorkinPozSmeta");
            int[] keyTrudozatratEachWork = chelChasforEachWork.Keys.ToArray();
            Dictionary<int, string> nameforEachWorkinSmeta = new Dictionary<int, string>();
            int numPozSmeta;
            string nameWorkinPozSmeta;
            Excel.Range keyCellNameWork = rangeoneSmeta.Find("Наименование");
            for (int j = keyCellNumberPozSmeta.Row + 4; j <= rangeoneSmeta.Rows.Count; j++)
            {

                Excel.Range cellsNumberPozColumnTabl = workSheetoneSmeta.Cells[j, keyCellNumberPozSmeta.Column];
                Excel.Range cellsNameWorkColumnTabl = workSheetoneSmeta.Cells[j, keyCellNameWork.Column];
                if (cellsNumberPozColumnTabl != null && cellsNumberPozColumnTabl.Value2 != null && !cellsNumberPozColumnTabl.MergeCells && cellsNumberPozColumnTabl.Value2.ToString() != "" && cellsNameWorkColumnTabl != null && cellsNameWorkColumnTabl.Value2 != null && !cellsNameWorkColumnTabl.MergeCells && cellsNameWorkColumnTabl.Value2.ToString() != "")
                {
                    try
                    {
                        for (int i = 0; i < keyTrudozatratEachWork.Length; i++)
                        {

                            numPozSmeta = Convert.ToInt32(cellsNumberPozColumnTabl.Value2);
                            if (numPozSmeta == keyTrudozatratEachWork[i])
                            {
                                nameWorkinPozSmeta = cellsNameWorkColumnTabl.Value.ToString();
                                nameforEachWorkinSmeta.Add(numPozSmeta, nameWorkinPozSmeta);
                            }
                        }
                    }
                    catch (ArgumentException ex)
                    {
                        Console.WriteLine($"{ex.Message} Проверьте чтобы в {AdresSmeta} не повторялись значения позиций по смете в строке {cellsNumberPozColumnTabl.Row}");
                    }
                    catch (FormatException ex)
                    {
                        Console.WriteLine($"{ex.Message} Вы ввели неверный формат для {AdresSmeta} в строке {cellsNumberPozColumnTabl.Row} в столбце {cellsNumberPozColumnTabl.Column}(не должно быть [., букв], только целые числа.");
                    }
                }
            }
            return nameforEachWorkinSmeta;
        }

        //меняет по ссылке лист, состоящий из словарей,где ключ - номер по смете, значение - номер строки в массиве наименования данных работ для всех разделов
        private void RasstanovkaAllWorkspoPoradky(Regex regulNameOfRazdel, Regex regulNameWorkOfRazdel, ref List<Dictionary<int, int>> allRazdelpoPoradky)
        {
            //Console.WriteLine("RasstanovkaAllWorkspoPoradky");
            int[] keynumerTrudozatratEachWork = chelChasforEachWork.Keys.ToArray();
            string[] valueNameofEachWork = nameforEachWorkinSmeta.Values.ToArray();
            Dictionary<int, int> inRazdelnumerPozandnumWorkinArr;
            ICollection keyColl2CellRazdelTrudozatrat = poRazdelyTrudozatrat.Keys;
            foreach (Excel.Range stringPoRazdely in keyColl2CellRazdelTrudozatrat)
            {
                string stringPoRazdelyforPoisk = stringPoRazdely.Value.ToString();
                MatchCollection mathesStringRazdel = regulNameOfRazdel.Matches(stringPoRazdelyforPoisk);
                int countRazdel = 0;
                if (mathesStringRazdel.Count > 0)
                {
                    inRazdelnumerPozandnumWorkinArr = ParserExc.PoradokRazdel(regulNameWorkOfRazdel, valueNameofEachWork, keynumerTrudozatratEachWork);
                    if (inRazdelnumerPozandnumWorkinArr.Count > 0)
                    {
                        allRazdelpoPoradky.Add(inRazdelnumerPozandnumWorkinArr);
                        countRazdel++;
                    }
                }

                if (countRazdel > 0) break;
            }
        }

        //закрашивает график в соответствие с данными
        public void ZapisGrafik(Excel.Application excelApp, Excel.Worksheet workSheetoneSmeta, string grafikAdress, int numberofWorkers)
        {
            //Console.WriteLine("ZapisGrafik");
            Excel.Workbook workBookGrafik = excelApp.Workbooks.Add();
            Excel.Worksheet workSheetGrafik = (Excel.Worksheet)workBookGrafik.Worksheets.get_Item(1);
            Excel.Range FirstCellGrafik = workSheetGrafik.Range["B4"];
            Excel.Range GrafikNext = workSheetGrafik.get_Range("B4", "B5");
            GrafikNext.Merge();
            GrafikNext.Value = "№";
            GrafikNext = workSheetGrafik.get_Range("C4", "C5");
            GrafikNext.Merge();
            GrafikNext.Value = "Наименование работ";
            GrafikNext = workSheetGrafik.get_Range("D4", "D5");
            GrafikNext.Merge();
            GrafikNext.Value = "Всего чел/час";
            GrafikNext = workSheetGrafik.get_Range("E4", "E5");
            GrafikNext.Merge();
            GrafikNext.Value = "Кол. чел.  бр";
            GrafikNext = workSheetGrafik.get_Range("F4", "F5");
            GrafikNext.Merge();
            GrafikNext.Value = "Кол-во рабоч. дней";
            Excel.Range firstMonth, lastMonth = null;
            List<int>[] valueAllWorkDaysforMonth = dayonEachWork.Values.ToArray();
            string[] keyNameDataWork = dayonEachWork.Keys.ToArray();
            for (int i = 0; i < valueAllWorkDaysforMonth.Length; i++)
            {
                firstMonth = workSheetGrafik.Cells[GrafikNext.Row, GrafikNext.Column + 1];
                lastMonth = workSheetGrafik.Cells[GrafikNext.Row, GrafikNext.Column + valueAllWorkDaysforMonth[i].Count];
                for (int j = 0; j < valueAllWorkDaysforMonth[i].Count; j++)
                {
                    workSheetGrafik.Cells[firstMonth.Row + 1, firstMonth.Column + j] = valueAllWorkDaysforMonth[i][j];
                }
                GrafikNext = workSheetGrafik.get_Range(firstMonth, lastMonth);
                GrafikNext.Merge();
                GrafikNext.Value = keyNameDataWork[i];
                GrafikNext = lastMonth;
            }
            Console.WriteLine("lastMonth= " + lastMonth.Column);
            int amountOfWorkinRazdel = 0;
            int[] numRazdelTablExcelGrafik = new int[cellsAllRazdel.Count];
            Zapisstrok(workSheetoneSmeta, numberofWorkers, workSheetGrafik, ref amountOfWorkinRazdel, ref numRazdelTablExcelGrafik);
            Excel.Range LastCellGrafik = workSheetGrafik.Cells[FirstCellGrafik.Row + amountOfWorkinRazdel + 1, lastMonth.Column];
            Excel.Range forIs = workSheetGrafik.get_Range(FirstCellGrafik, LastCellGrafik);
            forIs.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
            forIs.EntireColumn.Font.Size = 10;
            forIs.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
            forIs.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
            forIs.EntireColumn.AutoFit();
            Excel.Range cellforDaysSimilarSize = workSheetGrafik.get_Range("G5", LastCellGrafik);
            cellforDaysSimilarSize.ColumnWidth = 4;
            Excel.Range rangeForColour = workSheetGrafik.get_Range("E6", LastCellGrafik);
            int AmountofDaysOnEachRazdel, AmountofWorkerOnEachRazdel;
            int AmountofDaysOnAllRazdel = 0, summaAmountofDaysEachWork = 0, summaAmountofWorkerEachWork = 0, indexofRazdel = 0;
            Console.WriteLine(FirstCellGrafik.Row + amountOfWorkinRazdel + 1);
            Console.WriteLine(lastMonth.Column);
            for (int j = rangeForColour.Row; j < rangeForColour.Rows.Count + rangeForColour.Row; j++)
            {
                if (indexofRazdel < numRazdelTablExcelGrafik.Length)
                {
                    if (j == numRazdelTablExcelGrafik[indexofRazdel])
                    {
                        indexofRazdel++;
                        Excel.Range amountofDaysEachRazdelTabl = workSheetGrafik.Cells[numRazdelTablExcelGrafik[indexofRazdel - 1], 6];
                        AmountofDaysOnEachRazdel = (int)(amountofDaysEachRazdelTabl.Value2);
                        AmountofDaysOnAllRazdel += AmountofDaysOnEachRazdel;
                    }
                }
                if (indexofRazdel > 0)
                {
                    if (j >= numRazdelTablExcelGrafik[indexofRazdel - 1] + 1)
                    {
                        Excel.Range amountofWorkerEachRazdelTabl = workSheetGrafik.Cells[numRazdelTablExcelGrafik[indexofRazdel - 1], 5];
                        AmountofWorkerOnEachRazdel = (int)(amountofWorkerEachRazdelTabl.Value2);
                        Excel.Range numberofWorkerEachWorkTabl = workSheetGrafik.Cells[j, 5];
                        Excel.Range numberofDaysEachWorkTabl = workSheetGrafik.Cells[j, 6];
                        summaAmountofWorkerEachWork += (int)(numberofWorkerEachWorkTabl.Value2);
                        if (summaAmountofWorkerEachWork < AmountofWorkerOnEachRazdel)
                        {
                            Excel.Range firstFillColour = workSheetGrafik.Cells[j, 7 + summaAmountofDaysEachWork];
                            Excel.Range lastFillColour = workSheetGrafik.Cells[j, 7 + summaAmountofDaysEachWork + (int)(numberofDaysEachWorkTabl.Value2) - 1];
                            Excel.Range rangeFillColour = workSheetGrafik.get_Range(firstFillColour, lastFillColour);
                            //Console.WriteLine("(int)(forY7.Value2) " + (int)(forY7.Value2));
                            rangeFillColour.Interior.ColorIndex = 10;
                            //Console.WriteLine(j + "er=" + er + " s " + s);
                        }
                        else
                        {
                            //Console.WriteLine("(int)(forY7.Value2) " + (int)(forY7.Value2));
                            Excel.Range firstFillColour = workSheetGrafik.Cells[j, 7 + summaAmountofDaysEachWork];
                            Excel.Range lastFillColour = workSheetGrafik.Cells[j, 7 + summaAmountofDaysEachWork + (int)(numberofDaysEachWorkTabl.Value2) - 1];
                            Excel.Range rangeFillColour = workSheetGrafik.get_Range(firstFillColour, lastFillColour);
                            rangeFillColour.Interior.ColorIndex = 10;
                            summaAmountofDaysEachWork += (int)(numberofDaysEachWorkTabl.Value2);
                            //Console.WriteLine("before er=" + er + " sk " + skoldni);
                            if (summaAmountofDaysEachWork > AmountofDaysOnAllRazdel) summaAmountofDaysEachWork -= 1; //бригада переходит на следующие работы в тот же день 
                            //Console.WriteLine("after er=" + er + " sk " + skoldni);
                            //Console.WriteLine(j + "er=" + er + " s " + s);
                            summaAmountofWorkerEachWork -= AmountofWorkerOnEachRazdel;
                        }
                    }
                }

            }
            FirstCellGrafik = workSheetGrafik.Cells[FirstCellGrafik.Row + 2, FirstCellGrafik.Column + 1];
            LastCellGrafik = workSheetGrafik.Cells[FirstCellGrafik.Row + amountOfWorkinRazdel + 1, FirstCellGrafik.Column + 1];
            Excel.Range rangeCellsGrafik = workSheetGrafik.get_Range(FirstCellGrafik, LastCellGrafik);
            rangeCellsGrafik.EntireColumn.HorizontalAlignment = Excel.Constants.xlLeft;
            workBookGrafik.SaveAs(grafikAdress);
            object misValue = System.Reflection.Missing.Value;
            Marshal.FinalReleaseComObject(rangeCellsGrafik);
            Marshal.FinalReleaseComObject(workSheetGrafik);
            workBookGrafik.Close(true, misValue, misValue);
            Console.WriteLine("Вы сохранили данные в *.xlsx файле?");
        }
        //записывает в график строки, номер, наименование работ, трудозатраты, кол-во рабочих и кол-во дней
        public void Zapisstrok(Excel.Worksheet workSheetoneSmeta, int numberofWorkers, Excel.Worksheet workSheetGrafik, ref int amountOfWorkinRazdel, ref int[] numRazdelTablExcelGrafik)
        {
            Excel.Range firstCellAfterContent = workSheetGrafik.Range["B6"];
            int indexAmountWorkinRazdel = 0, AmountofWorkerinEachWork = 0, numPozGrafik = 0, indexOfPozRazdel = 0;
            double zapasPartofDayAfterWork = 0;
            double[] trudozatratForRazdel = poRazdelyTrudozatrat.Values.ToArray();
            string[] valueNameofEachWork = nameforEachWorkinSmeta.Values.ToArray();
            double[] valueTrudozatratEachWork = chelChasforEachWork.Values.ToArray();
            Console.WriteLine("poisk.Length " + valueNameofEachWork.Length);
            for (int i = 0; i < allRazdelpoPoradky.Count; i++)
            {
                //chet++;
                int indexAmountofRowEachWorkinRazdel;
                int[] keyNumPozSmetaRazdelpoPoradky = allRazdelpoPoradky[i].Keys.ToArray();
                int[] valueNumPozWorkinRazdelpoPoradky = allRazdelpoPoradky[i].Values.ToArray();
                for (int r = 0; r < cellsAllRazdel.Count; r++)
                {
                    int numRowofFirstWorkofRazdel = cellsAllRazdel[r].Row + 1;
                    Excel.Range cellFirstWorkinRazdel = workSheetoneSmeta.Cells[numRowofFirstWorkofRazdel, cellsAllRazdel[r].Column];
                    int pozFirstWorkinRazdel = Convert.ToInt32(cellFirstWorkinRazdel.Value2);
                    if (keyNumPozSmetaRazdelpoPoradky[indexAmountWorkinRazdel] == pozFirstWorkinRazdel)
                    {
                        indexAmountofRowEachWorkinRazdel = 0;
                        string nameOfRazdel = cellsAllRazdel[r].Value.ToString();
                        Console.WriteLine("valueNumPozWorkinRazdelpoPoradky.Length " + valueNumPozWorkinRazdelpoPoradky.Length);
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkinRazdel, firstCellAfterContent.Column] = ++numPozGrafik;
                        numRazdelTablExcelGrafik[indexOfPozRazdel++] = firstCellAfterContent.Row + amountOfWorkinRazdel;
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkinRazdel, firstCellAfterContent.Column + 1] = nameOfRazdel;
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkinRazdel, firstCellAfterContent.Column + 2] = trudozatratForRazdel[r];
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkinRazdel, firstCellAfterContent.Column + 3] = numberofWorkers;
                        int daysOfEachWork = (int)(trudozatratForRazdel[r] / (numberofWorkers * 8));
                        if (trudozatratForRazdel[r] / (numberofWorkers * 8) - daysOfEachWork > 0.1)
                        {
                            daysOfEachWork += 1;
                        }
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkinRazdel, firstCellAfterContent.Column + 4] = daysOfEachWork;
                        Console.WriteLine(i + " amountOfWorkinRazdel " + amountOfWorkinRazdel + " nameOfRazdel " + nameOfRazdel + " trudozatratForRazdel " + trudozatratForRazdel[i] + " numRowStartRazdel " + numRowStartRazdel[i]);
                        do
                        {
                            if (r < cellsAllRazdel.Count - 1 && keyNumPozSmetaRazdelpoPoradky[indexAmountWorkinRazdel] >= numRowStartRazdel[r + 1])
                            {
                                break;
                            }
                            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkinRazdel + indexAmountofRowEachWorkinRazdel + 1, firstCellAfterContent.Column] = ++numPozGrafik;
                            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkinRazdel + indexAmountofRowEachWorkinRazdel + 1, firstCellAfterContent.Column + 1] = valueNameofEachWork[valueNumPozWorkinRazdelpoPoradky[indexAmountWorkinRazdel]];
                            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkinRazdel + indexAmountofRowEachWorkinRazdel + 1, firstCellAfterContent.Column + 2] = valueTrudozatratEachWork[valueNumPozWorkinRazdelpoPoradky[indexAmountWorkinRazdel]];
                            int amountOfWorkersinOneTime = 0;
                            do
                            {
                                amountOfWorkersinOneTime++;
                                if (valueTrudozatratEachWork[valueNumPozWorkinRazdelpoPoradky[indexAmountWorkinRazdel]] > 8 * numberofWorkers)
                                {
                                    AmountofWorkerinEachWork = numberofWorkers;
                                    break;
                                }
                                if (valueTrudozatratEachWork[valueNumPozWorkinRazdelpoPoradky[indexAmountWorkinRazdel]] <= 8 * amountOfWorkersinOneTime)
                                {
                                    AmountofWorkerinEachWork = amountOfWorkersinOneTime;
                                    break;
                                }
                            } while (amountOfWorkersinOneTime <= numberofWorkers);
                            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkinRazdel + indexAmountofRowEachWorkinRazdel + 1, firstCellAfterContent.Column + 3] = AmountofWorkerinEachWork;
                            daysOfEachWork = (int)(valueTrudozatratEachWork[valueNumPozWorkinRazdelpoPoradky[indexAmountWorkinRazdel]] / (AmountofWorkerinEachWork * 8) + zapasPartofDayAfterWork);
                            zapasPartofDayAfterWork = valueTrudozatratEachWork[valueNumPozWorkinRazdelpoPoradky[indexAmountWorkinRazdel]] / (AmountofWorkerinEachWork * 8) - daysOfEachWork;
                            if (zapasPartofDayAfterWork > 0.5)
                            {
                                daysOfEachWork += 1;
                            }
                            if (daysOfEachWork == 0)
                            {
                                daysOfEachWork += 1;
                            }//задать парралельное выполнение или алгоритм уменьшения человек
                            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkinRazdel + indexAmountofRowEachWorkinRazdel + 1, firstCellAfterContent.Column + 4] = daysOfEachWork;
                            indexAmountWorkinRazdel++;
                            indexAmountofRowEachWorkinRazdel++;
                            if (indexAmountWorkinRazdel == valueNumPozWorkinRazdelpoPoradky.Length) { indexAmountWorkinRazdel = 0; break; }
                        } while (indexAmountWorkinRazdel > 0);
                        amountOfWorkinRazdel += indexAmountofRowEachWorkinRazdel + 1;
                    }
                }
            }

        }

    }
}