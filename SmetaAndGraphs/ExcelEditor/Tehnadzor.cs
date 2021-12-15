using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditor.bl
{
    public class Tehnadzor : Worker
    {
        public Tehnadzor() : base()
        { }
        protected override void ProcessSmeta(List<Excel.Workbook> listAktKStoOneSmeta, Excel.Workbook copySmeta, RangeFile processingArea, string adresSmeta,int size,ref string _textError)
        {
            object misValue;
            Excel.Worksheet sheetCopySmeta = copySmeta.Sheets[1];
            Excel.Range rangeSmetaOne = sheetCopySmeta.get_Range(processingArea.FirstCell, processingArea.LastCell);
            Excel.Range keyCellNumberPozSmeta = rangeSmetaOne.Find("№ пп");
            Excel.Range keyCellConstructWorkSmeta = rangeSmetaOne.Find("Кол.");
            if (keyCellNumberPozSmeta != null && keyCellConstructWorkSmeta != null)
            {
                int nextInsertColumn = keyCellConstructWorkSmeta.Column + 1;
                int lastRowCellsAfterDelete = 0;
                ParserExc.DeleteColumnandRow(sheetCopySmeta, rangeSmetaOne, keyCellNumberPozSmeta, adresSmeta, ref _textError, ref lastRowCellsAfterDelete);
                Console.WriteLine(lastRowCellsAfterDelete);
                Excel.Range newLastCell = sheetCopySmeta.Cells[lastRowCellsAfterDelete, rangeSmetaOne.Columns.Count];
                rangeSmetaOne = sheetCopySmeta.get_Range(keyCellNumberPozSmeta, newLastCell);//уменьшение области обработки 
                List<Excel.Workbook> aktKSinOrderSort = SortAktKSforTehnadzor(listAktKStoOneSmeta, ref _textError);
                List<Dictionary<int, double>> forRecordWorkColumnInSmeta = new List<Dictionary<int, double>>();
                string[] nameAktKSRecordColumn = new string[aktKSinOrderSort.Count];
                string error = null;
                if (aktKSinOrderSort.Count != 0)
                {
                    Parallel.For(0, aktKSinOrderSort.Count, numKS =>
                    {
                        Dictionary<int, double> totalScopeWorkAktKSone = new Dictionary<int, double>();
                        Excel.Worksheet workSheetAktKS = aktKSinOrderSort[numKS].Sheets[1];
                        Excel.Range rangeAktKS = workSheetAktKS.get_Range(processingArea.FirstCell, processingArea.LastCell);
                        string nameAktKS = null;
                        WorkWithAktKSTehnadzor(workSheetAktKS, rangeAktKS, aktKSinOrderSort[numKS].FullName, ref error, ref nameAktKS, ref totalScopeWorkAktKSone);
                        forRecordWorkColumnInSmeta.Add(totalScopeWorkAktKSone);
                        nameAktKSRecordColumn[numKS] = nameAktKS;
                        Marshal.FinalReleaseComObject(rangeAktKS);
                        Marshal.FinalReleaseComObject(workSheetAktKS);
                    });
                    for (int numKS = 0; numKS < listAktKStoOneSmeta.Count; numKS++)
                    {
                        misValue = System.Reflection.Missing.Value;
                        listAktKStoOneSmeta[numKS].Close(false, misValue, misValue);
                    }
                }
                _textError += error;
                RecordFileTehnadzor(sheetCopySmeta, rangeSmetaOne, forRecordWorkColumnInSmeta, nameAktKSRecordColumn, ref nextInsertColumn, ref _textError, adresSmeta);
                FormatRecordCopySmeta(sheetCopySmeta, rangeSmetaOne, adresSmeta, size, ref _textError);
                if (aktKSinOrderSort.Count != 0)
                {
                    RecordFormulaTehnadzor(sheetCopySmeta, rangeSmetaOne, keyCellConstructWorkSmeta, nextInsertColumn);
                    ZeroMinValue(sheetCopySmeta, rangeSmetaOne, nextInsertColumn);
                }
                misValue = System.Reflection.Missing.Value;
                Marshal.FinalReleaseComObject(rangeSmetaOne);
                Marshal.FinalReleaseComObject(sheetCopySmeta);
                copySmeta.Close(true, misValue, misValue);
                Marshal.FinalReleaseComObject(copySmeta);
            }
            else
            {
                misValue = System.Reflection.Missing.Value;
                for (int numKS = 0; numKS < listAktKStoOneSmeta.Count; numKS++)
                {
                    listAktKStoOneSmeta[numKS].Close(false, misValue, misValue);
                }
                Marshal.FinalReleaseComObject(sheetCopySmeta);
                copySmeta.Close(true, misValue, misValue);
                Marshal.FinalReleaseComObject(copySmeta);
                throw new NullvalueException($" Проверьте чтобы в {adresSmeta} было верно записано устойчивое выражение [№ пп] или [Кол.]\n");
            }        
        }

        //метод возврашает строку - наименование столбца выполненных объемов работ по КС-2 за определенный период и заполняет словарь
        //где ключ -номер позиции по смете из Актов КС, значение выполнение по смете
        private void WorkWithAktKSTehnadzor(Excel.Worksheet workSheetAktKS, Excel.Range rangeAktKS, string adresKs, ref string error,ref string nameAktKS, ref Dictionary<int, double> totalScopeWorkAktKSone)
        {          
            try
            {
                nameAktKS = "Акт КС-2 №";
                RegexReg regul = new RegexReg();
                Excel.Range keyNumPozpoSmeteinAktKS = rangeAktKS.Find("по смете");
                Excel.Range keyscopeWorkinAktKS = ParserExc.FindCellOfRegul(workSheetAktKS, rangeAktKS, regul.scopeWorkInAktKS);
                if (keyNumPozpoSmeteinAktKS != null && keyscopeWorkinAktKS != null)
                {
                    Excel.Range findCellNumAktKS = rangeAktKS.Find("Номер документа");
                    Excel.Range findCellDatAktKS = rangeAktKS.Find("Дата составления");
                    if (findCellNumAktKS != null && findCellDatAktKS != null)
                    {
                        findCellNumAktKS = FindCellforNameKS(workSheetAktKS, findCellNumAktKS);
                        findCellDatAktKS = FindCellforNameKS(workSheetAktKS, findCellDatAktKS);
                        string yearAktKS = ParserExc.FindDateAktKS(regul.regexYear, findCellDatAktKS);
                        string monthAktKS = ParserExc.FindDateAktKS(regul.regexMonth, findCellDatAktKS);
                        string monthAktKSpropis = ParserExc.MonthLetter(monthAktKS);
                        nameAktKS += $" {findCellNumAktKS.Value.ToString()} {monthAktKSpropis}{yearAktKS} ";
                        totalScopeWorkAktKSone = ParserExc.GetScopeWorkAktKSone(workSheetAktKS, rangeAktKS, keyNumPozpoSmeteinAktKS, keyscopeWorkinAktKS, adresKs,ref error);
                    }
                    else
                    {
                        throw new NullvalueException($" Проверьте чтобы в {adresKs} было верно записано устойчивое выражение [Номер документа] или [Дата составления]\n");
                    }
                }
                else
                {
                    throw new NullvalueException($"Проверьте чтобы в {adresKs} было верно записано устойчивое выражение [по смете] или [за отчетный|(К|к)оличество]\n");
                }
            }
            catch (COMException ex)
            {
                error += $"{ex.Message} Проверьте чтобы в {adresKs} было верно записано устойчивое выражение [Номер документа] или [Дата составления] или  [по смете] или [за отчетный|(К|к)оличество] \n";
            }
            catch (NullvalueException ex)
            {
                error += $"{ex.parName}";
            }
        }

        //метод записывает в файл копии сметы объемы из Актов КС-2, каждый месяц в новый столбец,
        //вставка столбцов идет за столбцом объемы по смете  
        private void RecordFileTehnadzor(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, List<Dictionary<int, double>> forRecordWorkColumninSmeta, string[] nameAktKSRecordColumn, ref int nextInsertColumn, ref string _textError, string adresSmeta)
        {
            //Console.WriteLine(" RecordFileTehnadzor");
            int pozSmeta;
            for (int i = 0; i < forRecordWorkColumninSmeta.Count; i++)
            {
                ICollection keyCollScopeWorkAktKSone = forRecordWorkColumninSmeta[i].Keys;
                for (int j = rangeSmetaOne.Row; j < rangeSmetaOne.Rows.Count + rangeSmetaOne.Row + 1; j++)
                {
                    Excel.Range cellsNextColumnTablInsert = SheetcopySmetaOne.Cells[j, nextInsertColumn];
                    cellsNextColumnTablInsert.Insert(XlInsertShiftDirection.xlShiftToRight);
                    if (j > rangeSmetaOne.Row + 4)
                    {
                        Excel.Range cellsNumPozColumnTabl = SheetcopySmetaOne.Cells[j, rangeSmetaOne.Column];
                        if (cellsNumPozColumnTabl != null && cellsNumPozColumnTabl.Value2 != null && cellsNumPozColumnTabl.Value2.ToString() != "" && !cellsNumPozColumnTabl.MergeCells)
                        {
                            try
                            {
                                pozSmeta = Convert.ToInt32(cellsNumPozColumnTabl.Value2);
                                foreach (int pozSmetaAktKS in keyCollScopeWorkAktKSone)
                                {
                                    if (pozSmeta == pozSmetaAktKS)
                                    {
                                        SheetcopySmetaOne.Cells[j, nextInsertColumn] = forRecordWorkColumninSmeta[i][pozSmetaAktKS];
                                    }
                                }
                            }
                            catch (FormatException ex)
                            {
                                 _textError+= $"{ex.Message} Вы ввели неверный формат для {adresSmeta} в строке {cellsNumPozColumnTabl.Row} в столбце {cellsNumPozColumnTabl.Column}(не должно быть [.,букв], только целые числа\n";
                            }
                        }
                    }
                }
                Excel.Range topCellmergeCellNameAktKS = SheetcopySmetaOne.Cells[rangeSmetaOne.Row, nextInsertColumn];
                Excel.Range bottomCellmergeCellNameAktKS = SheetcopySmetaOne.Cells[rangeSmetaOne.Row + 2, nextInsertColumn];
                Excel.Range mergeCellNameAktKS = SheetcopySmetaOne.get_Range(topCellmergeCellNameAktKS, bottomCellmergeCellNameAktKS);
                mergeCellNameAktKS.Merge();
                mergeCellNameAktKS.Value = nameAktKSRecordColumn[i];
                SheetcopySmetaOne.Cells[rangeSmetaOne.Row + 3, nextInsertColumn] = nextInsertColumn - rangeSmetaOne.Column + 1;
                nextInsertColumn += 1;
            }
        }
        //метод записывает в последний столбец "Остаток" формулу разности - остатка работ для технадзора
        private void RecordFormulaTehnadzor(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, Excel.Range keyCellVupolnSmeta, int nextInsertColumn)
        {
            //Console.WriteLine(" RecordFormulaTehnadzor");
            Excel.Range topInsertColumn = SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row, nextInsertColumn];
            Excel.Range bottomInsertColumn = SheetcopySmetaOne.Cells[rangeSmetaOne.Rows.Count, nextInsertColumn];
            Excel.Range restInsertColumn = SheetcopySmetaOne.get_Range(topInsertColumn, bottomInsertColumn);
            restInsertColumn.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight);
            Excel.Range topMergeCellContentRest = SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row, nextInsertColumn];
            Excel.Range bottomMergeCellContentRest = SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row + 2, nextInsertColumn];
            Excel.Range mergeCellContentRest = SheetcopySmetaOne.get_Range(topMergeCellContentRest, bottomMergeCellContentRest);
            mergeCellContentRest.Merge();
            mergeCellContentRest.Value = "Остаток";
            mergeCellContentRest.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
            mergeCellContentRest.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
            mergeCellContentRest.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
            mergeCellContentRest.EntireColumn.AutoFit();
            SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row + 3, nextInsertColumn] = nextInsertColumn - rangeSmetaOne.Column + 1;
            Excel.Range cellContentNumRest = SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row + 3, nextInsertColumn];
            cellContentNumRest.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
            cellContentNumRest.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
            cellContentNumRest.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
            cellContentNumRest.EntireColumn.AutoFit();
            int amountColumnAktKS = nextInsertColumn - keyCellVupolnSmeta.Column;
            if (amountColumnAktKS > 1)
            {
                for (int j = rangeSmetaOne.Row + 4; j < rangeSmetaOne.Rows.Count + rangeSmetaOne.Row; j++)
                {
                    Excel.Range restFormula = SheetcopySmetaOne.Cells[j, nextInsertColumn];
                    restFormula.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                    restFormula.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
                    restFormula.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
                    restFormula.EntireColumn.AutoFit();
                    Excel.Range cellsVupolnSmetaColumnTabl = SheetcopySmetaOne.Cells[j, keyCellVupolnSmeta.Column];
                    if (cellsVupolnSmetaColumnTabl != null && cellsVupolnSmetaColumnTabl.Value2 != null && cellsVupolnSmetaColumnTabl.Value2.ToString() != "" && !cellsVupolnSmetaColumnTabl.MergeCells)
                    {
                        switch (amountColumnAktKS)
                        {
                            case 2:
                                restFormula.FormulaR1C1 = "=RC[-2]-RC[-1]"; break;
                            case 3:
                                restFormula.FormulaR1C1 = "=RC[-3]-RC[-2]-RC[-1]"; break;
                            case 4:
                                restFormula.FormulaR1C1 = "=RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 5:
                                restFormula.FormulaR1C1 = "=RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 6:
                                restFormula.FormulaR1C1 = "=RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 7:
                                restFormula.FormulaR1C1 = "=RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 8:
                                restFormula.FormulaR1C1 = "=RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 9:
                                restFormula.FormulaR1C1 = "=RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 10:
                                restFormula.FormulaR1C1 = "=RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 11:
                                restFormula.FormulaR1C1 = "=RC[-11]-RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 12:
                                restFormula.FormulaR1C1 = "=RC[-12]-RC[-11]-RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 13:
                                restFormula.FormulaR1C1 = "=RC[-13]-RC[-12]-RC[-11]-RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            default: Console.WriteLine("Сводная таблица ведется до года, начните новую"); break;
                        }
                        restFormula.EntireColumn.AutoFit();
                    }
                }
            }
        }

        //метод возвращает отсортированный лист книг Иксель, акты КС для сметы
        private List<Excel.Workbook> SortAktKSforTehnadzor(List<Excel.Workbook> listAktKStoOneSmeta, ref string _textError)
        {
            List<Excel.Workbook> aktKSinOrderSort = new List<Excel.Workbook>();
            try
            {//Console.WriteLine(" SortAktKSforTehnadzor");
                if (listAktKStoOneSmeta != null)
                {
                    RegexReg reg = new RegexReg();
                    Dictionary<string, int> nomercifraList = new Dictionary<string, int>();
                    for (int i = 0; i < listAktKStoOneSmeta.Count; i++)
                    {
                        string monthAktKS = null, yearAktKS = null, dateVal = null;
                        int yearVal = 0, monthVal = 0;
                        string numerKS = listAktKStoOneSmeta[i].FullName;
                        MatchCollection mathesDate = reg.regexData.Matches(numerKS);
                        if (mathesDate.Count > 0)
                        {
                            foreach (Match date in mathesDate)
                            {
                                dateVal = date.Value;
                            }
                            MatchCollection mathesMonth = reg.regexMonth.Matches(dateVal);
                            MatchCollection mathesYear = reg.regexYear.Matches(dateVal);
                            if (mathesMonth.Count > 0 && mathesYear.Count > 0)
                            {
                                foreach (Match month in mathesMonth)
                                {
                                    monthAktKS = month.Value;
                                    monthAktKS = monthAktKS.Remove(monthAktKS.Length - 1, 1);
                                    monthVal = Convert.ToInt32(monthAktKS);
                                }
                                foreach (Match year in mathesYear)
                                {
                                    yearAktKS = year.Value;
                                    yearAktKS = yearAktKS.Remove(0, 1);
                                    yearVal = Convert.ToInt32(yearAktKS);
                                }

                                int nomercifra = yearVal * 100 + monthVal;
                                //Console.WriteLine(nomercifra + "mon " + monthAktKS + " year " + yearAktKS + " = " + dateVal);
                                nomercifraList.Add(dateVal, nomercifra);
                            }
                        }
                    }
                    if (nomercifraList.Count > 0)
                    {
                        int[] valueNomerCifra = nomercifraList.Values.ToArray();
                        string[] keyNomerCifra = nomercifraList.Keys.ToArray();
                        for (int i = 1; i < valueNomerCifra.Length; i++)
                        {
                            for (int j = i; j > 0; j--)
                            {
                                if (valueNomerCifra[j] < valueNomerCifra[j - 1])
                                {
                                    int temp = valueNomerCifra[j - 1];
                                    valueNomerCifra[j - 1] = valueNomerCifra[j];
                                    valueNomerCifra[j] = temp;
                                    string test = keyNomerCifra[j - 1];
                                    keyNomerCifra[j - 1] = keyNomerCifra[j];
                                    keyNomerCifra[j] = test;
                                }
                                else break;
                            }
                        }

                        for (int j = 0; j < keyNomerCifra.Length; j++)
                        {
                            for (int i = 0; i < listAktKStoOneSmeta.Count; i++)
                            {
                                string numerKS = listAktKStoOneSmeta[i].FullName;
                                if (numerKS.Contains(keyNomerCifra[j]))
                                {
                                    aktKSinOrderSort.Add(listAktKStoOneSmeta[i]);
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (listAktKStoOneSmeta.Count > 0)
                        {
                            _textError+= "Так как Вы не указали в названии Актов КС-2 дату в форме мм.гггг, то в итоговой таблице Акты не будут отсортированы по порядку\n";
                            aktKSinOrderSort = listAktKStoOneSmeta;
                        }
                    }
                }

            }
            catch (NullReferenceException ex)
            {
                _textError += $"{ex.Message} В названии сметы отсутствует символ № перед номером сметы\n";
            }
            return aktKSinOrderSort;
        }
    }
}
