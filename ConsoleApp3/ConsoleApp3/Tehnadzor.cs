using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp3
{

    public class Tehnadzor : Worker
    {
        public Tehnadzor() : base()
        { }
        protected override void ProcessSmeta(List<Excel.Workbook> listAktKStoOneSmeta, Excel.Worksheet sheetCopySmeta, RangeFile processingArea, string adresSmeta)
        {
            //Console.WriteLine("ProcessSmeta tehnadzor");
            try
            {     
                Excel.Range rangeSmetaOne = sheetCopySmeta.get_Range(processingArea.FirstCell, processingArea.LastCell);
                Excel.Range keyCellNumberPozSmeta = rangeSmetaOne.Find("№ пп");
                Excel.Range keyCellConstructionWorkSmeta = rangeSmetaOne.Find("Кол.");
                int nextInsertColumn = keyCellConstructionWorkSmeta.Column + 1;
                int lastRowCellsAfterDelete = 0;
                ParserExc.DeleteColumnandRow(sheetCopySmeta, rangeSmetaOne, keyCellNumberPozSmeta, adresSmeta, ref lastRowCellsAfterDelete);
                Console.WriteLine(lastRowCellsAfterDelete);
                Excel.Range newLastCell = sheetCopySmeta.Cells[lastRowCellsAfterDelete, rangeSmetaOne.Columns.Count];
                rangeSmetaOne = sheetCopySmeta.get_Range(keyCellNumberPozSmeta, newLastCell);//уменьшение области обработки 
                List<Excel.Workbook> aktKSinOrderSort = SortAktKSforTehnadzor(listAktKStoOneSmeta);
                List<Dictionary<int, double>> forRecordWorkColumninSmeta = new List<Dictionary<int, double>>();
                string[] nameAktKSRecordColumn = new string[aktKSinOrderSort.Count];
                for (int numKS = 0; numKS < aktKSinOrderSort.Count; numKS++)
                {
                    Dictionary<int, double> totalScopeWorkAktKSone=new Dictionary<int, double> ();
                    Console.WriteLine(aktKSinOrderSort[numKS].FullName);
                    Excel.Worksheet workSheetAktKS = aktKSinOrderSort[numKS].Sheets[1];
                    Excel.Range rangeAktKS = workSheetAktKS.get_Range(processingArea.FirstCell, processingArea.LastCell);
                    string nameAktKS=null;
                    WorkWithAktKSTehnadzor(workSheetAktKS, rangeAktKS, _adresAktKS[numKS],ref nameAktKS, ref totalScopeWorkAktKSone);
                    forRecordWorkColumninSmeta.Add(totalScopeWorkAktKSone);
                    nameAktKSRecordColumn[numKS] = nameAktKS;        
                    Marshal.FinalReleaseComObject(rangeAktKS);
                    Marshal.FinalReleaseComObject(workSheetAktKS); 
                }
                ZapisinfileTehnadzor(sheetCopySmeta, rangeSmetaOne, forRecordWorkColumninSmeta,nameAktKSRecordColumn, ref nextInsertColumn, adresSmeta);
                FormatZapisinCopySmeta(sheetCopySmeta, rangeSmetaOne, adresSmeta);
                ZapisFormulaTehnadzor(sheetCopySmeta, rangeSmetaOne, keyCellConstructionWorkSmeta, nextInsertColumn);
                ObnulenieMinValue(sheetCopySmeta, rangeSmetaOne, nextInsertColumn);           
            }
            catch (NullReferenceException ex)
            {
                Console.WriteLine($"{ex.Message} Проверьте чтобы в {adresSmeta} было верно записано устойчивое выражение [№ пп] или [Кол.]");
            }     
        }

        //метод возврашает строку - наименование столбца выполненных объемов работ по КС-2 за определенный период и заполняет словарь
        //где ключ -номер позиции по смете из Актов КС, значение выполнение по смете
        private void WorkWithAktKSTehnadzor(Excel.Worksheet workSheetAktKS, Excel.Range rangeAktKS, string adresKs, ref string nameAktKS,ref Dictionary<int, double> totalScopeWorkAktKSone)
        {
            //Console.WriteLine("WorkWithAktKSTehnadzor");
            nameAktKS = "Акт КС-2 №";
            try
            {
                RegexReg regul = new RegexReg();
                Excel.Range keyNumPozpoSmeteinAktKS = rangeAktKS.Find("по смете");
                Excel.Range keyscopeWorkinAktKS = ParserExc.FindCellofRegul(workSheetAktKS, rangeAktKS, regul.scopeWorkinAktKS);
                if (keyNumPozpoSmeteinAktKS != null && keyscopeWorkinAktKS != null)
                {
                    Excel.Range findCellNumAktKS = rangeAktKS.Find("Номер документа");
                    findCellNumAktKS = FindCellforNameKS(workSheetAktKS, findCellNumAktKS);
                    Excel.Range findCellDatAktKS = rangeAktKS.Find("Дата составления");
                    findCellDatAktKS = FindCellforNameKS(workSheetAktKS, findCellDatAktKS);
                    string yearAktKS = ParserExc.FinddateAktKS(regul.regexyear, findCellDatAktKS);
                    string monthAktKS = ParserExc.FinddateAktKS(regul.regexmonth, findCellDatAktKS);
                    string monthAktKSpropis = ParserExc.Monthpropis(monthAktKS);
                    nameAktKS += $" {findCellNumAktKS.Value.ToString()} {monthAktKSpropis}{yearAktKS} ";
                    totalScopeWorkAktKSone = ParserExc.GetScopeWorkAktKSone(workSheetAktKS, rangeAktKS, keyNumPozpoSmeteinAktKS, keyscopeWorkinAktKS, adresKs);
                }
                else
                {
                    Console.WriteLine($"Проверьте чтобы в {adresKs} было верно записано устойчивое выражение [по смете] или [за отчетный|(К|к)оличество]");
                }
            }
            catch (NullReferenceException ex)
            {
                Console.WriteLine($"{ex.Message} Проверьте чтобы в {adresKs} было верно записано устойчивое выражение Номер документа или Дата составления");
            }
        }

        //метод записывает в файл копии сметы объемы из Актов КС-2, каждый месяц в новый столбец,
        //вставка столбцов идет за столбцом объемы по смете  
        private void ZapisinfileTehnadzor(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, List<Dictionary<int, double>> forRecordWorkColumninSmeta,string[] nameAktKSRecordColumn,ref int nextInsertColumn, string adresSmeta)
        {
            //Console.WriteLine(" ZapisinfileTehnadzor");
            int pozSmeta;
            for (int i = 0; i < forRecordWorkColumninSmeta.Count ; i++)
            {
                ICollection keyCollScopeWorkAktKSone = forRecordWorkColumninSmeta[i].Keys;
                for (int j = rangeSmetaOne.Row; j < rangeSmetaOne.Rows.Count + rangeSmetaOne.Row; j++)
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
                                Console.WriteLine($"{ex.Message} Вы ввели неверный формат для {adresSmeta} в строке {cellsNumPozColumnTabl.Row} в столбце {cellsNumPozColumnTabl.Column}(не должно быть [.,букв], только целые числа ");
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
        private void ZapisFormulaTehnadzor(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, Excel.Range keyCellVupolnSmeta, int nextInsertColumn)
        {
            //Console.WriteLine(" ZapisFormulaTehnadzor");
            Excel.Range topInsertColumn = SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row, nextInsertColumn];
            Excel.Range bottomInsertColumn = SheetcopySmetaOne.Cells[rangeSmetaOne.Rows.Count, nextInsertColumn];
            Excel.Range ostatokInsertColumn = SheetcopySmetaOne.get_Range(topInsertColumn, bottomInsertColumn);
            ostatokInsertColumn.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight);
            Excel.Range topCellmergeCellContentOstatok = SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row, nextInsertColumn];
            Excel.Range bottomCellmergeCellContentOstatok = SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row + 2, nextInsertColumn];
            Excel.Range mergeCellContentOstatok = SheetcopySmetaOne.get_Range(topCellmergeCellContentOstatok, bottomCellmergeCellContentOstatok);
            mergeCellContentOstatok.Merge();
            mergeCellContentOstatok.Value = "Остаток";
            mergeCellContentOstatok.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
            mergeCellContentOstatok.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
            mergeCellContentOstatok.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
            mergeCellContentOstatok.EntireColumn.AutoFit();
            SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row + 3, nextInsertColumn] = nextInsertColumn - rangeSmetaOne.Column + 1;
            Excel.Range cellContentnumOstatok = SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row + 3, nextInsertColumn];
            cellContentnumOstatok.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
            cellContentnumOstatok.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
            cellContentnumOstatok.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
            cellContentnumOstatok.EntireColumn.AutoFit();
            int KolColumnAktKS = nextInsertColumn - keyCellVupolnSmeta.Column;
            if (KolColumnAktKS > 1)
            {
                for (int j = rangeSmetaOne.Row + 4; j < rangeSmetaOne.Rows.Count + rangeSmetaOne.Row; j++)
                {
                    Excel.Range ostatocFormula = SheetcopySmetaOne.Cells[j, nextInsertColumn];
                    ostatocFormula.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                    ostatocFormula.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
                    ostatocFormula.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
                    ostatocFormula.EntireColumn.AutoFit();
                    Excel.Range cellsVupolnSmetaColumnTabl = SheetcopySmetaOne.Cells[j, keyCellVupolnSmeta.Column];
                    if (cellsVupolnSmetaColumnTabl != null && cellsVupolnSmetaColumnTabl.Value2 != null && cellsVupolnSmetaColumnTabl.Value2.ToString() != "" && !cellsVupolnSmetaColumnTabl.MergeCells)
                    {
                        switch (KolColumnAktKS)
                        {
                            case 2:
                                ostatocFormula.FormulaR1C1 = "=RC[-2]-RC[-1]"; break;
                            case 3:
                                ostatocFormula.FormulaR1C1 = "=RC[-3]-RC[-2]-RC[-1]"; break;
                            case 4:
                                ostatocFormula.FormulaR1C1 = "=RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 5:
                                ostatocFormula.FormulaR1C1 = "=RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 6:
                                ostatocFormula.FormulaR1C1 = "=RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 7:
                                ostatocFormula.FormulaR1C1 = "=RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 8:
                                ostatocFormula.FormulaR1C1 = "=RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 9:
                                ostatocFormula.FormulaR1C1 = "=RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 10:
                                ostatocFormula.FormulaR1C1 = "=RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 11:
                                ostatocFormula.FormulaR1C1 = "=RC[-11]-RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 12:
                                ostatocFormula.FormulaR1C1 = "=RC[-12]-RC[-11]-RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            case 13:
                                ostatocFormula.FormulaR1C1 = "=RC[-13]-RC[-12]-RC[-11]-RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                            default: Console.WriteLine("Сводная таблица ведется до года, начните новую"); break;
                        }
                        ostatocFormula.EntireColumn.AutoFit();
                    }
                }
            }
        }

        //метод возвращает отсортированный лист книг Иксель, акты КС для сметы
        private List<Excel.Workbook> SortAktKSforTehnadzor(List<Excel.Workbook>  listAktKStoOneSmeta)
        {
            //Console.WriteLine(" SortAktKSforTehnadzor");
            List<Excel.Workbook> aktKSinOrderSort = new List<Excel.Workbook>();           
            List<int> nomercifraList = new List<int>();
            for (int i = 0; i < listAktKStoOneSmeta.Count; i++)
            {
                string nomerAktKS = null; ; ;
                string numerKS = listAktKStoOneSmeta[i].FullName;
                for (int j = numerKS.Length - 8; j < numerKS.Length - 5; j++)
                {
                    if (numerKS[j] >= '0' && numerKS[j] <= '9')
                    {
                        nomerAktKS += numerKS[j];
                    }
                }
                int nomercifra = Convert.ToInt32(nomerAktKS);
                nomercifraList.Add(nomercifra);
            }
            for (int i = 1; i < nomercifraList.Count; i++)
            {
                for (int j = i; j > 0; j--)
                {
                    if (nomercifraList[j] < nomercifraList[j - 1])
                    {
                        int temp = nomercifraList[j - 1];
                        nomercifraList[j - 1] = nomercifraList[j];
                        nomercifraList[j] = temp;
                    }
                    else break;
                }
            }

            for (int j = 0; j < nomercifraList.Count; j++)
            {
                string cifrapropis = nomercifraList[j].ToString();
                for (int i = 0; i < listAktKStoOneSmeta.Count; i++)
                {
                    int countcifr = 0;
                    string numerKS = listAktKStoOneSmeta[i].FullName;
                    if (numerKS.Contains(cifrapropis))
                    {
                        for (int v = 0; v < numerKS.Length; v++)
                        {
                            if (numerKS[v] >= '0' && numerKS[v] <= '9')
                            {
                                countcifr++;
                            }
                        }
                        if (countcifr == cifrapropis.Length)
                        {
                            aktKSinOrderSort.Add(listAktKStoOneSmeta[i]);
                            break;
                        }
                    }
                }
            }
            return aktKSinOrderSort;
        }
    }
}