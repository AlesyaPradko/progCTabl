using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditor.bl
{
    public class Expert : Worker
    {
        public Expert() : base()
        { }

        protected override void ProcessSmeta(List<Excel.Workbook> listAktKStoOneSmeta, Excel.Workbook copySmeta, RangeFile processingArea, string adresSmeta,int size,ref string _textError)
        {
            object misValue;
            Excel.Worksheet sheetCopySmeta = copySmeta.Sheets[1];
            Excel.Range rangeSmetaOne = sheetCopySmeta.get_Range(processingArea.FirstCell, processingArea.LastCell);
            Excel.Range keyCellNumberPozSmeta = rangeSmetaOne.Find("№ пп");
            Excel.Range keyCellConstructionWorkSmeta = rangeSmetaOne.Find("Кол.");
            if (keyCellNumberPozSmeta != null && keyCellConstructionWorkSmeta != null)
            {
                int lastRowCellsAfterDelete = 0;
                ParserExc.DeleteColumnandRow(sheetCopySmeta, rangeSmetaOne, keyCellNumberPozSmeta, adresSmeta, ref _textError, ref lastRowCellsAfterDelete);
                Excel.Range newLastCell = sheetCopySmeta.Cells[lastRowCellsAfterDelete, rangeSmetaOne.Columns.Count];
                rangeSmetaOne = sheetCopySmeta.get_Range(keyCellNumberPozSmeta, newLastCell);//уменьшение области обработки
                int insertColumnTotalScopeWork = keyCellConstructionWorkSmeta.Column + 1;
                Excel.Range firstCellNewColumn = sheetCopySmeta.Cells[keyCellNumberPozSmeta.Row, insertColumnTotalScopeWork];
                Excel.Range lastCellProcessing = sheetCopySmeta.Range[processingArea.LastCell];
                Excel.Range lastCellNewColumn = sheetCopySmeta.Cells[lastCellProcessing.Row, insertColumnTotalScopeWork];
                Excel.Range insertNewColumn = sheetCopySmeta.get_Range(firstCellNewColumn, lastCellNewColumn);
                insertNewColumn.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight);
                Dictionary<int, double> totalScopeWorkForSmeta = ParserExc.GetkeySmetaForRecord<double>(sheetCopySmeta, rangeSmetaOne, adresSmeta, ref _textError);
                Dictionary<int, string> periodTimeWorkForSmeta = ParserExc.GetkeySmetaForRecord<string>(sheetCopySmeta, rangeSmetaOne, adresSmeta, ref _textError);
                string[] valPeriodTimeWorkForSmeta = periodTimeWorkForSmeta.Values.ToArray();
                Excel.Range CellContentConstruct = rangeSmetaOne.Find("Выполнение по смете");
                if (CellContentConstruct == null)
                {
                    Excel.Range topMergeCellContentConstruct = sheetCopySmeta.Cells[keyCellNumberPozSmeta.Row, insertColumnTotalScopeWork];
                    Excel.Range bottomMergeCellContentConstruct = sheetCopySmeta.Cells[keyCellNumberPozSmeta.Row + 2, insertColumnTotalScopeWork];
                    Excel.Range mergeCellContentConstruct = sheetCopySmeta.get_Range(topMergeCellContentConstruct, bottomMergeCellContentConstruct);
                    mergeCellContentConstruct.Merge();
                    mergeCellContentConstruct.Value = "Выполнение по смете";
                    sheetCopySmeta.Cells[keyCellNumberPozSmeta.Row + 3, insertColumnTotalScopeWork] = insertColumnTotalScopeWork - keyCellNumberPozSmeta.Column + 1;
                }
                int numberLastColumnCellNote = GetColumforZapisNote(sheetCopySmeta, rangeSmetaOne);
                if (numberLastColumnCellNote == -1)
                {
                     throw new ZapredelException("Вы задали слишком малую область по ширине таблицы, задайте большую\n");
                }
                string error = null;
                string[] nameAktKS = new string[listAktKStoOneSmeta.Count];
                if (listAktKStoOneSmeta.Count != 0)
                {
                    Parallel.For(0, listAktKStoOneSmeta.Count, numKS =>
                    {
                         Excel.Worksheet workSheetAktKS = listAktKStoOneSmeta[numKS].Sheets[1];
                         Excel.Range firstAktKS = workSheetAktKS.Cells[1, 1];
                         Excel.Range lastAktKS = workSheetAktKS.Cells[rangeSmetaOne.Rows.Count + rangeSmetaOne.Row, rangeSmetaOne.Columns.Count + rangeSmetaOne.Column];
                         Excel.Range rangeAktKS = workSheetAktKS.get_Range(firstAktKS, lastAktKS);
                         Dictionary<int, double> totalScopeWorkAktKSone = new Dictionary<int, double>();
                         WorkWithAktKSExpert(workSheetAktKS, rangeAktKS, numKS, listAktKStoOneSmeta[numKS].FullName, ref error, ref nameAktKS, ref totalScopeWorkAktKSone);
                         int[] keyScopeWorkforSmeta = totalScopeWorkForSmeta.Keys.ToArray();
                         int[] keyWorkAktKS = totalScopeWorkAktKSone.Keys.ToArray();
                         for (int i = 0; i < totalScopeWorkForSmeta.Count; i++)
                         {
                             for (int j = 0; j < totalScopeWorkAktKSone.Count; j++)
                             {
                                if (keyScopeWorkforSmeta[i] == keyWorkAktKS[j])
                                {
                                    totalScopeWorkForSmeta[keyWorkAktKS[j]] += totalScopeWorkAktKSone[keyWorkAktKS[j]];
                                    periodTimeWorkForSmeta[keyWorkAktKS[j]] += nameAktKS[numKS];
                                }
                             }
                         }
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
                    RecordInFileExpert(sheetCopySmeta, rangeSmetaOne, insertColumnTotalScopeWork, numberLastColumnCellNote, totalScopeWorkForSmeta, periodTimeWorkForSmeta, adresSmeta, ref _textError);
                    RecordFormulaExpert(sheetCopySmeta, rangeSmetaOne, keyCellConstructionWorkSmeta, insertColumnTotalScopeWork);
                    FormatRecordCopySmeta(sheetCopySmeta, rangeSmetaOne, adresSmeta, size, ref _textError);
                    Excel.Range topLastColumnNote = sheetCopySmeta.Cells[rangeSmetaOne.Row, numberLastColumnCellNote + 1];
                    Excel.Range bottomLastColumnNote = sheetCopySmeta.Cells[rangeSmetaOne.Rows.Count, numberLastColumnCellNote + 1];
                    Excel.Range rangeLastColumnNote = sheetCopySmeta.get_Range(topLastColumnNote, bottomLastColumnNote);
                    rangeLastColumnNote.ColumnWidth = 50;
                    Marshal.FinalReleaseComObject(rangeSmetaOne);
                    misValue = System.Reflection.Missing.Value;
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

        //получение столбца где будет записан столбец примечание для записи в него из каких актов КС-2 взяты объемы
        private int GetColumforZapisNote(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne)
        {
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

        //метод записывает в последний столбец "Остаток" формулу разности - остатка работ для эксперта
        private void RecordFormulaExpert(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, Excel.Range keyCellVupolnSmeta, int vstavkaColumntotalScopeWork)
        {
            Excel.Range topInsertColumn = SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row, vstavkaColumntotalScopeWork + 1];
            Excel.Range bottomInsertColumn = SheetcopySmetaOne.Cells[rangeSmetaOne.Rows.Count, vstavkaColumntotalScopeWork + 1];
            Excel.Range restInsertColumn = SheetcopySmetaOne.get_Range(topInsertColumn, bottomInsertColumn);
            restInsertColumn.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight);
            Excel.Range topMergeCellContentRest = SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row, vstavkaColumntotalScopeWork + 1];
            Excel.Range bottomMergeCellContentRest = SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row + 2, vstavkaColumntotalScopeWork + 1];
            Excel.Range mergeCellContentRest = SheetcopySmetaOne.get_Range(topMergeCellContentRest, bottomMergeCellContentRest);
            mergeCellContentRest.Merge();
            mergeCellContentRest.Value = "Остаток";
            SheetcopySmetaOne.Cells[rangeSmetaOne.Row + 3, vstavkaColumntotalScopeWork + 1] = vstavkaColumntotalScopeWork - rangeSmetaOne.Column + 2;
            for (int j = rangeSmetaOne.Row + 4; j < rangeSmetaOne.Rows.Count + rangeSmetaOne.Row + 4; j++)
            {
                Excel.Range cellsVupolnSmetaColumnTabl = SheetcopySmetaOne.Cells[j, keyCellVupolnSmeta.Column];
                if (cellsVupolnSmetaColumnTabl != null && cellsVupolnSmetaColumnTabl.Value2 != null && cellsVupolnSmetaColumnTabl.Value2.ToString() != "" && !cellsVupolnSmetaColumnTabl.MergeCells)
                {
                    Excel.Range restFormula = SheetcopySmetaOne.Cells[j, vstavkaColumntotalScopeWork + 1];
                    restFormula.FormulaR1C1 = "=RC[-2]-RC[-1]";
                }
            }
        }

        //метод меняет по ссылке массив строк - наименование Актов КС-2 за определенный период и заполняет словарь
        //где ключ -номер позиции по смете из Актов КС, значение выполнение по смете
        private void WorkWithAktKSExpert(Excel.Worksheet workSheetAktKS, Excel.Range rangeAktKS, int numKS, string adresKs,ref string error, ref string[] nameAktKS, ref Dictionary<int, double> totalScopeWorkAktKSone)
        {
            try
            {
                nameAktKS[numKS] = "Акт КС-2 №";
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
                        nameAktKS[numKS] += $"{findCellNumAktKS.Value.ToString()} {monthAktKSpropis}{yearAktKS}\n";
                        totalScopeWorkAktKSone = ParserExc.GetScopeWorkAktKSone(workSheetAktKS, rangeAktKS, keyNumPozpoSmeteinAktKS, keyscopeWorkinAktKS, nameAktKS[numKS], ref error);
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
            catch(COMException ex)
            {
                error += $"{ex.Message} Проверьте чтобы в {adresKs} было верно записано устойчивое выражение [Номер документа] или [Дата составления] или  [по смете] или [за отчетный|(К|к)оличество] \n";
            }
            catch (NullvalueException ex)
            {
                error += $"{ex.parName}";
            }
        }

        //метод записывает в файл копии сметы объемы из Актов КС-2, все объемы работ по каждой позиции
        //суммируются в одном столбце, вставка столбца идет за столбцом объемы по смете  
        private void RecordInFileExpert(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, int insertColumnTotalScopeWork, int numberLastColumnCellNote, Dictionary<int, double> totalScopeWorkforSmeta, Dictionary<int, string> periodTimeWorkforSmeta, string adresSmeta,ref string _textError)
        {
            int pozSmeta;
            for (int j = rangeSmetaOne.Row + 4; j < rangeSmetaOne.Rows.Count + rangeSmetaOne.Row + 4; j++)
            {
                Excel.Range cellsNumPozColumnTabl = SheetcopySmetaOne.Cells[j, rangeSmetaOne.Column];
                if (cellsNumPozColumnTabl != null && cellsNumPozColumnTabl.Value2 != null && cellsNumPozColumnTabl.Value2.ToString() != "" && !cellsNumPozColumnTabl.MergeCells)
                {
                    try
                    {

                        pozSmeta = Convert.ToInt32(cellsNumPozColumnTabl.Value2);
                        SheetcopySmetaOne.Cells[j, insertColumnTotalScopeWork] = totalScopeWorkforSmeta[pozSmeta];
                        SheetcopySmetaOne.Cells[j, numberLastColumnCellNote] = periodTimeWorkforSmeta[pozSmeta];
                    }
                    catch (FormatException ex)
                    {                     
                        _textError+= $"{ex.Message} Вы ввели неверный формат для {adresSmeta} в строке {cellsNumPozColumnTabl.Row} в столбце {cellsNumPozColumnTabl.Column}(не должно быть [.,букв], только целые числа\n";
                    }
                }
            }
        }
    }
}
