using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp3
{
    public class Expert : Worker
    {
        private Dictionary<int, double> totalScopeWorkforSmeta;
        private Dictionary<int, string> periodTimeWorkforSmeta;
        private string[] nameAktKS;
        private Mutex mutexObj = new Mutex();
        public Expert() : base()
        { }

        protected override void ProcessSmeta(RangeFile oblastobrabotki)
        {
            //Console.WriteLine("ProcessSmeta expert");
            mutexObj.WaitOne();
            Console.WriteLine(Task.CurrentId + " получил мютех");
            int numSmeta = (int)Task.CurrentId - 1;
            try
            {
                Excel.Worksheet SheetcopySmetaOne = containCopySmeta[numSmeta].Sheets[1];
                Excel.Range rangeSmetaOne = SheetcopySmetaOne.get_Range(oblastobrabotki.FirstCell, oblastobrabotki.LastCell);
                Excel.Range keyCellNomerpozSmeta = rangeSmetaOne.Find("№ пп");
                Excel.Range keyCellVupolnSmeta = rangeSmetaOne.Find("Кол.");
                int lastRowCellsafterDelete = 0;
                ParserExc.DeleteColumnandRow(SheetcopySmetaOne, rangeSmetaOne, keyCellNomerpozSmeta, AdresSmeta[numSmeta], ref lastRowCellsafterDelete);
                Excel.Range newLastCell = SheetcopySmetaOne.Cells[lastRowCellsafterDelete, rangeSmetaOne.Columns.Count];
                rangeSmetaOne = SheetcopySmetaOne.get_Range(keyCellNomerpozSmeta, newLastCell);//уменьшение области обработки
                int vstavkaColumntotalScopeWork = keyCellVupolnSmeta.Column + 1;
                Excel.Range firstCellNewColumn = SheetcopySmetaOne.Cells[keyCellNomerpozSmeta.Row, vstavkaColumntotalScopeWork];
                Excel.Range lastCellobrabotki = SheetcopySmetaOne.Range[oblastobrabotki.LastCell];
                Excel.Range lastCellNewColumn = SheetcopySmetaOne.Cells[lastCellobrabotki.Row, vstavkaColumntotalScopeWork];
                Excel.Range insertNewColumn = SheetcopySmetaOne.get_Range(firstCellNewColumn, lastCellNewColumn);
                insertNewColumn.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight);
                totalScopeWorkforSmeta = ParserExc.GetkeySmetaForZapis<double>(SheetcopySmetaOne, rangeSmetaOne, AdresSmeta[numSmeta]);
                periodTimeWorkforSmeta = ParserExc.GetkeySmetaForZapis<string>(SheetcopySmetaOne, rangeSmetaOne, AdresSmeta[numSmeta]);
                string[] valperiodTimeWorkforSmeta = periodTimeWorkforSmeta.Values.ToArray();
                Excel.Range topCellmergeCellContentVupoln = SheetcopySmetaOne.Cells[keyCellNomerpozSmeta.Row, vstavkaColumntotalScopeWork];
                Excel.Range bottomCellmergeCellContentVupoln = SheetcopySmetaOne.Cells[keyCellNomerpozSmeta.Row + 2, vstavkaColumntotalScopeWork];
                Excel.Range mergeCellContentVupoln = SheetcopySmetaOne.get_Range(topCellmergeCellContentVupoln, bottomCellmergeCellContentVupoln);
                mergeCellContentVupoln.Merge();
                mergeCellContentVupoln.Value = "Выполнение по смете";
                SheetcopySmetaOne.Cells[keyCellNomerpozSmeta.Row + 3, vstavkaColumntotalScopeWork] = vstavkaColumntotalScopeWork - keyCellNomerpozSmeta.Column + 1;
                int numLastColumnCellNote = ParserExc.GetColumforZapisNote(SheetcopySmetaOne, rangeSmetaOne);
                Console.WriteLine("numLastColumnCellNote" + numLastColumnCellNote);
                if (numLastColumnCellNote == -1)
                {
                    throw new ZapredelException("Вы задали слишком малую область по ширине таблицы, задайте большую");
                }
                nameAktKS = new string[containPapkaKS.Count];
                int curNumKS = 0;
                for (int v = 0; v < aktAllKSforOneSmeta[adresSmeta[numSmeta]].Count; v++)
                {
                    for (int numKS = 0; numKS < containPapkaKS.Count; numKS++)
                    {
                        if (adresAktKS[numKS] != aktAllKSforOneSmeta[adresSmeta[numSmeta]][v]) continue;
                        else
                        {
                            Excel.Worksheet workSheetAktKS = containPapkaKS[numKS].Sheets[1];
                            Excel.Range firstAktKS = workSheetAktKS.Cells[1, 1];
                            Excel.Range lastAktKS = workSheetAktKS.Cells[rangeSmetaOne.Rows.Count + rangeSmetaOne.Row, rangeSmetaOne.Columns.Count + rangeSmetaOne.Column];
                            Excel.Range rangeAktKS = workSheetAktKS.get_Range(firstAktKS, lastAktKS);
                            WorkWithAktKSExpert(workSheetAktKS, rangeAktKS, numKS, adresAktKS[numKS], ref nameAktKS);
                            ZapisinfileExpert(SheetcopySmetaOne, rangeSmetaOne, vstavkaColumntotalScopeWork, numLastColumnCellNote, valperiodTimeWorkforSmeta, numKS, numSmeta);
                            curNumKS = numKS;
                            Marshal.FinalReleaseComObject(rangeAktKS);
                            Marshal.FinalReleaseComObject(workSheetAktKS);
                        }
                    }
                }
                ZapisFormulaExpert(SheetcopySmetaOne, rangeSmetaOne, keyCellVupolnSmeta, vstavkaColumntotalScopeWork);
                FormatZapisinCopySmeta(SheetcopySmetaOne, rangeSmetaOne, numSmeta);
                Excel.Range topLastColumnNote = SheetcopySmetaOne.Cells[rangeSmetaOne.Row, numLastColumnCellNote + 1];
                Excel.Range bottomLastColumnNote = SheetcopySmetaOne.Cells[rangeSmetaOne.Rows.Count, numLastColumnCellNote + 1];
                Excel.Range rangeLastColumnNote = SheetcopySmetaOne.get_Range(topLastColumnNote, bottomLastColumnNote);
                rangeLastColumnNote.ColumnWidth = 50;
                Zakrutie(SheetcopySmetaOne, rangeSmetaOne, numSmeta);
            }
            catch (ZapredelException exc)
            {
                Console.WriteLine(exc.parName);
            }
            catch (NullReferenceException ex)
            {
                Console.WriteLine($"{ex.Message} Проверьте чтобы в {AdresSmeta[numSmeta]} было верно записано устойчивое выражение [№ пп] или [Кол.]");
            }
            Console.WriteLine(Task.CurrentId + "освобождает");
            mutexObj.ReleaseMutex();
        }

        //метод записывает в последний столбец "Остаток" формулу разности - остатка работ для эксперта
        private void ZapisFormulaExpert(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, Excel.Range keyCellVupolnSmeta, int vstavkaColumntotalScopeWork)
        {
            //Console.WriteLine("ZapisFormulaExpert");
            Excel.Range topInsertColumn = SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row, vstavkaColumntotalScopeWork + 1];
            Excel.Range bottomInsertColumn = SheetcopySmetaOne.Cells[rangeSmetaOne.Rows.Count, vstavkaColumntotalScopeWork + 1];
            Excel.Range ostatokInsertColumn = SheetcopySmetaOne.get_Range(topInsertColumn, bottomInsertColumn);
            ostatokInsertColumn.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight);
            Excel.Range topCellmergeCellContentOstatok = SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row, vstavkaColumntotalScopeWork + 1];
            Excel.Range bottomCellmergeCellContentOstatok = SheetcopySmetaOne.Cells[keyCellVupolnSmeta.Row + 2, vstavkaColumntotalScopeWork + 1];
            Excel.Range mergeCellContentOstatok = SheetcopySmetaOne.get_Range(topCellmergeCellContentOstatok, bottomCellmergeCellContentOstatok);
            mergeCellContentOstatok.Merge();
            mergeCellContentOstatok.Value = "Остаток";
            SheetcopySmetaOne.Cells[rangeSmetaOne.Row + 3, vstavkaColumntotalScopeWork + 1] = vstavkaColumntotalScopeWork - rangeSmetaOne.Column + 2;
            for (int j = rangeSmetaOne.Row + 4; j < rangeSmetaOne.Rows.Count + rangeSmetaOne.Row; j++)
            {
                Excel.Range cellsVupolnSmetaColumnTabl = SheetcopySmetaOne.Cells[j, keyCellVupolnSmeta.Column];
                if (cellsVupolnSmetaColumnTabl != null && cellsVupolnSmetaColumnTabl.Value2 != null && cellsVupolnSmetaColumnTabl.Value2.ToString() != "" && !cellsVupolnSmetaColumnTabl.MergeCells)
                {
                    Excel.Range ostatocFormula = SheetcopySmetaOne.Cells[j, vstavkaColumntotalScopeWork + 1];
                    ostatocFormula.FormulaR1C1 = "=RC[-2]-RC[-1]";
                }
            }
        }

        //метод меняет по ссылке массив строк - наименование Актов КС-2 за определенный период и заполняет словарь
        //где ключ -номер позиции по смете из Актов КС, значение выполнение по смете
        private void WorkWithAktKSExpert(Excel.Worksheet workSheetAktKS, Excel.Range rangeAktKS, int numKS, string adresKs, ref string[] nameAktKS)
        {
            //Console.WriteLine(" WorkWithAktKSExpert");
            try
            {
                nameAktKS[numKS] = "Акт КС-2 №";
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
                    nameAktKS[numKS] += $"{findCellNumAktKS.Value.ToString()} {monthAktKSpropis}{yearAktKS}\n";
                    totalScopeWorkAktKSone = ParserExc.GetScopeWorkAktKSone(workSheetAktKS, rangeAktKS, keyNumPozpoSmeteinAktKS, keyscopeWorkinAktKS, nameAktKS[numKS]);
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

        //метод записывает в файл копии сметы объемы из Актов КС-2, все объемы работ по каждой позиции
        //суммируются в одном столбце, вставка столбца идет за столбцом объемы по смете  
        private void ZapisinfileExpert(Excel.Worksheet SheetcopySmetaOne, Excel.Range rangeSmetaOne, int vstavkaColumntotalScopeWork, int numLastColumnCellNote, string[] valperiodTimeWorkforSmeta, int numKS, int numSmeta)
        {
            //Console.WriteLine(" ZapisinfileExpert");
            int[] keytotalScopeWorkforSmeta = totalScopeWorkforSmeta.Keys.ToArray();
            double[] valtotalScopeWorkforSmeta = totalScopeWorkforSmeta.Values.ToArray();
            ICollection keyCollScopeWorkAktKSone = totalScopeWorkAktKSone.Keys;
            bool indexWasFound;
            int pozSmeta;
            for (int j = rangeSmetaOne.Row + 4; j < rangeSmetaOne.Rows.Count + rangeSmetaOne.Row; j++)
            {
                indexWasFound = false;
                int indexSmeta = 0;
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
                                indexWasFound = true;
                                indexSmeta = Array.IndexOf(keytotalScopeWorkforSmeta, pozSmetaAktKS);
                                valtotalScopeWorkforSmeta[indexSmeta] += totalScopeWorkAktKSone[pozSmetaAktKS];
                                totalScopeWorkforSmeta[pozSmetaAktKS] = valtotalScopeWorkforSmeta[indexSmeta];
                                SheetcopySmetaOne.Cells[j, vstavkaColumntotalScopeWork] = totalScopeWorkforSmeta[pozSmetaAktKS];
                            }
                        }
                        if (indexWasFound)
                        {
                            valperiodTimeWorkforSmeta[indexSmeta] += $"{nameAktKS[numKS]} ";
                            SheetcopySmetaOne.Cells[j, numLastColumnCellNote] = valperiodTimeWorkforSmeta[indexSmeta];
                        }
                    }
                    catch (FormatException ex)
                    {
                        Console.WriteLine($"{ex.Message} Вы ввели неверный формат для {AdresSmeta[numSmeta]} в строке {cellsNumPozColumnTabl.Row} в столбце {cellsNumPozColumnTabl.Column}(не должно быть [.,букв], только целые числа ");
                    }
                }
            }
        }
    }
}