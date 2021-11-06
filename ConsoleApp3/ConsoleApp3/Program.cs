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
public enum ChangeMod { expert=49, tehnadzor=50 };

namespace ConsoleApp3
{
   public class ZapredelException:Exception
    {
        public string parName;
        public ZapredelException(string s)
        { 
            parName = s;
        }
    }
    public class DonthaveExcelException : Exception
    {
        public string parName;
        public DonthaveExcelException(string s)
        { 
            parName = s;
        }
    }
    class Program
    {
      
        static void Main(string[] args)
        {
            Excel.Application excelApp = Proverka.Instance;
            try
            {
                //планируется по кнопке на выбор для каждого режима
                Console.WriteLine("Выберите режим эксперт(нажмите 1) или техназор(нажмите 2)");
                var selector = (ChangeMod)Console.ReadKey().Key;
                RangeFile obl = new RangeFile();
                obl.FirstCell = "A1";
                obl.LastCell = "L120";
                Console.WriteLine(obl.FirstCell + " " + obl.LastCell);
                switch (selector)
                {
                    case ChangeMod.expert:
                        {
                            Expert ob = new Expert();
                            ob.Initialization(excelApp);
                            ob.ProccessAll(obl);
                            break;
                        }
                    case ChangeMod.tehnadzor:
                        {
                            Tehnadzor ob = new Tehnadzor();
                            ob.Initialization(excelApp);
                            ob.ProccessAll(obl);
                            break;
                        }
                    default:
                        Console.WriteLine("Вы ввели неверный символ ");
                        break;
                }
                excelApp.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (ZapredelException exc)
            { Console.WriteLine(exc.parName); }
            catch (DirectoryNotFoundException exc)
            {
                Console.WriteLine(exc.Message);
            }
            catch (DonthaveExcelException ex)
            {
                Console.WriteLine(ex.parName);
            }
            catch (COMException exc)
            {
                Console.WriteLine(exc.Message);
            }
           
            Console.ReadLine();
        }
    }
}