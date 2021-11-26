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
using System.Timers;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
public enum ChangeMod { expert = 49, tehnadzor = 50, grafic = 51 };

namespace ConsoleApp3
{
    public class ZapredelException : Exception
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
        private static void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            Console.WriteLine("Время окончания работы над файлами {0:HH:mm:ss.fff}", e.SignalTime);
        }
        static void Main(string[] args)
        {
            Excel.Application excelApp = Proverka.Instance;
            try
            {
                Timer timer = new Timer();
                string adressforLogFile = @"log11.txt";
                FileStream logFile = new FileStream(adressforLogFile, FileMode.Append);
                StreamWriter logFileError = new StreamWriter(logFile);
                Stopwatch stopWatch = Stopwatch.StartNew();
                //планируется по кнопке на выбор для каждого режима
                Console.WriteLine("Выберите режим эксперт(нажмите 1) или техназор(нажмите 2) или график(3)");
                var selector = (ChangeMod)Console.ReadKey().Key;
                Console.SetOut(logFileError);
                Console.WriteLine("The application started at {0:HH:mm:ss.fff}", DateTime.Now);
                RangeFile oblastobrabotki = new RangeFile();
                oblastobrabotki.FirstCell = "A1";
                oblastobrabotki.LastCell = "Z1200";
                switch (selector)
                {
                    case ChangeMod.expert:
                        {
                            Expert ob = new Expert();
                            ob.Initialization(excelApp);
                            ob.ProccessAll(oblastobrabotki);
                            break;
                        }
                    case ChangeMod.tehnadzor:
                        {
                            Tehnadzor ob = new Tehnadzor();
                            ob.Initialization(excelApp);
                            ob.ProccessAll(oblastobrabotki);
                            break;
                        }
                    case ChangeMod.grafic:
                        {
                            Grafik ob = new Grafik();
                            ob.InitializationGrafik(excelApp);
                            ob.ProccessGrafik(oblastobrabotki, excelApp);
                            break;
                        }
                    default:
                        Console.WriteLine("Вы ввели неверный символ ");
                        break;
                }
                stopWatch.Stop();
                long elapsed = stopWatch.ElapsedMilliseconds; // or sw.ElapsedTicks
                Console.WriteLine("Total query time: {0} ms", elapsed);
                excelApp.Quit();
                logFileError.Close();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Console.ReadKey();
            }
            catch (DirectoryNotFoundException exc)
            {
                Console.WriteLine(exc.Message);
            }

            catch (IOException exc)
            {
                Console.WriteLine(exc.Message + "Ошибка при записи в файл.");
            }

            finally { }
            Console.ReadLine();
        }
    }
}






