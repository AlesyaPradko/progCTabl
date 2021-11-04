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
        { parName = s; }
    }
    public class DonthaveExcelException : Exception
    {
        public string paName;
        public DonthaveExcelException(string s)
        { paName = s; }
    }
    class Program
    {
        public delegate void regim(int n, List<Excel.Workbook> cop, List<string> adS, List<string> adK, List<Excel.Workbook> cPap, Dictionary<string, List<string>> kS,string s1, string s2);
        //static void Main(string[] args)
        //{
        //    Excel.Application excelApp = Proverka.Instance;
        //    //планируется выбор папки пользователем, где лежат сметы, поэтому значение полю задается в Main
        //    try
        //    {
        //        string usersmetu = @"D:\иксу";
        //        //Tabl.ContainPapkaSmeta = ParserExc.GetListKS(usersmetu, excelApp);
        //        var containPapkaSmeta = ParserExc.GetListKS(usersmetu, excelApp);

        //        //Tabl.AdresSmeta = ParserExc.Getstring(usersmetu);
        //        var adresSmeta = ParserExc.Getstring(usersmetu);
        //        if (containPapkaSmeta.Count==0 || adresSmeta.Count == 0)
        //        { 
        //            throw new DonthaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку");
        //        }


        //        //планируется выбор папки пользователем, где лежат Акты КС-2, поэтому значение полю задается в Main
        //        string userKS = @"D:\икси";
        //        //Tabl.AdresKS = ParserExc.Getstring(userKS);
        //        var adresKS = ParserExc.Getstring(userKS);

        //        //планируется выбор папки пользователем, куда сохранить измененные сметы, поэтому значение полю задается в Main
        //        string userwheresave = @"D:\икси 2";
        //        userwheresave += "\\Копия";
        //        //Tabl.CopySmet = ParserExc.MadeCopyExcbook(userwheresave, usersmetu, excelApp, Tabl.ContainPapkaSmeta, Tabl.AdresSmeta);
        //        //Tabl.ContainPapkaKS = ParserExc.GetListKS(userKS, excelApp);
        //        var copySmet = ParserExc.MadeCopyExcbook(userwheresave, usersmetu, excelApp, containPapkaSmeta, adresSmeta);
        //        var containPapkaKS = ParserExc.GetListKS(userKS, excelApp);
        //        if (containPapkaKS.Count == 0 || adresKS.Count == 0)
        //        {
        //            throw new DonthaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку");
        //        }
        //        //Tabl.KskSmete = ParserExc.GetContainSM(Tabl.ContainPapkaKS, Tabl.AdresSmeta, Tabl.AdresKS);
        //        var kskSmete = ParserExc.GetContainSM(containPapkaKS, adresSmeta, adresKS);

        //        //планируется по кнопке на выбор для каждого режима
        //        Console.WriteLine("Выберите режим эксперт(нажмите 1) или техназор(нажмите 2)");
        //        ChangeMod chan;
        //        int changeregim = (int)(Console.ReadKey().Key);
        //        chan = (ChangeMod)Console.ReadKey().Key;
        //        string sx1 = "A1";
        //        string sx2 = "L120";
        //        for (int u = 0; u < Tabl.CopySmet.Count; u++)
        //        {
        //            switch (chan)
        //            {
        //                case ChangeMod.expert:
        //                    {
        //                        Expert ob = new Expert();
        //                        ob.Worklikeexpert(u, Tabl.CopySmet, Tabl.AdresSmeta, Tabl.AdresKS, Tabl.ContainPapkaKS, Tabl.KskSmete, sx1, sx2);
        //                        break;
        //                    }
        //                case ChangeMod.tehnadzor:
        //                    {
        //                        Tehnadzor ob = new Tehnadzor();
        //                        del = ob.Workliketehnadzor;
        //                        del(u, Tabl.CopySmet, Tabl.AdresSmeta, Tabl.AdresKS, Tabl.ContainPapkaKS, Tabl.KskSmete, sx1, sx2);
        //                        break;
        //                    }
        //                default:
        //                    Console.WriteLine("Вы ввели неверный символ ");
        //                    break;
        //            }
        //        }
            
        //    excelApp.Quit();
        //    GC.Collect();
        //    GC.WaitForPendingFinalizers();
        //    GC.Collect();
        //    GC.WaitForPendingFinalizers();
        //    }
        //    catch (DirectoryNotFoundException exc)
        //    { Console.WriteLine(exc.Message); }
        //    catch (DonthaveExcelException ex)
        //    { Console.WriteLine(ex.paName); }
        //    catch (COMException exc)
        //    { Console.WriteLine(exc.Message); }
           
        //    Console.ReadLine();
        //}


        static void Main(string[] args)
        {
            Excel.Application excelApp = Proverka.Instance;
            //планируется выбор папки пользователем, где лежат сметы, поэтому значение полю задается в Main
            try
            {
                string usersmetu = @"D:\иксу";
                //Tabl.ContainPapkaSmeta = ParserExc.GetListKS(usersmetu, excelApp);
                var containPapkaSmeta = ParserExc.GetListKS(usersmetu, excelApp);

                //Tabl.AdresSmeta = ParserExc.Getstring(usersmetu);
                var adresSmeta = ParserExc.Getstring(usersmetu);
                if (containPapkaSmeta.Count == 0 || adresSmeta.Count == 0)
                {
                    throw new DonthaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку");
                }


                //планируется выбор папки пользователем, где лежат Акты КС-2, поэтому значение полю задается в Main
                string userKS = @"D:\икси";
                //Tabl.AdresKS = ParserExc.Getstring(userKS);
                var adresKS = ParserExc.Getstring(userKS);

                //планируется выбор папки пользователем, куда сохранить измененные сметы, поэтому значение полю задается в Main
                string userwheresave = @"D:\икси 2";
                userwheresave += "\\Копия";
                //Tabl.CopySmet = ParserExc.MadeCopyExcbook(userwheresave, usersmetu, excelApp, Tabl.ContainPapkaSmeta, Tabl.AdresSmeta);
                //Tabl.ContainPapkaKS = ParserExc.GetListKS(userKS, excelApp);
                var copySmet = ParserExc.MadeCopyExcbook(userwheresave, usersmetu, excelApp, containPapkaSmeta, adresSmeta);
                var containPapkaKS = ParserExc.GetListKS(userKS, excelApp);
                if (containPapkaKS.Count == 0 || adresKS.Count == 0)
                {
                    throw new DonthaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку");
                }
                //Tabl.KskSmete = ParserExc.GetContainSM(Tabl.ContainPapkaKS, Tabl.AdresSmeta, Tabl.AdresKS);
                var kskSmete = ParserExc.GetContainSM(containPapkaKS, adresSmeta, adresKS);

                //планируется по кнопке на выбор для каждого режима
                Console.WriteLine("Выберите режим эксперт(нажмите 1) или техназор(нажмите 2)");
                var selector = (ChangeMod)Console.ReadKey().Key;
                string sx1 = "A1";
                string sx2 = "L120";
                switch (selector)
                {
                    case ChangeMod.expert:
                        {
                            Expert ob = new Expert();
                            ob.ProccessAll(sx1, sx2);
                            break;
                        }
                    case ChangeMod.tehnadzor:
                        {
                            Tehnadzor ob = new Tehnadzor();
                            ob.ProccessAll(sx1, sx2);
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
            catch (DirectoryNotFoundException exc)
            { Console.WriteLine(exc.Message); }
            catch (DonthaveExcelException ex)
            { Console.WriteLine(ex.paName); }
            catch (COMException exc)
            { Console.WriteLine(exc.Message); }

            Console.ReadLine();
        }
    }
}