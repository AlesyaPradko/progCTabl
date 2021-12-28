using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfAppSmetaGraf.Model
{
    public class CheckIt
    {
        private static readonly Excel.Application instance = new Excel.Application();
        public static Excel.Application Instance
        {
            get
            {
                if (instance == null)
                {
                    Console.WriteLine("Excel is not installed!!");
                    return null;
                }
                return instance;
            }
        }
        static CheckIt()
        { }
        private CheckIt()
        { }
    }
}
