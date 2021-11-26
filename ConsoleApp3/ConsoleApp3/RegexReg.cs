using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ConsoleApp3
{
    public class RegexReg
    {
        //регулярные выражения, используемые в сметах и актах КС-2
        public Regex scopeWorkinAktKS = new Regex(@"(\bза отчетный|(К|к)оличество)", RegexOptions.IgnoreCase);
        public Regex regexmonth = new Regex(@"\.(?<month>\d{1,2})\.", RegexOptions.IgnoreCase);
        public Regex regexyear = new Regex(@"\.(?<year>\d{2,4})", RegexOptions.IgnoreCase);
        public Regex nameSsmeta = new Regex(@"((С|с)мета|\s*) №\s*\d+", RegexOptions.IgnoreCase);
        public Regex cellItogoPorazdely = new Regex("Итого по разделу");
        public Regex cellOfRazdel = new Regex(@"^Раздел");
    }
}