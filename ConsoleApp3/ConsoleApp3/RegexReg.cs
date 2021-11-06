using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ConsoleApp3
{
    public class RegexReg
    {
        //регулярные выражения, используемые в сметах и актах КС-2
        public Regex regexval = new Regex(@"(\bза отчетный|(К|к)оличество)", RegexOptions.IgnoreCase);
        public Regex regexmes = new Regex(@"\.(?<month>\d{1,2})\.", RegexOptions.IgnoreCase);
        public Regex regexgod = new Regex(@"\.(?<year>\d{2,4})", RegexOptions.IgnoreCase);
        public Regex namesmet = new Regex(@"(С|с)мета\s*№\d+", RegexOptions.IgnoreCase);
    }
}