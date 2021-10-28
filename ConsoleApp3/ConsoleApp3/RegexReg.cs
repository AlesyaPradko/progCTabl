using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ConsoleApp3
{
    public class RegexReg
    {
        //регулярные выражения, используемые в сметах и актах КС-2
        public Regex regexval = new Regex(@"(\bотчетный|Количество)", RegexOptions.IgnoreCase);
        public Regex regexKS = new Regex(@"^Акт №\s*\d+", RegexOptions.IgnoreCase);
        public Regex regexdat = new Regex(@"^\d{2}.\d{2}.\d{4}", RegexOptions.IgnoreCase);
        public Regex regexmes = new Regex(@"\.\d{2}\.", RegexOptions.IgnoreCase);
        public Regex regexgod = new Regex(@"\.\d{4}", RegexOptions.IgnoreCase);
    }
}