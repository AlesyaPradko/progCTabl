using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelEditor.bl
{
    public class RegexReg
    {
        public Regex scopeWorkInAktKS = new Regex(@"((К|к)оличество|Кол\.)", RegexOptions.IgnoreCase);
        public Regex regexMonth = new Regex(@"\.?(?<month>\d{2})\.", RegexOptions.IgnoreCase);
        public Regex regexYear = new Regex(@"\.(?<year>\d{4})", RegexOptions.IgnoreCase);
        public Regex regexData = new Regex(@"(?<month>\d{2})\.(?<year>\d{4})", RegexOptions.IgnoreCase);
        public Regex nameSmeta = new Regex(@"((С|с)мета|\s*) №\s*\d+", RegexOptions.IgnoreCase);
        public Regex cellTotalForChapter = new Regex("Итого по разделу");
        public Regex cellOfRazdel = new Regex(@"^Раздел");
    }
}
