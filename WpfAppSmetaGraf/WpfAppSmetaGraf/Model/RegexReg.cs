using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace WpfAppSmetaGraf.Model
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
