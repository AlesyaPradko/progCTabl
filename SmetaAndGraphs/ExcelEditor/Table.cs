using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEditor.bl
{
    public class Table
    {
        public int Id { get; set; }
        public string NameSmeta { get; set; }
        public ICollection<Graph> Graphs { get; set; }
    }
}
