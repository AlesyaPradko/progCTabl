using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfAppSmetaGraf.Model
{
    public class Graph
    {
        public int Id { get; set; }
        public string NameChapter { get; set; }
        public string NameWork { get; set; }
        public int? TableId { get; set; }
        public Table Table { get; set; }
    }
}
