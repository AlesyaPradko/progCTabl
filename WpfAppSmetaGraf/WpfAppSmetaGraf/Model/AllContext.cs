using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity;

namespace WpfAppSmetaGraf.Model
{
        public class AllContext : DbContext
        {
            public AllContext() : base("DbConnection")
            {

            }

            public DbSet<Graph> Graphs { get; set; }
            public DbSet<Table> Tables { get; set; }
            public DbSet<Day> Days { get; set; }
        }   
}
