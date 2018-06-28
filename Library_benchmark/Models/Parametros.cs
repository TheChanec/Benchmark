using System;
using System.Collections.Generic;

namespace Library_benchmark.Models
{
    public class Parametros
    {
        public int IdExcel { get; set; }
        public int Rows { get; set; }
        public int Sheets { get; set; }
        public bool Design { get; set; }
        public bool Mascaras { get; set; }
        public bool Resource { get; set; }
        public int Iteraciones { get; set; }
        public ICollection<object> Exceles { get; set; } = new List<Object>() { new { Id = 1, Nombre = "NPOI" }, new { Id = 2, Nombre = "EPPLUS" } };
    }
}