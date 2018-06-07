using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Library_benchmark.Models
{
    public class DummyView
    {
        public string Libreria { get; set; }
        public bool Recurso { get; set; }
        public int Registros { get; set; }
        public int Sheet { get; set; }


        public string TiempoDiseno { get; set; }
        public string TiempoCreacionDeExcel { get; set; }
        public string TiempoCreardescarga { get; set; }
        public string Total { get; set; }
    }
}