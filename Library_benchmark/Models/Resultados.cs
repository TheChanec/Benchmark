using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Library_benchmark.Models
{
    public class Resultados
    {
        public List<EPPLUS> conEPPLUS { get; set; }
        public List<NPOI> conNPOI { get; set; }

        private static Resultados instance = null;
        private static readonly object padlock = new object();

        private Resultados() { }

        public static Resultados Instance
        {
            get
            {
                if (instance == null)
                    instance = new Resultados();

                return instance;
            }
        }
    }
}