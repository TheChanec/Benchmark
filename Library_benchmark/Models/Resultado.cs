using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Library_benchmark.Models
{
    public class Resultado
    {
        public List<Tiempo> Tiempos { get; set; }
        
        private static Resultado instance = null;
        private static readonly object padlock = new object();

        private Resultado() { }

        public static Resultado Instance
        {
            get
            {
                if (instance == null) {
                    instance = new Resultado();
                    instance.Tiempos = new List<Tiempo>();
                }
                    

                return instance;
            }
        }
    }
}