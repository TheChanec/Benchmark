using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Library_benchmark.Models
{
    public class Singleton
    {
        private ICollection<Resultado> _resultados = new List<Resultado>();

        public ICollection<Resultado> Resultados { get => _resultados; set => _resultados = value; }
        private static Singleton instance = null;
        private static readonly object padlock = new object();

        private Singleton() { }

        public static Singleton Instance
        {
            get
            {
                if (instance == null) {
                    instance = new Singleton();
                    
                }
                    

                return instance;
            }
        }

        
    }
}