using System.Collections.Generic;

namespace Library_benchmark.Models
{
    public class Singleton
    {
        private ICollection<Resultado> _resultados = new List<Resultado>();

        public ICollection<Resultado> Resultados { get => _resultados; set => _resultados = value; }
        private static Singleton _instance;
        
        private Singleton() { }

        public static Singleton Instance
        {
            get
            {
                if (_instance == null) {
                    _instance = new Singleton();
                    
                }
                    

                return _instance;
            }
        }

        
    }
}