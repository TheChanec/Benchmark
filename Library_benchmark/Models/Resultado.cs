using System.Collections.Generic;

namespace Library_benchmark.Models
{
    public class Resultado
    {
        public int Intento { get; set; }
        private List<Tiempo> _tiempos = new List<Tiempo>();

        public Parametros Parametro { get; set; }

        public string Libreria { get; set; }
        public List<Tiempo> Tiempos { get => _tiempos; set => _tiempos = value; }
    }
}