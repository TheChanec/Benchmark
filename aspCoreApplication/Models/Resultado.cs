using System.Collections.Generic;

namespace aspCoreApplication.Models
{
    public class Resultado
    {
        public int Intento { get; set; }

        public Parametros Parametro { get; set; }

        public string Libreria { get; set; }
        public List<Tiempo> Tiempos { get; set; } = new List<Tiempo>();
    }
}