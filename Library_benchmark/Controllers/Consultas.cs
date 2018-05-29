using Library_benchmark.Models;
using System;
using System.Collections.Generic;

namespace Library_benchmark.Controllers
{
    internal class Consultas
    {
        private int rows;

        public Consultas(int rows)
        {
            this.rows = rows;
        }
        

        internal IList<Elemento1> GetInformacion()
        {
            if (rows > 0)
            {
                IList<Elemento1> respuesta = new List<Elemento1>();
                for (int i = 0; i <= rows; i++)
                {
                    respuesta.Add(new Elemento1
                    {
                        Propiedad1 = "row  " + i,
                        Propiedad2 = "row  " + i,
                        Propiedad3 = "row  " + i,
                        Propiedad4 = "row  " + i,
                        Propiedad5 = "row  " + i,
                        Propiedad6 = "row  " + i,
                        Propiedad7 = "row  " + i,
                        Propiedad8 = "row  " + i,
                        Propiedad9 = "row  " + i,
                        Propiedad10 = "row  " + i,
                        Propiedad11 = "row  " + i,
                        Propiedad12 = "row  " + i,
                        Propiedad13 = "row  " + i,
                        Propiedad14 = "row  " + i,
                        Propiedad15 = "row  " + i,
                        Propiedad16 = "row  " + i,
                        Propiedad17 = "row  " + i,
                        Propiedad18 = "row  " + i,
                        Propiedad19 = "row  " + i,
                        Propiedad20 = "row  " + i,
                        Propiedad21 = "row  " + i,
                        Propiedad22 = "row  " + i,
                        Propiedad23 = "row  " + i,
                        Propiedad24 = "row  " + i,
                        Propiedad25 = "row  " + i,
                        Propiedad26 = "row  " + i

                    });
                }
                return respuesta;
            }
            else
            {
                throw new NotImplementedException();
            }
        }
    }
}