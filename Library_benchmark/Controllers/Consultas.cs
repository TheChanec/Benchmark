using Library_benchmark.Models;
using System;
using System.Collections.Generic;

namespace Library_benchmark.Controllers
{
    internal class Consultas
    {
        public Consultas()
        {
        }

        internal object Cabeceras()
        {
            throw new NotImplementedException();
        }

        internal object Informacion()
        {
            throw new NotImplementedException();
        }

        internal IList<Elemento1> Informacion(int rows)
        {
            if (rows > 0)
            {
                IList<Elemento1> respuesta = new List<Elemento1>();
                for (int i = 0; i <= rows; i++)
                {
                    respuesta.Add(new Elemento1
                    {
                        Propiedad1 = "col " + i,
                        Propiedad2 = "col " + i,
                        Propiedad3 = "col " + i,
                        Propiedad4 = "col " + i,
                        Propiedad5 = "col " + i,
                        Propiedad6 = "col " + i,
                        Propiedad7 = "col " + i,
                        Propiedad8 = "col " + i,
                        Propiedad9 = "col " + i,
                        Propiedad10 = "col " + i,
                        Propiedad11 = "col " + i,
                        Propiedad12 = "col " + i,
                        Propiedad13 = "col " + i,
                        Propiedad14 = "col " + i,
                        Propiedad15 = "col " + i,
                        Propiedad16 = "col " + i,
                        Propiedad17 = "col " + i,
                        Propiedad18 = "col " + i,
                        Propiedad19 = "col " + i,
                        Propiedad20 = "col " + i,
                        Propiedad21 = "col " + i,
                        Propiedad22 = "col " + i,
                        Propiedad23 = "col " + i,
                        Propiedad24 = "col " + i,
                        Propiedad25 = "col " + i,
                        Propiedad26 = "col " + i,

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