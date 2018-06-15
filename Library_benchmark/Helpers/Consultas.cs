using Library_benchmark.Models;
using System;
using System.Collections.Generic;

namespace Library_benchmark.Helpers
{
    internal class Consultas
    {
        private int rows;

        public Consultas(int rows)
        {
            this.rows = rows;
        }


        internal IList<Dummy> GetInformacion()
        {
            if (rows > 0)
            {
                IList<Dummy> respuesta = new List<Dummy>();
                for (int i = 0; i <= rows; i++)
                {
                    respuesta.Add(new Dummy
                    {
                        Fecha1 = DateTime.Now.Date,
                        Fecha2 = DateTime.Now.Date,
                        Fecha3 = DateTime.Now.Date,
                        Fecha4 = DateTime.Now.Date,
                        Moneda1 = (decimal)i,
                        Moneda2 = (decimal)i,
                        Moneda3 = (decimal)i,
                        Moneda4 = (decimal)i,

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