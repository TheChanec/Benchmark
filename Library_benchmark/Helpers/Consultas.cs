using Library_benchmark.Models;
using System;
using System.Collections.Generic;

namespace Library_benchmark.Helpers
{
    public class Consultas
    {
        private readonly int _rows;

        public Consultas()
        {
        }

        /// <summary>
        /// Contructor Inicial de la clase consultas 
        /// </summary>
        /// <param name="rows"> Numero de registros que se retornaran en un List </param>
        public Consultas(int rows)
        {
            _rows = rows;
        }

        /// <summary>
        /// Obtiene el la lista generada de entidad ExcelDummy
        /// </summary>
        /// <returns></returns>
        internal IList<ExcelDummy> GetExcelInformacion()
        {

            IList<ExcelDummy> respuesta = new List<ExcelDummy>();
            for (var i = 1; i <= _rows; i++)
            {
                respuesta.Add(new ExcelDummy
                {
                    Fecha1 = DateTime.Now.Date,
                    Fecha2 = DateTime.Now.Date,
                    Fecha3 = DateTime.Now.Date,
                    Fecha4 = DateTime.Now.Date,
                    Fecha5 = DateTime.Now.Date,
                    Fecha6 = DateTime.Now.Date,

                    Moneda1 = i,
                    Moneda2 = i,
                    Moneda3 = i,
                    Moneda4 = i,
                    Moneda5 = i,
                    Moneda6 = i,

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

        /// <summary>
        /// Obtiene valores dumi para la generacion de un PDF
        /// </summary>
        /// <returns></returns>
        internal PdfDummy GetPdfInformacion()
        {
            return new PdfDummy
            {
                Title = "",
                SubTitle = "",
                Date = new DateTime().Date,
                Driver = "",
                Truck = "",
                OdometerStatr = 9,
                PressurePsi = 9,
                DevicePsi = 9,
                Water = "",
                MechanicalComments = ""
            };

        }
    }
}