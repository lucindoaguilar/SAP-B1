using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Addon_Facturas_Proveedores.Comunes
{
    public class Filas
    {
        public String FebosId { get; set; }
        public Int32 LineNum { get; set; }
        public String Servicio { get; set; }
        public Double Total { get; set; }
        public String Impuesto { get; set; }
        public String Cuenta { get; set; }
        public String Especie { get; set; }
        public String Variedad { get; set; }
        public String Category { get; set; }
        public String CatPacking { get; set; }
        public String RolPrivado { get; set; }

        public Filas() { }
    }

    public static class ListaFilas
    {
        public static List<Filas> ListFilas { get; set; }

        static ListaFilas()
        {
            ListFilas = new List<Filas>();
        }
    }
}
