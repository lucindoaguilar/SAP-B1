using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Addon_Facturas_Proveedores.Comunes
{
    public class BaseFactura
    {
        public string BaseType { get; set; }
        public string BaseRef { get; set; }
        public string BaseLine { get; set; }
        public string BaseEntry { get; set; }
        public double Total { get; set; }
        public string Descripcion { get; set; }
        public string CuentaMayor { get; set; }
        public string IndImpuesto { get; set; }
        public string Category { get; set; }
        public string[] Campo { get; set; }
        public string CodMaquinaria { get; set; }
        public string CodMantencion { get; set; }
        public string FechaMantencion { get; set; }
        public double Horometro { get; set; }

        public BaseFactura()
        {
        }
    }
}
