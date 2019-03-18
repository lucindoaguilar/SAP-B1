using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Addon_Facturas_Proveedores.Comunes
{
    public class AsignacionMultiple
    {
        public class AsignacionHec
        {
            public String BaseType { get; set; }
            public String BaseRef { get; set; }
            public String BaseLine { get; set; }
            public String BaseEntry { get; set; }
            public Double Total { get; set; }
            public String Descripcion { get; set; }
            public String CuentaMayor { get; set; }
            public String IndImpuesto { get; set; }
            public String Variedad { get; set; }
            public String Especie { get; set; }
            public String Category { get; set; }
            public String CatPackin { get; set; }
            public String RolPrivado { get; set; }
            //public string CodMaquinaria { get; set; }
            //public string CodMant { get; set; }
            //public string FechaMantencion { get; set; }
            //public double Horometro { get; set; }

            public AsignacionHec()
            {
            }
        }

        public static class AsignadosHec
        {
            public static List<AsignacionHec> ListaAsignacionHec { get; set; }

            static AsignadosHec()
            {
                ListaAsignacionHec = new List<AsignacionHec>();
            }
        }
    }
}
