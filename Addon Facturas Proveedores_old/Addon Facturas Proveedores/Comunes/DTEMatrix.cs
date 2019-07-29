using System;
using System.Collections.Generic;
using Addon_Facturas_Proveedores.Documento;

namespace Addon_Facturas_Proveedores.Comunes
{
    public class DTEMatrix
    {
        public String FebosID { get; set; }
        public String DteID { get; set; }
        public DTE objDTE { get; set; }

        public DTEMatrix() { }
    }

    public static class ListaDTEMatrix
    {
        public static List<DTEMatrix> ListaDTE { get; set; }

        static ListaDTEMatrix()
        {
            ListaDTE = new List<DTEMatrix>();
        }
    }
}
