using System.Collections.Generic;

namespace Addon_Facturas_Proveedores.Documento
{
    public class DTE
    {
        // Encabezado Doc
        public IdDoc IdDoc { get; set; }
        public Emisor Emisor { get; set; }
        public Receptor Receptor { get; set; }
        public Transporte Transporte { get; set; }
        public Totales Totales { get; set; }
        public OtraMoneda OtraMoneda { get; set; }

        // Detalle
        public List<Detalle> Detalle { get; set; }
        // SubTotales
        public List<SubTotInfo> SubTotInfo { get; set; }
        // Descuento Recargo
        public List<DscRcgGlobal> DscRcgGlobal { get; set; }
        // Referencias
        public List<Referencia> Referencia { get; set; }
        // Comisiones
        public List<Comisiones> Comisiones { get; set; }

        public DTE() {
            IdDoc = new IdDoc();
            Emisor = new Emisor();
            Receptor = new Receptor();
            Transporte = new Transporte();
            Totales = new Totales();
            OtraMoneda = new OtraMoneda();
            Detalle = new List<Detalle>();
            SubTotInfo = new List<SubTotInfo>();
            DscRcgGlobal = new List<DscRcgGlobal>();
            Referencia = new List<Referencia>();
            Comisiones = new List<Comisiones>();
        }
    }
}
