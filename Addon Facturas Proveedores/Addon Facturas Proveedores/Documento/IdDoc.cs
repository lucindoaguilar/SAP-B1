using System.Collections.Generic;
using System;

namespace Addon_Facturas_Proveedores.Documento
{
    public class IdDoc
    {
        public String TipoDTE { get; set; }
        public Int64 Folio { get; set; }
        public String FchEmis { get; set; }
        public Int32 IndNoRebaja { get; set; }
        public Int32 TipoDespacho { get; set; }
        public Int32 IndTraslado { get; set; }
        public String TpoImpresion { get; set; }
        public Int32 IndServicio { get; set; }
        public Int32 MntBruto { get; set; }
        public Int32 FmaPago { get; set; }
        public Int32 FmaPagExp { get; set; }
        public String FchCancel { get; set; }
        public Int64 MntCancel { get; set; }
        public Int64 SaldoInsol { get; set; }
        public List<MntPagos> MntPagos { get; set; }
        public String PeriodoDesde { get; set; }
        public String PeriodoHasta { get; set; }
        public String MedioPago { get; set; }
        public String TipoCtaPago { get; set; }
        public String NumCtaPago { get; set; }
        public String BcoPago { get; set; }
        public String TermPagoCdg { get; set; }
        public String TermPagoGlosa { get; set; }
        public String TermPagoDias { get; set; }
        public String FchVenc { get; set; }

        public IdDoc() {
            MntPagos = new List<MntPagos>();
        }
    }

    public class MntPagos
    {
        public String FchPago { get; set; }
        public Int64 MntPago { get; set; }
        public String GlosaPagos { get; set; }

        public MntPagos() { }
    }
}
