using System.Collections.Generic;
using System;

namespace Addon_Facturas_Proveedores.Documento
{
    public class Detalle
    {
        public Int32 NroLinDet { get; set; }
        public List<CdgItem> CdgItem { get; set; }
        public String TpoDocLiq { get; set; }
        public Int32 IndExe { get; set; }
        public String IndAgente { get; set; }
        public Int64 MntBaseFaenaRet { get; set; }
        public Int64 MntMargComer { get; set; }
        public Int64 PrcConsFinal { get; set; }
        public String NmbItem { get; set; }
        public String DscItem { get; set; }
        public Double QtyRef { get; set; }
        public String UnmdRe { get; set; }
        public Double PrcRef { get; set; }
        public Double QtyItem { get; set; }
        public List<Subcantidad> Subcantidad { get; set; }
        public String FchElabor { get; set; }
        public String FchVencim { get; set; }
        public String UnmdItem { get; set; }
        public Double PrcItem { get; set; }
        public List<OtrMnda> OtrMnda { get; set; }
        public Double DescuentoPct { get; set; }
        public Int64 DescuentoMonto { get; set; }
        public List<SubDscto> SubDscto { get; set; }
        public Double RecargoPct { get; set; }
        public Int64 RecargoMonto { get; set; }
        public List<SubRecargo> SubRecargo { get; set; }
        public List<CodImpAdic> CodImpAdic { get; set; }
        public Int64 MontoItem { get; set; }
        public String ItemCode { get; set; }
        public Int32 LineNumBase { get; set; }

        public Detalle() {
            CdgItem = new List<CdgItem>();
            Subcantidad = new List<Subcantidad>();
            OtrMnda = new List<OtrMnda>();
            SubDscto = new List<SubDscto>();
            SubRecargo = new List<SubRecargo>();
            CodImpAdic = new List<CodImpAdic>();
        }
    }

    public class CdgItem
    {
        public String TpoCodigo { get; set; }
        public String VlrCodigo { get; set; }

        public CdgItem() { }
    }

    public class Subcantidad
    {
        public Double SubQty { get; set; }
        public String SubCod { get; set; }
        public String TipCodSubQty { get; set; }

        public Subcantidad() { }
    }

    public class OtrMnda
    {
        public Double PrcOtrMon { get; set; }
        public String Moneda { get; set; }
        public Double FctConv { get; set; }
        public Double DctoOtrMnda { get; set; }
        public Double RecargoOtrMnda { get; set; }
        public Double MontoItemOtrMnda { get; set; }

        public OtrMnda() { }
    }

    public class SubDscto
    {
        public String TipoDscto { get; set; }
        public Double ValorDscto { get; set; }

        public SubDscto() { }
    }

    public class SubRecargo
    {
        public String TipoRecargo { get; set; }
        public Double ValorRecargo { get; set; }

        public SubRecargo() { }
    }

    public class CodImpAdic
    {
        public String sCodImpAdic { get; set; }

        public CodImpAdic() { }
    }

}
