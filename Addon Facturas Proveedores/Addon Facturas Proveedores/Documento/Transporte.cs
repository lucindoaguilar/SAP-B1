using System.Collections.Generic;
using System;

namespace Addon_Facturas_Proveedores.Documento
{
    public class Transporte
    {
        public String Patente { get; set; }
        public String RUTTrans { get; set; }
        public String RUTChofer { get; set; }
        public String NombreChofer { get; set; }        
        public String DirDest { get; set; }
        public String CmnaDest { get; set; }
        public String CiudadDest { get; set; }
        public Aduana Aduana { get; set; }

        public Transporte() {
            Aduana = new Aduana();

        }
    }

    public class Aduana
    {
        public Int32 CodModVenta { get; set; }
        public Int32 CodClauVenta { get; set; }
        public Double TotClauVenta { get; set; }
        public Int32 CodViaTransp { get; set; }
        public String NombreTransp { get; set; }
        public String RUTCiaTransp { get; set; }
        public String NomCiaTransp { get; set; }
        public String IdAdicTransp { get; set; }
        public String Booking { get; set; }
        public String Operador { get; set; }
        public Int32 CodPtoEmbarque { get; set; }
        public String IdAdicPtoEmb { get; set; }
        public Int32 CodPtoDesemb { get; set; }
        public String IdAdicPtoDesemb { get; set; }
        public Int32 Tara { get; set; }
        public Int32 CodUnidMedTara { get; set; }
        public Double PesoBruto { get; set; }
        public Int32 CodUnidPesoBruto { get; set; }
        public Double PesoNeto { get; set; }
        public Int32 CodUnidPesoNeto { get; set; }
        public Int64 TotItems { get; set; }
        public Int64 TotBultos { get; set; }
        public List<TipoBultos> TipoBultos { get; set; }
        public Double MntFlete { get; set; }
        public Double MntSeguro { get; set; }
        public String CodPaisRecep { get; set; }
        public String CodPaisDestin { get; set; }

        public Aduana() {
            TipoBultos = new List<TipoBultos>();
        }
    }

    public class TipoBultos
    {
        public Int32 CodTpoBultos { get; set; }
        public Int64 CantBultos { get; set; }
        public String Marcas { get; set; }
        public String IdContainer { get; set; }
        public String Sello { get; set; }
        public String EmisorSello { get; set; }

        public TipoBultos() { }
    }
}
