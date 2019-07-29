using System;
namespace Addon_Facturas_Proveedores.Documento
{
    public class Emisor
    {
        public String RUTEmisor { get; set; }
        public String RznSoc { get; set; }
        public String GiroEmis { get; set; }
        public String Telefono { get; set; }
        public String CorreoEmisor { get; set; }
        public String Acteco { get; set; }
        public Int32 CdgTraslado { get; set; }
        public Int32 FolioAut { get; set; }
        public String FchAut { get; set; }
        public String Sucursal { get; set; }
        public String CdgSIISucur { get; set; }
        public String DirOrigen { get; set; }
        public String CmnaOrigen { get; set; }
        public String CiudadOrigen { get; set; }
        public String CdgVendedor { get; set; }
        public String IdAdicEmisor { get; set; }

        public Emisor() { }
    }
}
