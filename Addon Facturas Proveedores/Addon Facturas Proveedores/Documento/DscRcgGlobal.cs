using System;
namespace Addon_Facturas_Proveedores.Documento
{
    public class DscRcgGlobal
    {
        public Int32 NroLinDR { get; set; }
        public String TpoMov { get; set; }
        public String GlosaDR { get; set; }
        public String TpoValor { get; set; }
        public Double ValorDR { get; set; }
        public Double ValorDROtrMnda { get; set; }
        public Int32 IndExeDR { get; set; }

        public DscRcgGlobal() { }
    }
}
