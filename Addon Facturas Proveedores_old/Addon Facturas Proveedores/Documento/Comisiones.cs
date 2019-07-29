using System;
namespace Addon_Facturas_Proveedores.Documento
{
    public class Comisiones
    {
        public Int32 NroLinCom { get; set; }
        public String TipoMovim { get; set; } 
        public String Glosa { get; set; }
        public Double TasaComision { get; set; }
        public Int64 ValComNeto { get; set; }
        public Int64 ValComExe { get; set; }
        public Int64 ValComIVA { get; set; }

        public Comisiones() { }
    }
}
