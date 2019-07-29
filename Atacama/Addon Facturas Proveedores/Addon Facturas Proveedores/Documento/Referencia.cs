using System;
namespace Addon_Facturas_Proveedores.Documento
{
    public class Referencia
    {
        public Int32 NroLinRef { get; set; }
        public String TpoDocRef { get; set; }
        public Int32 IndGlobal { get; set; }
        public String FolioRef { get; set; }
        public String RUTOtr { get; set; }
        public String FchRef { get; set; }
        public Int32 CodRef { get; set; }
        public String RazonRef { get; set; }

        public Referencia() { }
    }
}
