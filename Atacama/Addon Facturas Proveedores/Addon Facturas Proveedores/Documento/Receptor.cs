using System;
namespace Addon_Facturas_Proveedores.Documento
{
    public class Receptor
    {
        public String RUTRecep { get; set; }
        public String CdgIntRecep { get; set; } 
        public String RznSocRecep { get; set; }
        public String NumId { get; set; }
        public String Nacionalidad { get; set; }
        public String IdAdicRecep { get; set; }
        public String GiroRecep { get; set; }
        public String Contacto { get; set; }
        public String CorreoRecep { get; set; }
        public String DirRecep { get; set; }
        public String CmnaRecep { get; set; }
        public String CiudadRecep { get; set; }
        public String DirPostal { get; set; }
        public String CmnaPostal { get; set; }
        public String CiudadPostal { get; set; }

        public Receptor() { }
    }
}
