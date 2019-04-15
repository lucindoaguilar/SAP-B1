
namespace E_Money_Nominas.Comunes
{
    public class ClasePagoMasivo
    {
        public string Directorio { get; set; }
        public string NombreArchivo { get; set; }
        public string RutProveedor { get; set; }
        public string DigVerProveedor { get; set; }
        public string NombreProveedor { get; set; }
        public string CodigoBcoProveedor { get; set; }
        public string CuentaBcoProveedor { get; set; }
        public string TipoDocProveedor { get; set; }
        public string FolioDocProveedor { get; set; }
        public string FechaDocProveedor { get; set; }
        public string MontoDocPRoveedor { get; set; }
        public string FechaVctoDoc { get; set; }
        public string BancoLocal { get; set; }
        public string NombreBancoLocal { get; set; }
        public string Moneda { get; set; }
        public string CuentaOrigen { get; set; }
        public string Correo { get; set; }

        public ClasePagoMasivo()
        {
        }
    }
}
