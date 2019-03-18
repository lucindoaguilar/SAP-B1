using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Addon_Facturas_Proveedores.Comunes
{    
    public class Documento
    {
        public String febosId { get; set; }
        public Int32 tipoDocumento { get; set; }
        public Int32 folio { get; set; }
        public String fechaEmision { get; set; }
        public String fechaRecepcion { get; set; }
        public String rutEmisor { get; set; }
        public String razonSocialEmisor { get; set; }
        public Double montoTotal { get; set; }
        public Int32? formaDePago { get; set; }
        public String estadoSii { get; set; }
        public Int32 plazo { get; set; }
    }

    public class RootObjectDescarga
    {
        public String totalElementos { get; set; }
        public String totalPaginas { get; set; }
        public String paginaActual { get; set; }
        public String elementosPorPagina { get; set; }
        public List<Documento> documentos { get; set; }
        public Int32 codigo { get; set; }
        public String mensaje { get; set; }
        public String seguimientoId { get; set; }
        public List<String> errores { get; set; }
        public Int32 duracion { get; set; }
        public String hora { get; set; }
    }

    public class RootObjectGetXML
    {
        public String febosId { get; set; }
        public String rutEmisor { get; set; }
        public String rutReceptor { get; set; }
        public String razonSocialEmisor { get; set; }
        public String razonSocialReceptor { get; set; }
        public String fechaEmision { get; set; }
        public String folio { get; set; }
        public String tipo { get; set; }
        public String fechaRecepcion { get; set; }
        public String xmlData { get; set; }
        public String imagenLink { get; set; }
        public Int32 codigo { get; set; }
        public String mensaje { get; set; }
        public String seguimientoId { get; set; }
        public Int32 duracion { get; set; }
        public String hora { get; set; }
    }
}
