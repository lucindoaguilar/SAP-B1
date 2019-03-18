using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Addon_Facturas_Proveedores
{

    public class Documents
    {
        public string RutEmisor { get; set; }
        public string Tipo { get; set; }
        public string Folio { get; set; }
        public string FebId { get; set; }
        public string CardCode { get; set; }
        public string FchEmis { get; set; }
        public string FchVenc { get; set; }
        public string MntTotal { get; set; }
        public string IVA { get; set; }
    }
}
