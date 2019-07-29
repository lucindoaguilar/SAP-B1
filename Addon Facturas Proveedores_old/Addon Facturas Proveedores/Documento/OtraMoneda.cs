using System.Collections.Generic;
using System;

namespace Addon_Facturas_Proveedores.Documento
{
    public class OtraMoneda
    {
        public String TpoMoneda { get; set; }
        public Double TpoCambio { get; set; }
        public Double MntNetoOtrMnda { get; set; }
        public Double MntExeOtrMnda { get; set; }
        public Double MntFaeCarneOtrMnda { get; set; }
        public Double MntMargComOtrMnda { get; set; }
        public Double IVAOtrMnda { get; set; }
        public List<ImpRetOtrMnda> ImpRetOtrMnda { get; set; }
        public Double IVANoRetOtrMnda { get; set; }
        public Double MntTotOtrMnda { get; set; }

        public OtraMoneda() {
            ImpRetOtrMnda = new List<ImpRetOtrMnda>();
        }
    }

    public class ImpRetOtrMnda
    {
        public String TipoImpOtrMnda { get; set; }
        public Double TasaImpOtrMnda { get; set; }
        public Int64 VlrImpOtrMnda { get; set; }

        public ImpRetOtrMnda(){ }
    }
}
