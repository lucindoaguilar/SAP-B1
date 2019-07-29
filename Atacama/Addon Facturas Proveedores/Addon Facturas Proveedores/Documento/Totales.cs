using System.Collections.Generic;
using System;

namespace Addon_Facturas_Proveedores.Documento
{
    public class Totales
    {
        public Int64 MntNeto { get; set; }
        public Int64 MntExe { get; set; }
        public Int64 MntBase { get; set; }
        public Int64 MntMargenCom { get; set; }
        public Double TasaIVA { get; set; }
        public Int64 IVA { get; set; }
        public Int64 IVAProp { get; set; }
        public Int64 IVATerc { get; set; }
        public List<ImptoReten> ImptoReten { get; set; }
        public Int64 IVANoRet { get; set; }
        public Int64 CredEC { get; set; }
        public Int64 GrntDep { get; set; }
        public ComisionesTotal ComisionesTotal { get; set; }
        public Int64 MntTotal { get; set; }
        public Int64 MontoNF { get; set; }
        public Int64 MontoPeriodo { get; set; }
        public Int64 SaldoAnterior { get; set; }
        public Int64 VlrPagar { get; set; }

        public Totales() {
            ImptoReten = new List<ImptoReten>();
            ComisionesTotal = new ComisionesTotal();
        }
    }

    public class ImptoReten
    {
        public String TipoImp { get; set; }
        public Double TasaImp { get; set; }
        public Int64 MontoImp { get; set; }

        public ImptoReten(){}
    }

    public class ComisionesTotal
    {
        public Int64 ValComNeto { get; set; }
        public Int64 ValComExe { get; set; }
        public Int64 ValComIVA { get; set; }

        public ComisionesTotal() { }
    }
}
