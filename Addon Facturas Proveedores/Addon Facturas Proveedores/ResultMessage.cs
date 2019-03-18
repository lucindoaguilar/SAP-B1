using System.Runtime.Serialization;

namespace Addon_Facturas_Proveedores
{
    [DataContract()]
    public class ResultMessage
    {
        [DataMember]
        public int Id;

        [DataMember]
        public string Mensaje = "";

        [DataMember]
        public bool Success;

        [DataMember]
        public object DTE;
    }
}
