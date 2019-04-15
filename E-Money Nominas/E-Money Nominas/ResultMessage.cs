using System.Runtime.Serialization;

namespace E_Money_Nominas
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
    }
}
