using System;
using System.IO;
using System.Collections.Generic;
using System.Text;

namespace E_Money_Nominas.Comunes
{
    public class NominaBci
    {
        /// <summary>
        /// Genera un txt con la nomina de pago masivo para la estructura del banco BCI
        /// </summary>
        /// <param name="RutProveedor"></param>
        /// <param name="RznSocialProveedor"></param>
        /// <param name="CodigoBcoProveedor"></param>
        /// <param name="CuentaBcoProveedor"></param>
        /// <param name="TipoDocProveedor"></param>
        /// <param name="FolioDocProveedor"></param>
        /// <param name="FechaDocProveedor"></param>
        /// <param name="MontoDocProveedor"></param>
        /// <param name="NombreArchivo"></param>
        /// <param name="FechaVctoDoc"></param>
        /// <returns></returns>
        public static ResultMessage GenerarNomina(List<ClasePagoMasivo> listaPagos, string NombreArchivo)
        {
            ResultMessage result = new ResultMessage();

            try
            {                
                StreamWriter oStringWriter = null;
                string RutProveedor = string.Empty;
                string DigVerProveedor = string.Empty;
                string ModalidadServicio = "G";
                string Unidad = string.Empty.PadLeft(5);
                string MedioAviso = "E";                
                string CodigoComuna = string.Empty.PadLeft(4, '0');
                string FormaDePago = string.Empty.PadLeft(3);
                string CodigoOfAbono = string.Empty.PadLeft(3, '0');
                string FolioRelacionado = string.Empty.PadLeft(12, '0');
                string EstadoPago = "OK".PadRight(3);
                string Glosa = string.Empty.PadLeft(200);
                string RutRetirador = string.Empty.PadLeft(8);
                string DigVerRutRetirador = string.Empty.PadLeft(1);
                string ApellidoPaterno = string.Empty.PadLeft(15);
                string ApellidoMaterno = string.Empty.PadLeft(15);
                string NombresRet1 = string.Empty.PadLeft(15);
                string RutRetirador2 = string.Empty.PadLeft(8);
                string DigVerRutRetirador2 = string.Empty.PadLeft(1);
                string ApellidoPaterno2 = string.Empty.PadLeft(15);
                string ApellidoMaterno2 = string.Empty.PadLeft(15);
                string NombresRet2 = string.Empty.PadLeft(15);
                string IdentificadorDoc = "001".PadLeft(12, '0');
                
                // Instanciar archivo, si existe lo sobreescribe.
                oStringWriter = new StreamWriter(NombreArchivo, false, Encoding.ASCII);

                foreach (ClasePagoMasivo objPago in listaPagos)
                {   
                    // Obtener rut y dígito verificador                    
                    DigVerProveedor = objPago.RutProveedor.Substring(objPago.RutProveedor.Length - 1, 1);
                    RutProveedor = objPago.RutProveedor.Substring(0, objPago.RutProveedor.Length - 1);

                    if (RutProveedor.Length < 8)
                    {
                        RutProveedor = RutProveedor.PadLeft(8, '0');
                    }

                    // Ajustar razon social
                    objPago.NombreProveedor = objPago.NombreProveedor.PadRight(45);

                    // Ajustar cuenta de proveedor
                    objPago.CuentaBcoProveedor = FuncionesComunes.ValidarCaracteres(objPago.CuentaBcoProveedor);
                    objPago.CuentaBcoProveedor = objPago.CuentaBcoProveedor.PadLeft(20, '0');

                    // Obtener Tipo de documento, SEGUN BANCO DEBE SER ODP
                    //objPago.TipoDocProveedor = TipoDocumentoPalabras(objPago.TipoDocProveedor);
                    objPago.TipoDocProveedor = "ODP";

                    // Ajustar folio
                    objPago.FolioDocProveedor = objPago.FolioDocProveedor.PadLeft(12, '0');

                    // Ajustar monto
                    objPago.MontoDocPRoveedor = objPago.MontoDocPRoveedor.PadLeft(11, '0');

                    // Obtener fecha de documento
                    objPago.FechaDocProveedor = FuncionesComunes.ValidarCaracteres(objPago.FechaDocProveedor);
                    objPago.FechaVctoDoc = FuncionesComunes.ValidarCaracteres(objPago.FechaVctoDoc);

                    // Obtener forma de pago
                    FormaDePago = FormaDePagoBCI(objPago.CodigoBcoProveedor);

                    // Ajustar codigo de banco, si es nulo o vacio asociar a BCI
                    if (string.IsNullOrEmpty(objPago.CodigoBcoProveedor))
                    {
                        objPago.CodigoBcoProveedor = "016";
                    }

                    // Ajustar correo 
                    objPago.Correo = objPago.Correo.PadLeft(35);


                    oStringWriter.Write(ModalidadServicio);     // 00 - Modalidad de servicio
                    oStringWriter.Write(RutProveedor);          // 01 - Rut del proveedor
                    oStringWriter.Write(DigVerProveedor);       // 02 . Dígito verificador del proveedor
                    oStringWriter.Write(Unidad);                // 03 - Unidad
                    oStringWriter.Write(objPago.NombreProveedor);    // 04 - Razón social proveedor
                    oStringWriter.Write(MedioAviso);            // 05 - Medio de aviso
                    oStringWriter.Write(objPago.Correo);        // 06 - Dirección de aviso
                    oStringWriter.Write(CodigoComuna);          // 07 - Código de comuna
                    oStringWriter.Write(FormaDePago);           // 08 - Forma de pago
                    oStringWriter.Write(objPago.CodigoBcoProveedor);    // 09 - Código del banco destino
                    oStringWriter.Write(objPago.CuentaBcoProveedor);    // 10 - Cuenta de abono
                    oStringWriter.Write(CodigoOfAbono);         // 11 - Código de oficina de abono
                    oStringWriter.Write(objPago.TipoDocProveedor);      // 12 - Tipo de documento
                    oStringWriter.Write(IdentificadorDoc);     // 13 - Identificador de documento
                    oStringWriter.Write(FolioRelacionado);      // 14 - Identificador documento relacionado
                    oStringWriter.Write(objPago.MontoDocPRoveedor);     // 15 - Monto informado documento 
                    oStringWriter.Write(objPago.MontoDocPRoveedor);     // 16 - Monto a pago
                    oStringWriter.Write(EstadoPago);            // 17 - Estado del pago
                    oStringWriter.Write(objPago.FechaVctoDoc);          // 18 - Fecha de vencimiento documento
                    oStringWriter.Write(objPago.FechaDocProveedor);     // 19 - Fecha de pago
                    oStringWriter.Write(Glosa);                 // 20 - Glosa
                    oStringWriter.Write(RutRetirador);          // 21 - Rut retirador # 1
                    oStringWriter.Write(DigVerRutRetirador);    // 22 - Dígito verificador Rut retirador # 1 
                    oStringWriter.Write(ApellidoPaterno);       // 23 - Apeliido paterno retirador # 1
                    oStringWriter.Write(ApellidoMaterno);       // 24 - Apellido materno retirador # 1
                    oStringWriter.Write(NombresRet1);           // 25 - Nombre retirador # 1
                    oStringWriter.Write(RutRetirador2);         // 26 - Rut retirador # 2
                    oStringWriter.Write(DigVerRutRetirador2);   // 27 - Dígito verificador Rut retirador # 2
                    oStringWriter.Write(ApellidoPaterno2);      // 28 - Apeliido paterno retirador # 2
                    oStringWriter.Write(ApellidoMaterno2);      // 29 - Apeliido materno retirador # 2
                    oStringWriter.Write(NombresRet2);           // 30 - Nombre retirador # 2

                    // Salto de linea nuevo pago
                    oStringWriter.Write(oStringWriter.NewLine);
                }

                oStringWriter.Close();
                oStringWriter.Dispose();

                result.Success = true;
                result.Mensaje = string.Format("Nomina generada en : {0}", NombreArchivo);
            }
            catch (Exception ex)
            {
                result.Success = true;
                result.Mensaje = ex.Message + ex.Source;
            }

            return result;
        }

        /// <summary>
        /// Retorna el tipo de documentos en palabra según tipo.
        /// </summary>
        /// <param name="TipoDoc"></param>
        /// <returns></returns>
        private static string TipoDocumentoPalabras(string TipoDoc)
        {
            switch (TipoDoc)
            {
                case "30":
                    TipoDoc = "FAC";
                    break;
                case "33":
                    TipoDoc = "FAC";
                    break;
                case "34":
                    TipoDoc = "FAC";
                    break;
                case "55":
                    TipoDoc = "NDB";
                    break;
                case "56":
                    TipoDoc = "NDB";
                    break;
                case "60":
                    TipoDoc = "NCR";
                    break;
                case "61":
                    TipoDoc = "NCR";
                    break;
            }

            return TipoDoc;
        }

        /// <summary>
        /// Devuelve la forma de pago que usa banco BCI
        /// </summary>
        /// <param name="Banco"></param>
        /// <param name="TipoCuenta"></param>
        /// <returns></returns>
        private static string FormaDePagoBCI(string Banco)
        {
            string FormaPago = string.Empty;

            // Es BCI
            if (Banco.Equals("016"))
            {
                FormaPago = "CCT";
            }
            // Sin banco, Vale vista
            else if (Banco.Equals("999"))
            {
                FormaPago = "VVC";
            }
            // Otro banco cuenta corriente
            else
            {
                FormaPago = "OTC";
            }

            return FormaPago;
        }
    }
}
