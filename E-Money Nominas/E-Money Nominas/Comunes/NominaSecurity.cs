using System;
using System.IO;
using System.Collections.Generic;

namespace E_Money_Nominas.Comunes
{
    public class NominaSecurity
    {
        /// <summary>
        /// Genera un txt con la nomina de pago masivo para la estructura del banco SECURITY
        /// </summary>
        /// <param name="RutProveedor"></param>
        /// <param name="RznSocialProveedor"></param>
        /// <param name="CodigoBcoProveedor"></param>
        /// <param name="CuentaBcoProveedor"></param>
        /// <param name="MontoDocProveedor"></param>
        /// <param name="NombreArchivo"></param>
        /// <returns></returns>
        public static ResultMessage GenerarNomina(List<ClasePagoMasivo> listaPagos, string NombreArchivo)
        {
            ResultMessage result = new ResultMessage();

            try
            {
                StreamWriter oStringWriter = null;
                string TipoRegistro = "2";
                string MedioPago = "1".PadLeft(2, '0');
                string DireccionBeneficiario = string.Empty.PadRight(30);
                string ComunaBeneficiario = string.Empty.PadRight(20);
                string CiudadBeneficiario = string.Empty.PadRight(20);
                string ActividadEconomica = string.Empty.PadRight(2);
                string Oficina = string.Empty.PadRight(4);
                string MotivoPago = string.Empty.PadRight(300);
                string ControlIntegridad = string.Empty.PadLeft(13, '0');
                string Filler = string.Empty.PadRight(466);

                // Instanciar archivo, si existe lo sobreescribe.
                oStringWriter = new StreamWriter(NombreArchivo, false);

                foreach (ClasePagoMasivo objPago in listaPagos)
                {
                    // Ajustar Rut
                    objPago.RutProveedor = objPago.RutProveedor.PadLeft(10, '0');

                    // Ajustar razon social
                    objPago.NombreProveedor = objPago.NombreProveedor.PadRight(50);

                    // Ajustar cuenta
                    objPago.CuentaBcoProveedor = objPago.CuentaBcoProveedor.PadRight(22);

                    // Ajustar monto
                    objPago.MontoDocPRoveedor = objPago.MontoDocPRoveedor.PadLeft(14, '0');
                    objPago.MontoDocPRoveedor = objPago.MontoDocPRoveedor.PadRight(16, '0');

                    // Instanciar archivo, si existe lo sobreescribe.
                    oStringWriter = new StreamWriter(NombreArchivo, true);
                    oStringWriter.Write(TipoRegistro);         // 00 - Tipo de registro
                    oStringWriter.Write(objPago.RutProveedor);         // 01 - Rut beneficiario
                    oStringWriter.Write(objPago.NombreProveedor);   // 02 - Nombre beneficiario
                    oStringWriter.Write(objPago.CodigoBcoProveedor);   // 03 - Código de banco destino
                    oStringWriter.Write(MedioPago);            // 04 - Medio de pago
                    oStringWriter.Write(objPago.CuentaBcoProveedor);   // 05 - Número cuenta destino
                    oStringWriter.Write(objPago.MontoDocPRoveedor);    // 06 - Monto del pago
                    oStringWriter.Write(DireccionBeneficiario);// 07 - Dirección beneficiario
                    oStringWriter.Write(ComunaBeneficiario);   // 08 - Comuna beneficiario
                    oStringWriter.Write(CiudadBeneficiario);   // 09 - Ciudad beneficiario
                    oStringWriter.Write(ActividadEconomica);   // 10 - Actividad económica
                    oStringWriter.Write(Oficina);              // 11 - Oficina o sucursal
                    oStringWriter.Write(MotivoPago);           // 12 - Motivo del pago                    

                    // Salto de linea nuevo pago
                    oStringWriter.Write(oStringWriter.NewLine);
                }

                // Seguridad
                TipoRegistro = "4";
                oStringWriter.Write(oStringWriter.NewLine);
                oStringWriter.Write(TipoRegistro);          // 00 - Tipo de registro
                oStringWriter.Write(ControlIntegridad);     // 01 - Control de integridad
                oStringWriter.Write(Filler);                // 02 - Filler

                oStringWriter.Close();
                oStringWriter.Dispose();

                result.Success = true;
                result.Mensaje = string.Format("Nomina generada en : {0}", NombreArchivo);
            }
            catch (Exception ex)
            {
                result.Success = true;
                result.Mensaje = ex.Message;
            }

            return result;
        }
    }
}
