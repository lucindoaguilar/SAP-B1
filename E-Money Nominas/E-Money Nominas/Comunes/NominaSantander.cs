using System;
using System.IO;
using System.Collections.Generic;

namespace E_Money_Nominas.Comunes
{
    public class NominaSantander
    {
        public static ResultMessage GenerarNomina(List<ClasePagoMasivo> listaPagos, string NombreArchivo)
        {
            ResultMessage result = new ResultMessage();

            // Generar archivo CSV para poder usar xsl
            NombreArchivo = NombreArchivo.Split('.')[0];
            NombreArchivo = string.Format("{0}.csv", NombreArchivo);

            try
            {
                StreamWriter oStringWriter = null;
                string GlosaTEF = string.Empty;
                string GlosaCorreo = string.Empty;
                string GlosaCartolaCliente = string.Empty;
                string GlosaCartolaBeneficiario = string.Empty;

                // Instanciar archivo, si existe lo sobreescribe.
                oStringWriter = new StreamWriter(NombreArchivo, false);

                // Cabecera
                oStringWriter.Write("Cta_origen");
                oStringWriter.Write(",");
                oStringWriter.Write("moneda_origen");
                oStringWriter.Write(",");
                oStringWriter.Write("Cta_destino");
                oStringWriter.Write(",");
                oStringWriter.Write("moneda_destino");
                oStringWriter.Write(",");
                oStringWriter.Write("Cod_banco");
                oStringWriter.Write(",");
                oStringWriter.Write("RUT benef.");
                oStringWriter.Write(",");
                oStringWriter.Write("nombre benef.");
                oStringWriter.Write(",");
                oStringWriter.Write("Mto Total");
                oStringWriter.Write(",");
                oStringWriter.Write("Glosa TEF");
                oStringWriter.Write(",");
                oStringWriter.Write("Correo");
                oStringWriter.Write(",");
                oStringWriter.Write("Glosa correo");
                oStringWriter.Write(",");
                oStringWriter.Write("Glosa Cartola Cliente");
                oStringWriter.Write(",");
                oStringWriter.Write("Glosa Cartola Beneficiario");

                foreach (ClasePagoMasivo objPago in listaPagos)
                {
                    // Ajustar Moneda
                    if (objPago.Moneda.Equals("$"))
                    {
                        objPago.Moneda = "CLP";
                    }                    

                    // Datos
                    oStringWriter.Write(oStringWriter.NewLine);
                    oStringWriter.Write(objPago.CuentaOrigen);          // 00 - Cuenta origen
                    oStringWriter.Write(",");
                    oStringWriter.Write(objPago.Moneda);                   // 01 - Moneda origen
                    oStringWriter.Write(",");
                    oStringWriter.Write(objPago.CuentaBcoProveedor);       // 02 - Cuenta destino
                    oStringWriter.Write(",");
                    oStringWriter.Write(objPago.Moneda);                   // 03 - Moneda destino
                    oStringWriter.Write(",");
                    oStringWriter.Write(objPago.CodigoBcoProveedor);       // 04 - Código banco
                    oStringWriter.Write(",");
                    oStringWriter.Write(objPago.RutProveedor);             // 05 - Rut beneficiario
                    oStringWriter.Write(",");
                    oStringWriter.Write(objPago.NombreProveedor);       // 06 - Nombre beneficiario
                    oStringWriter.Write(",");
                    oStringWriter.Write(objPago.MontoDocPRoveedor);        // 07 - Monto total
                    oStringWriter.Write(",");
                    oStringWriter.Write(GlosaTEF);                 // 08 - Glosa TEF
                    oStringWriter.Write(",");
                    oStringWriter.Write(objPago.Correo);                   // 09 - Correo
                    oStringWriter.Write(",");
                    oStringWriter.Write(GlosaCorreo);              // 10 - Glosa correo
                    oStringWriter.Write(",");
                    oStringWriter.Write(GlosaCartolaCliente);      // 11 - Glosa cartoa cliente
                    oStringWriter.Write(",");
                    oStringWriter.Write(GlosaCartolaBeneficiario); // 12 - Glosa cartola beneficiario
                }
            
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
