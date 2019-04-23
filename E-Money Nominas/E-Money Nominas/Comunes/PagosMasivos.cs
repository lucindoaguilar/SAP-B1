using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using E_Money_Nominas.Conexiones;

namespace E_Money_Nominas.Comunes
{
    /// <summary>
    /// Clase que controla los querys de la BD
    /// </summary>
    public class PagosMasivos
    {
        /// <summary>
        /// Metodo que devuelve una lista de pagos de la BD.
        /// </summary>
        /// <returns></returns>
        public static List<string> ObtenerPagosExistentes(string Fecha)
        {
            List<string> ListaPagos = new List<string>();

            // Ajustar fecha a formato indicado AAAA/DD/MM
            DateTime dt = Convert.ToDateTime(Fecha);
            string Mes = (dt.Month.ToString().Length == 1) ? dt.Month.ToString().PadLeft(2, '0') : dt.Month.ToString();
            string Dia = (dt.Day.ToString().Length == 1) ? dt.Day.ToString().PadLeft(2, '0') : dt.Day.ToString();

            Fecha = string.Format("{0}/{1}/{2}", Dia, Mes, dt.Year.ToString());
            //Fecha = string.Format("{0}/{1}/{2}", dt.Year.ToString(), Mes, Dia);

            // Establecer objeto de datos y consulta
            SAPbobsCOM.Recordset oRecordSet = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery(FuncionesComunes.QueryPagosExistentes(Fecha));

            while (!oRecordSet.EoF)
            {
                ListaPagos.Add(oRecordSet.Fields.Item(0).Value.ToString());
                oRecordSet.MoveNext();
            }

            return ListaPagos;
        }        

        /// <summary>
        /// Gestiona la creación de archivo TXT para pagos masivos.
        /// </summary>
        /// <param name="Pago"></param>
        public static ResultMessage GestionarPagoMasivo(string Pago)
        {
            ResultMessage result = new ResultMessage();

            try
            {
                // Variables.
                SAPbobsCOM.Recordset oRecordSet = null;
                string Directorio = string.Empty;
                string NombreBancoLocal = string.Empty;
                string NombreArchivo = string.Empty;
                string FechaDocProveedor = string.Empty;
                string BancoLocal = string.Empty;
                List<ClasePagoMasivo> listaPagos = new List<ClasePagoMasivo>();
                ClasePagoMasivo objPago = null;
                int index = 0;

                // Obtener banco de destino de nomina para ejecutar pago masivo
                oRecordSet = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(FuncionesComunes.QueryObtenerBanco(Pago));

                if (!oRecordSet.EoF)
                {
                    BancoLocal = oRecordSet.Fields.Item(0).Value.ToString();
                }
                else
                {
                    result.Success = false;
                    result.Mensaje = "El pago no se puede procesar";
                    return result;
                }

                // Obtener información básica del pago según banco.
                // Santander se hace de forma distinta
                if(BancoLocal.Equals("037"))
                {
                    oRecordSet = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery(FuncionesComunes.QueryPagosMasivosSantander(Pago));

                    while (!oRecordSet.EoF)
                    {
                        objPago = new ClasePagoMasivo();

                        objPago.CuentaBcoProveedor = oRecordSet.Fields.Item(0).Value.ToString();
                        objPago.CodigoBcoProveedor = oRecordSet.Fields.Item(1).Value.ToString();
                        objPago.RutProveedor = oRecordSet.Fields.Item(2).Value.ToString();
                        objPago.NombreProveedor = oRecordSet.Fields.Item(3).Value.ToString();
                        objPago.MontoDocPRoveedor = oRecordSet.Fields.Item(4).Value.ToString();
                        objPago.BancoLocal = oRecordSet.Fields.Item(5).Value.ToString();
                        BancoLocal = oRecordSet.Fields.Item(5).Value.ToString();
                        objPago.NombreBancoLocal = oRecordSet.Fields.Item(6).Value.ToString();
                        NombreBancoLocal = oRecordSet.Fields.Item(6).Value.ToString();
                        objPago.Moneda = oRecordSet.Fields.Item(7).Value.ToString();
                        objPago.CuentaOrigen = oRecordSet.Fields.Item(8).Value.ToString();
                        objPago.Correo = oRecordSet.Fields.Item(9).Value.ToString();

                        listaPagos.Add(objPago);
                        index++;
                        oRecordSet.MoveNext();
                    }

                    oRecordSet = null;

                }
                else
                {
                    oRecordSet = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery(FuncionesComunes.QueryPagosMasivos(Pago));

                    while (!oRecordSet.EoF)
                    {
                        objPago = new ClasePagoMasivo();

                        objPago.RutProveedor = oRecordSet.Fields.Item(2).Value.ToString();
                        objPago.NombreProveedor = oRecordSet.Fields.Item(3).Value.ToString();
                        objPago.CodigoBcoProveedor = oRecordSet.Fields.Item(1).Value.ToString();
                        objPago.CuentaBcoProveedor = oRecordSet.Fields.Item(0).Value.ToString();
                        objPago.TipoDocProveedor = oRecordSet.Fields.Item(7).Value.ToString();
                        objPago.FolioDocProveedor = oRecordSet.Fields.Item(5).Value.ToString();
                        DateTime dtFecha = Convert.ToDateTime(oRecordSet.Fields.Item(6).Value.ToString());
                        objPago.FechaDocProveedor = dtFecha.ToShortDateString();
                        FechaDocProveedor = dtFecha.ToShortDateString();
                        objPago.MontoDocPRoveedor = oRecordSet.Fields.Item(4).Value.ToString();
                        dtFecha = Convert.ToDateTime(oRecordSet.Fields.Item(8).Value.ToString());
                        objPago.FechaVctoDoc = dtFecha.ToShortDateString();
                        objPago.BancoLocal = oRecordSet.Fields.Item(9).Value.ToString();
                        BancoLocal = oRecordSet.Fields.Item(9).Value.ToString();
                        objPago.NombreBancoLocal = oRecordSet.Fields.Item(10).Value.ToString();
                        NombreBancoLocal = oRecordSet.Fields.Item(10).Value.ToString();
                        objPago.Moneda = oRecordSet.Fields.Item(11).Value.ToString();
                        objPago.CuentaOrigen = oRecordSet.Fields.Item(12).Value.ToString();
                        objPago.Correo = oRecordSet.Fields.Item(13).Value.ToString();

                        listaPagos.Add(objPago);
                        index++;
                        oRecordSet.MoveNext();
                    }

                    oRecordSet = null;
                }
                
                // Generar Archivo 
                Directorio = ConfigurationManager.AppSettings["RutaNominas"].ToString();
                Directorio = string.Format("{0}{1}\\", Directorio, NombreBancoLocal);
                // Si no existe directorio se crea
                if (!Directory.Exists(Directorio))
                {
                    Directory.CreateDirectory(Directorio);
                }
                NombreArchivo = string.Format("{0}{1}.txt", Directorio, FuncionesComunes.ValidarCaracteres(Pago));

                // Generar documento según banco
                switch(BancoLocal)
                {
                    case "016":
                        result = NominaBci.GenerarNomina(listaPagos, NombreArchivo);
                        break;
                    case "049":
                        result = NominaSecurity.GenerarNomina(listaPagos, NombreArchivo);
                        break;
                    case "037":
                        result = NominaSantander.GenerarNomina(listaPagos, NombreArchivo);
                        break;
                    case "001":
                        result = NominaBancoDeChile.GenerarNomina(listaPagos, NombreArchivo);
                        break;
                    default:
                        result.Success = false;
                        result.Mensaje = "El banco del pago seleccionado no esta implementado";
                        break;
                }
                // Banco BCI
                if (BancoLocal.Equals("016"))
                {
                    
                }
                // Banco SECURITY
                else if(BancoLocal.Equals("049"))
                {
                    
                }
                // Banco SANTANDER
                else if (BancoLocal.Equals("037"))
                {
                   
                }
                else
                {
                    
                }
                               
            }
            catch (IOException exIO)
            {
                result.Success = false;
                result.Mensaje = exIO.Message;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Mensaje = ex.Message;
            }

            return result;
        }

        
    }
}
