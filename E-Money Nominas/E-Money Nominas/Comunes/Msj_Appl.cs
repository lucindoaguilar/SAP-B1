using System;

namespace E_Money_Nominas.Comunes
{
    /// <summary>
    /// Clase donde se realiza la gestion de los mensajes de la aplicacion
    /// </summary>
    class Msj_Appl
    {
        public static ResultMessage Result = new ResultMessage();

        #region Metodos

        /// <summary>
        /// Metodo que retorna el mensaje de error para un codigo determinado.
        /// </summary>
        /// <param name="errorCode">Codigo del error</param>
        /// <param name="errMsj">Mensaje del error(Complementario)</param>
        public static ResultMessage Errores(int errorCode, string errMsj)
        {
            string Msj = "";
            try
            {
                switch (errorCode)
                {
                    case -2:
                        Msj = errMsj;
                        Result = MostrarMsjWF(errorCode, Msj, false);
                        break;
                    case -1:
                        Msj = errMsj;
                        Result = MostrarMsjWF(errorCode, Msj, false);
                        break;
                    case 1:
                        Msj = "Error. Folio no asignado.";
                        Result = MostrarMsjWF(errorCode, Msj, false);
                        break;
                    case 2:
                        Msj = "Error. Indicador no asignado.";
                        Result = MostrarMsjWF(errorCode, Msj, false);
                        break;
                    case 3:
                        Msj = "Error. DocEntry no existe.";
                        Result = MostrarMsjWF(errorCode, Msj, false);
                        break;
                    case 14:
                        Msj = "Error originado en: " + errMsj;
                        Result = MostrarMsjWF(errorCode, Msj, false);
                        break;
                    default:
                        Msj = errMsj;
                        Result = MostrarMsjWF(errorCode, Msj, false);
                        break;
                }
                return Result;
            }
            catch (Exception ex)
            {
                Msj = ex.Message;
                Result = MostrarMsjWF(1000, Msj, false);
                return Result;
            }
        }

        /// <summary>
        /// Metodo que retorna un mensaje de Exito para un codigo especifico
        /// </summary>
        /// <param name="SuccesCode">Codigo del Mensaje</param>
        /// <param name="Msj">Mensaje del Exito (Complementario)</param>
        public static ResultMessage Exitos(int SuccesCode, string Msj)
        {
            try
            {
                switch (SuccesCode)
                {
                    case 1:
                        Msj = "Conexión realizada con exito";
                        Result = MostrarMsjWF(SuccesCode, Msj, true);
                        break;
                    case 2:
                        Msj = "Parametros correctos";
                        Result = MostrarMsjWF(SuccesCode, Msj, true);
                        break;
                    case 6:
                        Msj = "Documento de texto creado con exito en la ruta: " + Msj;
                        Result = MostrarMsjWF(SuccesCode, Msj, true);
                        break;
                    case 7:
                        Msj = "Documento Enviado con exito";
                        Result = MostrarMsjWF(SuccesCode, Msj, true);
                        break;
                    default:
                        Result = MostrarMsjWF(SuccesCode, Msj, true);
                        break;
                }
                return Result;
            }
            catch (Exception ex)
            {
                Msj = ex.Message;
                Result = MostrarMsjWF(1000, Msj, false);
                return Result;
            }
        }

        /// <summary>
        /// Metodo que retorna un mensaje de Adventencia para un codigo especifico
        /// </summary>
        /// <param name="WarningCode">Codigo de la Advertecia</param>
        /// <param name="Msj">Mensaje de Advertencia (Complementatio)</param>
        //public static ResultMessage Advertencias(int WarningCode, string Msj)
        //{
        //    try
        //    {
        //        switch (WarningCode)
        //        {
        //            case 1:
        //                Msj = "(1) Se genrará la estructura para el Addon: " + Msj;
        //                MostrarMsjSBO(Msj, 1);
        //                break;
        //            case 2:
        //                Msj = "(2) Se actualizará la estructura de datos para el Addon: " + Msj;
        //                MostrarMsjSBO(Msj, 1);
        //                break;
        //            case 3:
        //                Msj = "(3) Se detecto una version superior del Addon: " + Msj;
        //                MostrarMsjSBO(Msj, 1);
        //                break;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Msj = ex.Message;
        //        MostrarMsjWF(Msj);
        //    }

        //}

        /// <summary>
        /// Metodo para mostrar los mensajes en la barra de estatus de SBO
        /// </summary>
        /// <param name="Msj">Mensaje</param>
        /// /// <param name="Tipo">Tipo de Mensaje: 0 Error, 1 Advertencia, 2 Exito, 3 sin tipo</param>
        //private static void MostrarMsjSBO(string Msj, int Tipo)
        //{
        //    try
        //    {
        //        switch (Tipo)
        //        {
        //            case 0:
        //                Conexiones.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(Msj, SAPbouiCOM.BoMessageTime.bmt_Short,
        //                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);//Muestra el mensaje el barra de estatus de SBO
        //                break;
        //            case 1:
        //                Conexiones.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(Msj, SAPbouiCOM.BoMessageTime.bmt_Short,
        //                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);//Muestra el mensaje el barra de estatus de SBO
        //                break;
        //            case 2:
        //                Conexiones.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(Msj, SAPbouiCOM.BoMessageTime.bmt_Short,
        //                    SAPbouiCOM.BoStatusBarMessageType.smt_Success);//Muestra el mensaje el barra de estatus de SBO
        //                break;
        //            case 3:
        //                Conexiones.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(Msj, SAPbouiCOM.BoMessageTime.bmt_Short,
        //                    SAPbouiCOM.BoStatusBarMessageType.smt_None);//Muestra el mensaje el barra de estatus de SBO
        //                break;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MostrarMsjWF(ex.Message);
        //    }
        //}

        /// <summary>
        /// Metodo para mostrar los mensajes en formularios estandar de windows "System.Windows.Forms"
        /// </summary>
        /// <param name="Msj">Mensaje</param>
        private static ResultMessage MostrarMsjWF(int id, string Msj, bool tipo)
        {
            ResultMessage response = new ResultMessage();
            response.Id = id;
            response.Mensaje = Msj;
            response.Success = tipo;

            return response;
            //System.Windows.Forms.MessageBox.Show(Msj);
        }

        #endregion
    }
}
