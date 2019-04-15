using System;

namespace E_Money_Nominas.Conexiones
{
    class Conexion_SBO
    {
        #region Atributos

        /// <summary>
        /// Variable para el control y acceso a la instancia de la Aplicacion SBO en la cual se ejecuta nuestra solucion.
        /// </summary>
        public static SAPbouiCOM.Application m_SBO_Appl = null;
        /// <summary>
        /// Variable para el control y acceso a la instancia de la Compañia SBO en la cual se ejecuta nuestra solucion.
        /// </summary>
        public static SAPbobsCOM.Company m_oCompany = null;

        #endregion

        #region Constructores

        /// <summary>
        /// Establece la coneccion hacia la instacia de SBO que se esta ejecutando
        /// </summary>
        public Conexion_SBO()
        {
            try
            {
                ObtenerAplicacion();
                ConectarCompany();
            }
            catch (Exception ex)
            {
                Comunes.Msj_Appl.Errores(4, ex.Message);//mensaje de error por defecto que genera la aplicacion
                throw ex;
            }

        }

        /// <summary>
        /// Metodo para conectarse a la Aplicacion SBO que se esta ejecutando (UIAPI)
        /// </summary>
        private void ObtenerAplicacion()
        {
            try
            {
                string strConexion = ""; //variable que almacena el codigo de identificacion de conexion con SBO
                string[] strArgumentos = new string[4];
                SAPbouiCOM.SboGuiApi oSboGuiApi = null; //Variable para obtener la instacia activa de SBO

                oSboGuiApi = new SAPbouiCOM.SboGuiApi();//Instancia nueva para la gestion de la conexion
                strArgumentos = System.Environment.GetCommandLineArgs();//obtenemos el codigo de conexion del entorno configurado en "Propiedades -> Depurar -> Argumentos de la linea de comandos"

                if (strArgumentos.Length > 0)
                {
                    if (strArgumentos.Length > 1)
                    {
                        //Verificamos que la aplicacion se este ejecutando en un ambiente SBO
                        if (strArgumentos[0].LastIndexOf("\\") > 0) strConexion = strArgumentos[1];
                        else strConexion = strArgumentos[0];
                    }
                    else
                    {
                        //Verificamos que la aplicacion se este ejecutando en un ambiente SBO
                        if (strArgumentos[0].LastIndexOf("\\") > -1) strConexion = strArgumentos[0];
                        else Comunes.Msj_Appl.Errores(1, "");//mensaje de erro por no tener SBO activo
                    }
                }
                else Comunes.Msj_Appl.Errores(1, "");//mensaje de erro por no tener SBO activo


                oSboGuiApi.Connect(strConexion);//Establecemos la conexion
                m_SBO_Appl = oSboGuiApi.GetApplication(-1);//Asignamos la conexion a la aplicacion
            }
            catch (Exception ex)
            {
                Comunes.Msj_Appl.Errores(-2, ex.Message);
            }
        }

        /// <summary>
        /// Metodo para conectar a la compañia de la instacia de SBO que se esta ejecutando (DIAPI)
        /// </summary>
        public static void ConectarCompany()
        {
            try
            {
                m_oCompany = (SAPbobsCOM.Company)m_SBO_Appl.Company.GetDICompany();
            }
            catch (Exception ex)
            {
                Comunes.Msj_Appl.Errores(5, ex.Message);
            }
        }

        /// <summary>
        /// Metodo para desconeconectar la compañia de la instacia de SBO que se esta ejecutando (DIAPI)
        /// </summary>
        public static void DesconectarCompany()
        {
            try
            {
                m_oCompany.Disconnect();
            }
            catch (Exception ex)
            {
                Comunes.Msj_Appl.Errores(6, ex.Message);
            }
        }

        #endregion
    }
}
