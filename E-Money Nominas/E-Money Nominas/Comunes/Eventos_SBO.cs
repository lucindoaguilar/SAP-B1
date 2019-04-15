using System;
using E_Money_Nominas.ClaseFormulario;
using E_Money_Nominas.Conexiones;

namespace E_Money_Nominas.Comunes
{
    /// <summary>
    /// Clase encargada de gestionar los eventos generados en SBO
    /// </summary>
    public class Eventos_SBO
    {
        #region Atributos

        public static string menuID = "";
        public static string errMsg = "";
        public static int ret = 0;

        #endregion

        #region Constructores

        /// <summary>
        /// Inicializador de la clase, en este se invocan los metodos de creacion de objeto basicos para el funcionamiento del Addon
        /// </summary>
        public Eventos_SBO()
        {
            try
            {
                RegistrarEventos();

                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Cargando Menu de Nominas...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                RegistrarMenuNominas();

                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("AddOn Nominas SBO Conectado con exito.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Msj_Appl.Errores(12, "Eventos_SBO() " + ex.Message);
            }

        }

        #endregion

        #region Eventos

        /// <summary>
        /// Metodo encargado de Gestionar los eventos de Aplicacion
        /// </summary>
        /// <param name="EventType">Objeto con la informacion completa del Evento</param>
        void m_SBO_Appl_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            try
            {
                switch (EventType)
                {
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                        System.Windows.Forms.Application.Exit();//terminamos la ejecucion del Addon
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                        System.Windows.Forms.Application.Exit();//terminamos la ejecucion del Addon
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                        System.Windows.Forms.Application.Exit();//terminamos la ejecucion del Addon
                        break;
                }
            }
            catch (Exception ex)
            {
                Msj_Appl.Errores(14, "m_SBO_Appl_AppEvent() > Eventos_SBO.cs" + ex.Message);
            }
        }

        /// <summary>
        /// Metodo para gestionar los eventos generados por las actividades en los datos.
        /// </summary>
        /// <param name="BusinessObjectInfo">Objeto con la informacion completa del evento</param>
        /// <param name="BubbleEvent">Indicador booleano para detener la cola de eventos generada</param>
        void m_SBO_Appl_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (BusinessObjectInfo.FormTypeEx)
                {
                    /* COMENTADO TODO, NO SE SABE SI SE NECESITA VALIDAR ALGUN DATO 
                //case "65081":
                //    SEI_ImprimeFolio.m_SBO_Appl_dataEvent(ref BusinessObjectInfo, out BubbleEvent);
                //    break;            
                case "133": //Factura de cliente
                    if ((System.Configuration.ConfigurationManager.AppSettings["ValidaFactura"].ToString()).Equals("Y"))
                        SEI_FormFactura.m_SBO_Appl_dataEvent(ref BusinessObjectInfo, out BubbleEvent);
                    break;
                case "65302":
                    if ((System.Configuration.ConfigurationManager.AppSettings["ValidaFacturaExenta"].ToString()).Equals("Y"))
                        SEI_FormFacturaExenta.m_SBO_Appl_dataEvent(ref BusinessObjectInfo, out BubbleEvent);
                    break;
                case "179":
                    if ((System.Configuration.ConfigurationManager.AppSettings["ValidaNotaCredito"].ToString()).Equals("Y"))
                        SEI_FormNotaCredito.m_SBO_Appl_dataEvent(ref BusinessObjectInfo, out BubbleEvent);
                    break;
                case "65303":
                    if ((System.Configuration.ConfigurationManager.AppSettings["ValidaNotaDebito"].ToString()).Equals("Y"))
                        SEI_FormNotaDebito.m_SBO_Appl_dataEvent(ref BusinessObjectInfo, out BubbleEvent);
                    break;
                case "65304":
                    if ((System.Configuration.ConfigurationManager.AppSettings["ValidaBoleta"].ToString()).Equals("Y"))
                        SEI_FormBoleta.m_SBO_Appl_dataEvent(ref BusinessObjectInfo, out BubbleEvent);
                    break;
                case "65307":
                    if ((System.Configuration.ConfigurationManager.AppSettings["ValidaFacturaExportacion"].ToString()).Equals("Y"))
                        SEI_FormFactura.m_SBO_Appl_dataEvent(ref BusinessObjectInfo, out BubbleEvent);
                    break;
                case "140":
                    if ((System.Configuration.ConfigurationManager.AppSettings["ValidaGuiaDespacho"].ToString()).Equals("Y"))
                        SEI_FormGuiaDespacho.m_SBO_Appl_dataEvent(ref BusinessObjectInfo, out BubbleEvent);
                    break;
                //case "60006":
                //    SEI_FormSerieFolio.m_SBO_Appl_dataEvent(ref BusinessObjectInfo, out BubbleEvent);
                //    break;
                case "940":
                    if ((System.Configuration.ConfigurationManager.AppSettings["ValidaTrasladoStock"].ToString()).Equals("Y"))
                        SEI_FormGuiaDespacho_TS.m_SBO_Appl_dataEvent(ref BusinessObjectInfo, out BubbleEvent);
                    break;
                //case "60004":
                //    SEI_Parametrizar.m_SBO_Appl_dataEvent(ref BusinessObjectInfo, out BubbleEvent);
                //    break;

                /*  case "65300"://Factura Anticipo
                      if ((System.Configuration.ConfigurationManager.AppSettings["ValidaFacturaAnticipo"].ToString()).Equals("Y"))
                          SEI_FormFacturaAnticipo.m_SBO_Appl_dataEvent(ref BusinessObjectInfo, out BubbleEvent);
                      break;*/
                    /*
                case "60091"://Factura Reserva
                    if ((System.Configuration.ConfigurationManager.AppSettings["ValidaFacturaReserva"].ToString()).Equals("Y"))
                        SEI_FormFacturaReserva.m_SBO_Appl_dataEvent(ref BusinessObjectInfo, out BubbleEvent);
                    break;
                    */
                }

            }
            catch (Exception ex)
            {
                Msj_Appl.Errores(14, "m_SBO_Appl_DataEvent() > Eventos_SBO.cs " + ex.Message);
            }
        }

        /// <summary>
        /// Metodo encargado de gestionar los eventos generado por los Items del sistema
        /// </summary>
        /// <param name="FormUID">Identificador del formulario</param>
        /// <param name="pVal">Objeto con el listado completo de variables de control del evento</param>
        /// <param name="BubbleEvent">Indicador booleano para detener la cola de eventos generada</param>
        void m_SBO_Appl_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.FormTypeEx)
                {
                    case "SEI_NOM":
                        SEI_FormNominas.m_SBO_Appl_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        break;
                }
            }
            catch (Exception ex)
            {
                Msj_Appl.Errores(14, "m_SBO_Appl_ItemEvent() > Eventos_SBO.cs" + ex.Message + pVal.EventType.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        /// <summary>
        /// Eventos generado por los diferentes menus del sistema
        /// </summary>
        /// <param name="pVal">Objeto con el listado completo de variables de control del evento</param>
        /// <param name="BubbleEvent">Indicador booleano para detener la cola de eventos generada</param>
        void m_SBO_Appl_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.BeforeAction)
                {
                    switch (pVal.MenuUID)
                    {
                        case "MSEI_NOM":
                            SEI_FormNominas oCAF = new SEI_FormNominas("SEI_NOM");
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Msj_Appl.Errores(14, "m_SBO_Appl_MenuEvent() > Eventos_SBO.cs " + ex.Message);
            }

        }

        #endregion

        #region Metodos

        /// <summary>
        /// Metodo encargado de resgistrar los eventos de SBO que van a ser manejados por la aplicacion.
        /// </summary>
        private void RegistrarEventos()
        {
            try
            {
                Conexion_SBO.m_SBO_Appl.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(m_SBO_Appl_AppEvent);
                Conexion_SBO.m_SBO_Appl.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(m_SBO_Appl_FormDataEvent);
                Conexion_SBO.m_SBO_Appl.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(m_SBO_Appl_ItemEvent);
                Conexion_SBO.m_SBO_Appl.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(m_SBO_Appl_MenuEvent);
                //Conexion_SBO.m_SBO_Appl.UDOEvent += new SAPbouiCOM._IApplicationEvents_UDOEventEventHandler(m
                //Conexion_SBO.m_SBO_Appl.PrintEvent += new SAPbouiCOM._IApplicationEvents_PrintEventEventHandler(m_SBO_Appl_PrintEvent);
            }
            catch (Exception ex)
            {
                Msj_Appl.Errores(10, ex.Message);
                throw ex;
            }
        }

        /// <summary>
        /// Metodo encargado de añadir nuevas entradas al menu de Activo Fijo.
        /// </summary>
        private void RegistrarMenuNominas()
        {
            try
            {
                CreaMenu("MSEI_NOM", "Nominas", "43537", SAPbouiCOM.BoMenuType.mt_STRING);
            }
            catch
            {
            }
        }

        /// <summary>
        /// Agrega una entrada a un Menu de SBO
        /// <alert class="note">
        /// <para>Para mayor sobre los tipo de menu referencia de los tipos revisar la ayuda del SDK de SBO</para>
        /// </alert>
        /// </summary>
        /// <param name="UniqueId">Identificador Unico del Menu que será creado</param>
        /// <param name="Name">Nombre que sera mostrado en SBO</param>
        /// <param name="PrincipalMenuId">Identificador Unico del Menu que contendra la nueva entrada</param>
        /// <param name="type">Tipo de Menu</param>
        private void CreaMenu(string uniqueId, string name, string principalMenuId, SAPbouiCOM.BoMenuType type)
        {
            SAPbouiCOM.MenuCreationParams objParams;
            SAPbouiCOM.Menus objSubMenu;
            int posmenu = 0;
            try
            {
                objSubMenu = Conexiones.Conexion_SBO.m_SBO_Appl.Menus.Item(principalMenuId).SubMenus;

                if (Conexiones.Conexion_SBO.m_SBO_Appl.Menus.Exists(uniqueId) == false)
                {
                    posmenu = objSubMenu.Count;
                    objParams = (SAPbouiCOM.MenuCreationParams)Conexiones.Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objParams.Type = type;
                    objParams.UniqueID = uniqueId;
                    objParams.String = name;
                    objParams.Position = posmenu + 1;
                    /*if (uniqueId == "MSEI_NOM")
                    {
                        string sPath = null;

                        sPath = Application.StartupPath.ToString();
                        sPath += "\\";
                        sPath = sPath + "Icono\\lyc.ico";
                        objParams.Image = sPath;

                    }*/
                    objSubMenu.AddEx(objParams);

                }
            }
            catch (Exception ex)
            {
                Msj_Appl.Errores(13, "CreaMenu() " + ex.Message);
            }
        }

        #endregion
    }
}
