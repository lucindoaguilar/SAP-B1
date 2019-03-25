using System;
using System.Configuration;
using System.Windows.Forms;
using Addon_Facturas_Proveedores.ClaseFormulario;
using Addon_Facturas_Proveedores.Conexiones;
using SAPbobsCOM;

namespace Addon_Facturas_Proveedores.Comunes
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

                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Cargando Estructura de datos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                CargarEstructuraDatos();
                CargarBusquedasFormateadas();         
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Cargando Menu de Factura Proveedores...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                RegistrarMenu();

                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("AddOn Factura Proveedores SBO Conectado con exito.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
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
                /*switch (BusinessObjectInfo.FormTypeEx)
                {
                    
                }*/

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
                    case "SEI_INT":
                        SEI_FormIntegracion.m_SBO_Appl_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        break;
                    case "FormDocS":
                        SEI_FormDocS.m_SBO_Appl_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        break;
                    case "FormDocSMas":
                        SEI_FormDocSMas.m_SBO_Appl_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        break;
                    case "FormSetVal":
                        SEI_FormSetVal.ItemEventEventHandler(FormUID, ref pVal, out BubbleEvent);
                        break;
                        /*
                    case "SEI_INTC":
                        SEI_FormIntegracionContado.m_SBO_Appl_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        break;
                    case "SEI_DAT":
                        SEI_FormDat.m_SBO_Appl_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        break;
                    case "SEI_ERR":
                        SEI_FormErr.m_SBO_Appl_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        break;
                        */


                        /*case "SEI_NREC":
                            SEI_FormNrec.m_SBO_Appl_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                            break;

                        case "SEI_MERC":
                            SEI_FormMerc.m_SBO_Appl_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                            break;*/

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
                        case "SEI_INT":
                            SEI_FormIntegracion oProv = new SEI_FormIntegracion("SEI_INT");
                            break;
                        case "SET_VAL":
                            SEI_FormSetVal sv = new SEI_FormSetVal();
                            break;
                            /*
                        case "SEI_INTC":
                            SEI_FormIntegracionContado oIntc = new SEI_FormIntegracionContado("SEI_INTC");
                            break;                        
                        case "SEI_ERR":
                            SEI_FormErr oErr = new SEI_FormErr("SEI_ERR");
                            break;
                            */
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
        private void RegistrarMenu()
        {
            try
            {
                CreaMenu("SEI_FELP", "Integración DTE Proveedores", "2304", SAPbouiCOM.BoMenuType.mt_POPUP);
                CreaMenu("SEI_CRE", "Crédito", "SEI_FELP", SAPbouiCOM.BoMenuType.mt_POPUP);
                CreaMenu("SEI_INT", "Integración SAP", "SEI_CRE", SAPbouiCOM.BoMenuType.mt_STRING);
                CreaMenu("SET_VAL", "Integración SAP", "SEI_FELP", SAPbouiCOM.BoMenuType.mt_STRING);
                //CreaMenu("SEI_CONT", "Contado", "SEI_FELP", SAPbouiCOM.BoMenuType.mt_POPUP);
                //CreaMenu("SEI_INTC", "Integración SAP", "SEI_CONT", SAPbouiCOM.BoMenuType.mt_STRING);
                //CreaMenu("SEI_ERR", "ver DTE erroneo no contado", "SEI_CONT", SAPbouiCOM.BoMenuType.mt_STRING);
                //CreaMenu("SEI_ERM", "Enviar Recibo mercaderia", "SEI_FELP", SAPbouiCOM.BoMenuType.mt_STRING);
                //CreaMenu("SEI_NREC", "Documentos no recibidos", "SEI_FELP", SAPbouiCOM.BoMenuType.mt_STRING);
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
                    if (uniqueId == "SEI_FELP")
                    {
                        string sPath = null;

                        sPath = Application.StartupPath.ToString();
                        sPath += "\\";
                        sPath = sPath + "Icono\\lyc.jpg";
                        objParams.Image = sPath;

                    }
                    objSubMenu.AddEx(objParams);

                }
            }
            catch (Exception ex)
            {
                Msj_Appl.Errores(13, "CreaMenu() " + ex.Message);
            }
        }

        /// <summary>
        /// Funcion que permite cargar la estructura de datos del Addon (Tablas, campos, UDO...etc)
        /// </summary>
        private void CargarEstructuraDatos()
        {            
            CargarCampos();
        }

        
        /// <summary>
        /// Funcion que permite cargar los campos necesarios en sus tablas respectivas
        /// </summary>
        private static void CargarCampos()
        {
            try
            {
                CreaCampoMD("OADM", "SEI_TOKEN", "Token Febos 3", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null);
                CreaCampoMD("OADM", "SEI_RECINTO", "Recinto", 150, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null);
                ValorValido[] valoresCont = new ValorValido[] 
                { new ValorValido("S", " "),
                  new ValorValido("N", "Documento generado como contado por error")
                };
                CreaCampoMD("OPCH", "SEI_CONTADO", "Es Contado", 1, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, "S", valoresCont, null);
                CreaCampoMD("OPCH", "SEI_FEBOSID", "Febos ID", 1, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, "S", valoresCont, null);
                CreaCampoMD("OPOR", "SEI_FCHV", "Fecha de Vencimiento", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null);
            }
            catch (Exception e)
            {
                Msj_Appl.Errores(13, "Carga de Campo de Usuario" + e.Message);
            }
        }

        /// <summary>
        /// Funcion que permite crear los campos sus respectivas tablas
        /// </summary>
        /// <param name="nombretabla">Se especifica la tabla del campo a crear</param>
        /// <param name="nombrecampo">Se especifica el campo a crear</param>
        /// <param name="descripcion">Se especifica la descripción del campo a crear</param>
        /// <param name="longitud">Se especifica la longitud del campo a crear</param>
        /// <param name="tipo">Se especifica el tipo de dato del campo </param>
        /// <param name="subtipo">Se especifica el subtipo de dato del campo. Solo apliaca para tipo de datos Float</param>
        public static void CreaCampoMD(String nombretabla, String nombrecampo, String descripcion, int longitud, SAPbobsCOM.BoFieldTypes tipo, SAPbobsCOM.BoFldSubTypes subtipo, SAPbobsCOM.BoYesNoEnum mandatory, String defaultValue, ValorValido[] valores, String linkTable)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldsMD;
            try
            {
                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = nombretabla;//Se obtiene el nombre de la tabla de usario
                oUserFieldsMD.Name = nombrecampo;//Se asigna el nombre del campo de usuario
                oUserFieldsMD.Description = descripcion;//Se asigna una descripcion al campo de usuario
                oUserFieldsMD.Mandatory = mandatory;
                if (longitud > 0)
                {
                    oUserFieldsMD.EditSize = longitud;//Se define una longitud al campo de usuario
                }
                oUserFieldsMD.Type = tipo;//Se asigna el tipo de dato al campo de usuario
                oUserFieldsMD.SubType = subtipo;                

                if (defaultValue != null) oUserFieldsMD.DefaultValue = defaultValue;

                if (valores != null && valores.Length > 0)
                {
                    foreach (ValorValido vv in valores)
                    {
                        oUserFieldsMD.ValidValues.Value = vv.valor;
                        oUserFieldsMD.ValidValues.Description = vv.descripcion;
                        oUserFieldsMD.ValidValues.Add();
                    }
                }

                oUserFieldsMD.LinkedTable = linkTable;

                ret = oUserFieldsMD.Add();//Se agrega el campo de usuario

                if (ret != 0 && ret != -2035) //&& ret != -5002)
                {
                    Conexion_SBO.m_oCompany.GetLastError(out ret, out errMsg);
                    Msj_Appl.Errores(8, "CargarTabla -> " + errMsg);
                }

                Comunes.FuncionesComunes.LiberarObjetoGenerico(oUserFieldsMD);
            }
            catch 
            {

            }
        }

        /// <summary>
        /// Crea y enlaza querys para uso del addon
        /// </summary>
        public static void CargarBusquedasFormateadas()
        {
            try
            {
                String Query = String.Empty;
                String CategoryName = "Consultas Addon Recepcion DTE Proveedores";
                Int32 CategoryID = 0;
                SAPbobsCOM.QueryCategories oCategory = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.oQueryCategories);
                oCategory.Name = CategoryName;
                Int32 ret = oCategory.Add();

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        Query = "SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\" = '" + CategoryName + "'";
                        break;
                    default:
                        Query = "SELECT CategoryId FROM OQCN WHERE CatName = '" + CategoryName + "'";
                        break;
                }
             
               
                SAPbobsCOM.Recordset oRecordSet = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(Query);

                if (!oRecordSet.EoF)
                {
                    CategoryID = Convert.ToInt32(oRecordSet.Fields.Item(0).Value);
                }

                SAPbobsCOM.UserQueries oQuery = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.oUserQueries);
                oQuery.QueryCategory = CategoryID;
                oQuery.QueryDescription = "Trae Dimension 1";

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        oQuery.Query = "Select \"PrcCode\", \"PrcName\" from OPRC where \"DimCode\"= 1 and \"Active\" = 'Y' and \"PrcCode\" NOT IN ('Stelle_Z')";
                        break;
                    default:
                        oQuery.Query = "Select PrcCode, PrcName from OPRC where DimCode= 1 and Active = 'Y' and PrcCode NOT IN ('Stelle_Z')";
                        break;
                }
                
                oQuery.QueryType = UserQueryTypeEnum.uqtWizard;
                oQuery.Add();

                oQuery = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.oUserQueries);
                oQuery.QueryCategory = CategoryID;
                oQuery.QueryDescription = "Trae Dimension 2";

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        oQuery.Query = "Select \"PrcCode\", \"PrcName\" from OPRC where \"DimCode\"= 2 and \"Active\" = 'Y' and \"PrcCode\" NOT IN ('Stelle_2')";
                        break;
                    default:
                        oQuery.Query = "Select PrcCode, PrcName from OPRC where DimCode= 2 and Active = 'Y' and PrcCode NOT IN ('Stelle_2')";
                        break;
                }
                
                oQuery.QueryType = UserQueryTypeEnum.uqtWizard;
                oQuery.Add();

                oQuery = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.oUserQueries);
                oQuery.QueryCategory = CategoryID;
                oQuery.QueryDescription = "Trae Dimension 3";

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        oQuery.Query = "Select \"PrcCode\", \"PrcName\" from OPRC where \"DimCode\"= 3 and \"Active\" = 'Y' and \"PrcCode\" NOT IN ('Stelle_3')";
                        break;
                    default:
                        oQuery.Query = "Select PrcCode, PrcName FROM OPRC where DimCode = 3 and Active = 'Y' and PrcCode NOT IN ('Stelle_3')";
                        break;
                }
                
                oQuery.QueryType = UserQueryTypeEnum.uqtWizard;
                oQuery.Add();

                oQuery = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.oUserQueries);
                oQuery.QueryCategory = CategoryID;
                oQuery.QueryDescription = "Trae Dimension 4";

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        oQuery.Query = "Select \"PrcCode\", \"PrcName\" from OPRC where \"DimCode\"= 4 and \"Active\" = 'Y' and \"PrcCode\" NOT IN ('Stelle_4')";
                        break;
                    default:
                        oQuery.Query = "Select PrcCode, PrcName FROM OPRC where DimCode = 4 and Active = 'Y' and PrcCode NOT IN ('Stelle_4')";
                        break;
                }                

                oQuery.QueryType = UserQueryTypeEnum.uqtWizard;
                oQuery.Add();

                oQuery = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.oUserQueries);
                oQuery.QueryCategory = CategoryID;
                oQuery.QueryDescription = "Trae Dimension 5";

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        oQuery.Query = "Select \"PrcCode\", \"PrcName\" from OPRC where \"DimCode\"= 5 and \"Active\" = 'Y' and \"PrcCode\" NOT IN ('Stelle_5')";
                        break;
                    default:
                        oQuery.Query = "Select PrcCode, PrcName FROM OPRC where DimCode = 5 and Active = 'Y' and PrcCode NOT IN ('Stelle_5')";
                        break;
                }                              

                oQuery.QueryType = UserQueryTypeEnum.uqtWizard;
                oQuery.Add();

                oQuery = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.oUserQueries);
                oQuery.QueryCategory = CategoryID;
                oQuery.QueryDescription = "Trae Cuenta Contable";

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        oQuery.Query = "Select \"AcctCode\", \"AcctName\" from OACT order by \"AcctCode\"";
                        break;
                    default:
                        oQuery.Query = "Select AcctCode, AcctName FROM OACT order by AcctCode";
                        break;
                }
                
                oQuery.QueryType = UserQueryTypeEnum.uqtWizard;
                oQuery.Add();

                Int32 QueryID = 0;
                String QueryName = "Trae Dimension 1";

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        Query = "SELECT \"IntrnalKey\" FROM OUQR WHERE \"QName\" = '" + QueryName + "' and \"QCategory\" = " + CategoryID + "";
                        break;
                    default:
                        Query = "SELECT IntrnalKey FROM OUQR WHERE QName = '" + QueryName + "' and QCategory = " + CategoryID + "";
                        break;
                }                
                
                oRecordSet = null;
                oRecordSet = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                oRecordSet.DoQuery(Query);

                if (!oRecordSet.EoF)
                {
                    QueryID = Convert.ToInt32(oRecordSet.Fields.Item(0).Value);
                }

                SAPbobsCOM.FormattedSearches oFtts = null;
                oFtts = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.oFormattedSearches);

                oFtts.Action = BoFormattedSearchActionEnum.bofsaQuery;
                oFtts.FormID = "SEI_DAT";
                oFtts.ItemID = "Item_0";
                oFtts.ColumnID = "col_Esp";
                oFtts.QueryID = QueryID;
                oFtts.ByField = BoYesNoEnum.tNO;
                oFtts.ForceRefresh = BoYesNoEnum.tNO;
                oFtts.Refresh = BoYesNoEnum.tNO;
                oFtts.Add();

                QueryID = 0;
                QueryName = "Trae Dimension 2";

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        Query = "SELECT \"IntrnalKey\" FROM OUQR WHERE \"QName\" = '" + QueryName + "' and \"QCategory\" = " + CategoryID + "";
                        break;
                    default:
                        Query = "SELECT IntrnalKey FROM OUQR WHERE QName = '" + QueryName + "' and QCategory = " + CategoryID + "";
                        break;
                }                
                
                oRecordSet = null;
                oRecordSet = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                oRecordSet.DoQuery(Query);

                if (!oRecordSet.EoF)
                {
                    QueryID = Convert.ToInt32(oRecordSet.Fields.Item(0).Value);
                }

                oFtts = null;
                oFtts = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.oFormattedSearches);

                oFtts.Action = BoFormattedSearchActionEnum.bofsaQuery;
                oFtts.FormID = "SEI_DAT";
                oFtts.ItemID = "Item_0";
                oFtts.ColumnID = "col_Var";
                oFtts.QueryID = QueryID;
                oFtts.ByField = BoYesNoEnum.tNO;
                oFtts.ForceRefresh = BoYesNoEnum.tNO;
                oFtts.Refresh = BoYesNoEnum.tNO;
                oFtts.Add();

                QueryID = 0;
                QueryName = "Trae Dimension 3";

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        Query = "SELECT \"IntrnalKey\" FROM OUQR WHERE \"QName\" = '" + QueryName + "' and \"QCategory\" = " + CategoryID + "";
                        break;
                    default:
                        Query = "SELECT IntrnalKey FROM OUQR WHERE QName = '" + QueryName + "' and QCategory = " + CategoryID + "";
                        break;
                }               
               
                oRecordSet = null;
                oRecordSet = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                oRecordSet.DoQuery(Query);

                if (!oRecordSet.EoF)
                {
                    QueryID = Convert.ToInt32(oRecordSet.Fields.Item(0).Value);
                }

                oFtts = null;
                oFtts = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.oFormattedSearches);

                oFtts.Action = BoFormattedSearchActionEnum.bofsaQuery;
                oFtts.FormID = "SEI_DAT";
                oFtts.ItemID = "Item_0";
                oFtts.ColumnID = "col_Cat";
                oFtts.QueryID = QueryID;
                oFtts.ByField = BoYesNoEnum.tNO;
                oFtts.ForceRefresh = BoYesNoEnum.tNO;
                oFtts.Refresh = BoYesNoEnum.tNO;
                oFtts.Add();

                QueryID = 0;
                QueryName = "Trae Dimension 4";

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        Query = "SELECT \"IntrnalKey\" FROM OUQR WHERE \"QName\" = '" + QueryName + "' and \"QCategory\" = " + CategoryID + "";
                        break;
                    default:
                        Query = "SELECT IntrnalKey FROM OUQR WHERE QName = '" + QueryName + "' and QCategory = " + CategoryID + "";
                        break;
                }                
                
                oRecordSet = null;
                oRecordSet = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                oRecordSet.DoQuery(Query);

                if (!oRecordSet.EoF)
                {
                    QueryID = Convert.ToInt32(oRecordSet.Fields.Item(0).Value);
                }

                oFtts = null;
                oFtts = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.oFormattedSearches);

                oFtts.Action = BoFormattedSearchActionEnum.bofsaQuery;
                oFtts.FormID = "SEI_DAT";
                oFtts.ItemID = "Item_0";
                oFtts.ColumnID = "col_Catpa";
                oFtts.QueryID = QueryID;
                oFtts.ByField = BoYesNoEnum.tNO;
                oFtts.ForceRefresh = BoYesNoEnum.tNO;
                oFtts.Refresh = BoYesNoEnum.tNO;
                oFtts.Add();

                QueryID = 0;
                QueryName = "Trae Dimension 5";

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        Query = "SELECT \"IntrnalKey\" FROM OUQR WHERE \"QName\" = '" + QueryName + "' and \"QCategory\" = " + CategoryID + "";
                        break;
                    default:
                        Query = "SELECT IntrnalKey FROM OUQR WHERE QName = '" + QueryName + "' and QCategory = " + CategoryID + "";
                        break;
                }                
                
                oRecordSet = null;
                oRecordSet = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                oRecordSet.DoQuery(Query);

                if (!oRecordSet.EoF)
                {
                    QueryID = Convert.ToInt32(oRecordSet.Fields.Item(0).Value);
                }

                oFtts = null;
                oFtts = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.oFormattedSearches);

                oFtts.Action = BoFormattedSearchActionEnum.bofsaQuery;
                oFtts.FormID = "SEI_DAT";
                oFtts.ItemID = "Item_0";
                oFtts.ColumnID = "col_Rolp";
                oFtts.QueryID = QueryID;
                oFtts.ByField = BoYesNoEnum.tNO;
                oFtts.ForceRefresh = BoYesNoEnum.tNO;
                oFtts.Refresh = BoYesNoEnum.tNO;
                oFtts.Add();

                QueryID = 0;
                QueryName = "Trae Cuenta Contable";

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        Query = "SELECT \"IntrnalKey\" FROM OUQR WHERE \"QName\" = '" + QueryName + "' and \"QCategory\" = " + CategoryID + "";
                        break;
                    default:
                        Query = "SELECT IntrnalKey FROM OUQR WHERE QName = '" + QueryName + "' and QCategory = " + CategoryID + "";
                        break;
                }
                
                oRecordSet = null;
                oRecordSet = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                oRecordSet.DoQuery(Query);

                if (!oRecordSet.EoF)
                {
                    QueryID = Convert.ToInt32(oRecordSet.Fields.Item(0).Value);
                }

                oFtts = null;
                oFtts = Conexion_SBO.m_oCompany.GetBusinessObject(BoObjectTypes.oFormattedSearches);

                oFtts.Action = BoFormattedSearchActionEnum.bofsaQuery;
                oFtts.FormID = "SEI_DAT";
                oFtts.ItemID = "Item_0";
                oFtts.ColumnID = "col_Cta";
                oFtts.QueryID = QueryID;
                oFtts.ByField = BoYesNoEnum.tNO;
                oFtts.ForceRefresh = BoYesNoEnum.tNO;
                oFtts.Refresh = BoYesNoEnum.tNO;
                oFtts.Add();

            }
            catch (Exception ex)
            {
                String err = ex.Message;
            }
        }

        public class ValorValido
        {
            public String valor;
            public String descripcion;

            public ValorValido(String valor, String descripcion)
            {
                this.valor = valor;
                this.descripcion = descripcion;
            }
        }
        #endregion
    }
}
