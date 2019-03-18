using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Addon_Facturas_Proveedores.Comunes;
using Addon_Facturas_Proveedores.Conexiones;
using Addon_Facturas_Proveedores.Documento;
using RestSharp;
using Newtonsoft.Json;

namespace Addon_Facturas_Proveedores.ClaseFormulario
{
    public class SEI_FormIntegracionContado
    {
        public SEI_FormIntegracionContado(String FormUID)
        {
            CargarXML();
            CargarFormulario(FormUID);
        }

        /// <summary>
        /// Carga el formulario de ingreso múltiple a traves de su XML
        /// </summary>
        public static void CargarXML()
        {
            XmlDocument oXmlDoc = null;
            ResultMessage result = new ResultMessage();
            bool bFormAbierto = false;
            SAPbouiCOM.Form oForm = null;
            int index;

            try
            {
                for (index = 0; index < Conexion_SBO.m_SBO_Appl.Forms.Count; index++)
                {
                    SAPbouiCOM.Form oFormAbierto = Conexion_SBO.m_SBO_Appl.Forms.Item(index);
                    if (oFormAbierto.UniqueID.Equals("SEI_INTC"))
                    {
                        bFormAbierto = true;
                        break;
                    }
                }

                // Si el form no esta abierto se abre
                if (!bFormAbierto)
                {
                    oXmlDoc = new XmlDocument();

                    String xml = "";
                    xml = Application.StartupPath.ToString();
                    xml += "\\";
                    xml = xml + "Formularios\\" + "FormIntc.srf";
                    oXmlDoc.Load(xml);

                    String sXML = oXmlDoc.InnerXml.ToString();
                    SAPbouiCOM.FormCreationParams creationPackage = null;
                    creationPackage = (SAPbouiCOM.FormCreationParams)Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                    creationPackage.UniqueID = "SEI_INTC";
                    creationPackage.FormType = "SEI_INTC";
                    creationPackage.Modality = SAPbouiCOM.BoFormModality.fm_None;
                    creationPackage.XmlData = sXML;
                    oForm = Conexion_SBO.m_SBO_Appl.Forms.AddEx(creationPackage);
                }
                // Si no traer al frente
                else
                {
                    Conexion_SBO.m_SBO_Appl.Forms.Item(index).Select();
                }
            }
            catch (Exception ex)
            {
                result = Msj_Appl.Errores(14, "CargarXML > SEI_FormIntc.cs " + ex.Message);
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Forma el formulario de documentos, con los campos necesarios.
        /// </summary>
        private void CargarFormulario(String FormUID)
        {
            SAPbouiCOM.Form oForm = Conexion_SBO.m_SBO_Appl.Forms.Item(FormUID);            
            oForm.Freeze(true);

            Int16 Top = 10;
            Int16 Left = 10;

            // static Desde Fecha
            SAPbouiCOM.Item oItem = oForm.Items.Add("stDesde", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            SAPbouiCOM.StaticText itemST = ((SAPbouiCOM.StaticText)(oItem.Specific));
            itemST.Item.Top = Top;
            itemST.Item.Width = 80;
            itemST.Item.Left = Left;
            itemST.Caption = "Desde Fecha:";

            // Campo  Desde Fecha
            oItem = oForm.Items.Add("etDesde", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = Top;
            oItem.Left = Left + oForm.Items.Item("stDesde").Width;
            oItem.Width = 110;
            SAPbouiCOM.EditText oEditText = ((SAPbouiCOM.EditText)oItem.Specific);
            oForm.DataSources.UserDataSources.Add("dsDesde", SAPbouiCOM.BoDataType.dt_DATE);
            oEditText.DataBind.SetBound(true, "", "dsDesde");

            // static Hasta Fecha
            oItem = oForm.Items.Add("stHasta", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            itemST = ((SAPbouiCOM.StaticText)(oItem.Specific));
            itemST.Item.Top = Top;
            itemST.Item.Width = 80;
            oItem.Left = Left + oForm.Items.Item("stDesde").Width + oForm.Items.Item("etDesde").Width + 20;
            itemST.Caption = "Hasta Fecha:";

            // Campo  Hasta Fecha
            oItem = oForm.Items.Add("etHasta", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = Top;
            oItem.Left = Left + oForm.Items.Item("stDesde").Width + oForm.Items.Item("etDesde").Width + oForm.Items.Item("stHasta").Width + 20;
            oItem.Width = 110;
            oEditText = ((SAPbouiCOM.EditText)oItem.Specific);
            oForm.DataSources.UserDataSources.Add("dsHasta", SAPbouiCOM.BoDataType.dt_DATE);
            oEditText.DataBind.SetBound(true, "", "dsHasta");

            // Boton para obtener informacion
            oItem = oForm.Items.Add("Obtener", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = Left + oForm.Items.Item("stDesde").Width + oForm.Items.Item("etDesde").Width + oForm.Items.Item("stHasta").Width + oForm.Items.Item("etHasta").Width + 40;
            oItem.Width = 120;
            oItem.Top = Top;
            oItem.Height = 19;
            SAPbouiCOM.Button oButton = ((SAPbouiCOM.Button)(oItem.Specific));
            oButton.Caption = "Obtener Documentos";

            Top += 20;
            // static Socio de Negocio
            oItem = oForm.Items.Add("stSN", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            itemST = ((SAPbouiCOM.StaticText)(oItem.Specific));
            itemST.Item.Top = Top;
            itemST.Item.Width = 80;
            itemST.Item.Left = Left;
            itemST.Caption = "Socio Negocio:";

            // Campo Socio de Negocio
            oItem = oForm.Items.Add("etSN", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = Top;
            oItem.Left = Left + oForm.Items.Item("stSN").Width;
            oItem.Width = 110;
            oEditText = ((SAPbouiCOM.EditText)oItem.Specific);
            // Choose From list Socio de negocio
            FuncionesComunes.CrearCFL_SocioNegocio(oForm, oEditText);

            // static Tipo DTE
            oItem = oForm.Items.Add("stTipo", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            itemST = ((SAPbouiCOM.StaticText)(oItem.Specific));
            itemST.Item.Top = Top;
            itemST.Item.Width = 80;
            oItem.Left = Left + oForm.Items.Item("stSN").Width + oForm.Items.Item("etSN").Width + 20;
            itemST.Caption = "Documentos:";

            // Campo  Tipo DTE
            oItem = oForm.Items.Add("etTipo", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oItem.Top = Top;
            oItem.Left = Left + oForm.Items.Item("stSN").Width + oForm.Items.Item("etSN").Width + oForm.Items.Item("stTipo").Width + 20;
            oItem.Width = 110;
            SAPbouiCOM.ComboBox oComboBox = ((SAPbouiCOM.ComboBox)oItem.Specific);
            oComboBox.ValidValues.Add("", "");
            oComboBox.ValidValues.Add("33", "Factura Electrónica");
            oComboBox.ValidValues.Add("34", "Factura Exenta de IVA Electrónica");
            oComboBox.ValidValues.Add("56", "Nota de Débito Electrónica");
            oComboBox.ValidValues.Add("61", "Nota de Crédito Electrónica");
            
            Top += 20;
            // Matrix para mostrar documentos
            oItem = oForm.Items.Add("oMtx", SAPbouiCOM.BoFormItemTypes.it_MATRIX);            
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)oItem.Specific);
            oMatrix.Item.Top = Top;
            oMatrix.Item.Left = Left;
            oMatrix.Item.Height = 340;
            oMatrix.Item.Width = 1100;

            // Columnas para Rut Emisor, Razón social emisor, Tipo Doc, folio, monto total, estado y Razon de reparo y/o rechazo
            SAPbouiCOM.Columns oColumns;
            oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;

            oMatrix.Layout = SAPbouiCOM.BoMatrixLayoutType.mlt_Normal;
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
            
            // Datasource para blindear matrix
            oForm.DataSources.DataTables.Add("DOCUMENTOS");
            SAPbouiCOM.DataTable dt = oForm.DataSources.DataTables.Item("DOCUMENTOS");

            dt.Columns.Add("co_FebId", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Rut", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_RznSoc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Tipo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Folio", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Fecha", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_FechaR", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Total", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Estado", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Pago", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);            
            dt.Columns.Add("co_Exissn", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);            
            dt.Columns.Add("co_Card", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                        
            oColumn = oColumns.Add("co_Rut", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Rut Emisor";
            oColumn.Editable = false;            
            oColumn.Width = 70;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Rut");            

            oColumn = oColumns.Add("co_RznSoc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Razón Social";
            oColumn.Editable = false;
            oColumn.Width = 250;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_RznSoc");

            oColumn = oColumns.Add("co_Tipo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Tipo Documento";
            oColumn.Editable = false;
            oColumn.Width = 130;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Tipo");

            oColumn = oColumns.Add("co_Folio", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Folio";
            oColumn.Editable = false;
            oColumn.Width = 60;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Folio");

            oColumn = oColumns.Add("co_Fecha", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Fecha Emisión";
            oColumn.Editable = false;
            oColumn.Width = 90;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Fecha");

            oColumn = oColumns.Add("co_FechaR", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Fecha Recep SII";
            oColumn.Editable = false;
            oColumn.Width = 90;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_FechaR");

            oColumn = oColumns.Add("co_Total", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Monto Total";
            oColumn.Editable = false;
            oColumn.Width = 80;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Total");

            oColumn = oColumns.Add("co_Estado", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Estado SII";
            oColumn.Editable = false;
            oColumn.Width = 140;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Estado");
            
            oColumn = oColumns.Add("co_Pago", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Pago";
            oColumn.Editable = false;
            oColumn.Width = 60;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Pago");
            
            oColumn = oColumns.Add("co_Exissn", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Existe SN";
            oColumn.Editable = false;
            oColumn.ValOn = "Y";
            oColumn.ValOff = "N";
            oColumn.Width = 70;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Exissn");
            
            oColumn = oColumns.Add("co_Card", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Cardcode";
            oColumn.Visible = false;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Card");
            
            oColumn = oColumns.Add("co_FebId", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Febos ID";
            oColumn.Visible = false;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_FebId"); 
               
            oForm.Freeze(false);
        }

        /// <summary>
        /// Controla los eventos de items del formularios
        /// </summary>
        public static void m_SBO_Appl_ItemEvent(String FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = null;
            oForm = Conexion_SBO.m_SBO_Appl.Forms.Item(FormUID);
            SAPbouiCOM.Matrix oMatrix = null;
            ResultMessage result = new ResultMessage();
            Int32 Row = 0;
            String MensajeMsg = String.Empty;
            SAPbouiCOM.ComboBox oCombo = null;

            try
            {
                if (pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CLICK) && pVal.ItemUID.Equals("bt_Proc"))
                {
                    oCombo = oForm.Items.Item("cb_Options").Specific;
                    if (oCombo.Selected != null)
                    {
                        #region Integración
                        if (oCombo.Selected.Value.Equals("1"))
                        {
                            oMatrix = oForm.Items.Item("oMtx").Specific;
                            Row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
                            if (Row != -1)
                            {
                                GestionarIntegracion(oForm, Row, 1);
                            }
                            else
                            {
                                BubbleEvent = false;
                                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Seleccione un documento para integrar documento.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            }
                        }
                        #endregion  
                        #region Integración
                        if (oCombo.Selected.Value.Equals("2"))
                        {
                            oMatrix = oForm.Items.Item("oMtx").Specific;
                            Row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
                            if (Row != -1)
                            {
                                GestionarIntegracion(oForm, Row, 2);                                
                            }
                            else
                            {
                                BubbleEvent = false;
                                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Seleccione un documento para integrar documento error contado.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            }
                        }
                        #endregion  
                        #region Visualizar
                        else if (oCombo.Selected.Value.Equals("3"))
                        {
                            oMatrix = oForm.Items.Item("oMtx").Specific;
                            Row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
                            if (Row != -1)
                            {
                                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Visualizando 1 documento. Espere un momento.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                result = GestionarVisualizar(oForm, Row);

                                if (!result.Success) Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }
                            else
                            {
                                BubbleEvent = false;
                                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Seleccione un documento para visualizar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            }
                        }
                        #endregion                        
                        #region LLenar SN
                        else if (oCombo.Selected.Value.Equals("4"))
                        {
                            oMatrix = oForm.Items.Item("oMtx").Specific;
                            Row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
                            if (Row != -1)
                            {
                                LLenarSocioNegocio(oForm, Row);
                            }
                            else
                            {
                                BubbleEvent = false;
                                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Seleccione un documento para llenar Socio de negocio", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            }
                        }
                        #endregion                        
                    }
                    else
                    {
                        BubbleEvent = false;
                        Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Seleccione una opción para procesar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                }
                if (pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CLICK) && pVal.ItemUID.Equals("Obtener"))
                {
                    GestionarDescarga(oForm.Items.Item("oMtx").Specific, oForm);
                }
                if (pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CLICK) && pVal.ItemUID.Equals("bt_Salir"))
                {
                    ListaFilas.ListFilas.Clear();
                    oForm.Close();
                }
                if (!pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST))
                {
                    GestionarSeleccionCFL(oForm, pVal);
                }
                if (!pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) && pVal.ItemUID.Equals("cb_Options"))
                {
                    oCombo = oForm.Items.Item("cb_Options").Specific;
                    if (oCombo.Selected.Value.Equals("1") || oCombo.Selected.Value.Equals("2"))
                    {
                        oMatrix = oForm.Items.Item("oMtx").Specific;
                        Row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
                        if (Row != -1)
                        {   
                            listarFilas(Row, oForm.Items.Item("oMtx").Specific);
                            SEI_FormDat oFormDat = new SEI_FormDat("SEI_DAT");                                                     
                        }
                        else
                        {
                            Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Seleccione un documento para integrar documento", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                    }
                }
            }            
            catch(Exception ex)
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Obtiene los documentos desde febos y los muestra en grilla
        /// </summary>
        public static void GestionarDescarga(SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.Form oForm)
        {
            ListaDTEMatrix.ListaDTE.Clear();
            Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Descargando información. Espere unos momentos", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            oForm.Freeze(true);

            SAPbouiCOM.DataTable dtDoc = oForm.DataSources.DataTables.Item("DOCUMENTOS");

            try
            {
                dtDoc.Rows.Clear();
                oMatrix.Clear();

                String[] Ruts = new String[] { };
                String FechaFinal = String.Empty;
                String FechaInicial = String.Empty;
                String Filtros = String.Empty;

                // armar filtro de consulta
                // Rut receptor para documentos recibidos y estado comercial sin estado
                Filtros = String.Format("rutReceptor:{0}|estadoSii:4,5|estadoComercial:2,7|incompleto:N", FuncionesComunes.ObtenerRut()); // |formaDePago:1
                
                // Tipo de documento
                SAPbouiCOM.ComboBox oCombobox = oForm.Items.Item("etTipo").Specific;
                if (!String.IsNullOrEmpty(oCombobox.Selected.Value))
                {
                    Filtros += String.Format("|tipoDocumento:{0}", oCombobox.Selected.Value);
                }

                // Socios de negocios
                if (!String.IsNullOrEmpty(oForm.Items.Item("etSN").Specific.Value))
                {
                    String[] ArrayRuts = oForm.Items.Item("etSN").Specific.Value.Split(';');
                    Filtros += String.Format("|rutEmisor:");
                    foreach (String rut in ArrayRuts)
                    {
                        Filtros += String.Format("{0},", rut);
                    }
                    Filtros = Filtros.Substring(0, Filtros.Length - 1);
                }

                // Rango de fechas
                String DesdeFecha = oForm.Items.Item("etDesde").Specific.Value;
                String HastaFecha = oForm.Items.Item("etHasta").Specific.Value;

                if (!String.IsNullOrEmpty(DesdeFecha) && !String.IsNullOrEmpty(HastaFecha))
                {
                    FechaInicial = String.Format("{0}-{1}-{2}", DesdeFecha.Substring(0, 4), DesdeFecha.Substring(4, 2), DesdeFecha.Substring(6, 2));
                    FechaFinal = String.Format("{0}-{1}-{2}", HastaFecha.Substring(0, 4), HastaFecha.Substring(4, 2), HastaFecha.Substring(6, 2));
                }
                else
                {
                    DateTime dt;
                    String Mes = String.Empty;
                    String Dia = String.Empty;

                    dt = new DateTime(DateTime.Now.Year, DateTime.Today.Month, DateTime.Today.Day);
                    Mes = dt.Month.ToString().PadLeft(2, '0');
                    Dia = dt.Day.ToString().PadLeft(2, '0');

                    FechaFinal = String.Format("{0}-{1}-{2}", dt.Year.ToString(), Mes, Dia);

                    dt = DateTime.Today.AddDays(-8);
                    Mes = dt.Month.ToString().PadLeft(2, '0');
                    Dia = dt.Day.ToString().PadLeft(2, '0');

                    FechaInicial = String.Format("{0}-{1}-{2}", dt.Year.ToString(), Mes, Dia);
                }
                Filtros += String.Format("|fechaRecepcion:{0}--{1}", FechaInicial, FechaFinal);

                var clientDescarga = new RestClient();
                var requestDescarga = new RestRequest(ConfigurationManager.AppSettings["Descarga"], Method.GET);

                requestDescarga.RequestFormat = DataFormat.Json;
                requestDescarga.AddHeader("token", FuncionesComunes.ObtenerToken());
                requestDescarga.AddHeader("empresa", FuncionesComunes.ObtenerRut());
                requestDescarga.AddHeader("pagina", "1");
                requestDescarga.AddHeader("itemsPorPagina", "200");
                requestDescarga.AddHeader("campos", "tipoDocumento,folio,fechaEmision,fechaRecepcion,formaDePago,rutEmisor,razonSocialEmisor,montoTotal,estadoSii,plazo,fechaRecepcionSii");
                requestDescarga.AddHeader("filtros", Filtros);
                requestDescarga.AddHeader("orden", "-fechaRecepcion");

                IRestResponse responseDescarga = clientDescarga.Execute(requestDescarga);

                if (responseDescarga.StatusDescription.Equals("OK"))
                {
                    RootObjectDescarga d = JsonConvert.DeserializeObject<RootObjectDescarga>(responseDescarga.Content);
                    Int32 IndexMatrix = 0;
                    String CardCode = String.Empty;
                    String DocEntryBase = String.Empty;
                    String Query = String.Empty;

                    foreach (Comunes.Documento dto in d.documentos)
                    {
                        if (!dto.tipoDocumento.Equals(39) && !dto.tipoDocumento.Equals(41) && !dto.tipoDocumento.Equals(52))
                        {
                            dtDoc.Rows.Add();
                            dtDoc.SetValue("co_FebId", IndexMatrix, dto.febosId);
                            dtDoc.SetValue("co_Rut", IndexMatrix, dto.rutEmisor);
                            dtDoc.SetValue("co_RznSoc", IndexMatrix, dto.razonSocialEmisor);
                            dtDoc.SetValue("co_Tipo", IndexMatrix, FuncionesComunes.ObtenerTipoDocumento(dto.tipoDocumento.ToString()));
                            dtDoc.SetValue("co_Folio", IndexMatrix, dto.folio);
                            dtDoc.SetValue("co_Fecha", IndexMatrix, dto.fechaEmision);
                            dtDoc.SetValue("co_FechaR", IndexMatrix, dto.fechaRecepcion.Substring(0, 10));
                            dtDoc.SetValue("co_Total", IndexMatrix, String.Format(System.Globalization.CultureInfo.GetCultureInfo("es-CL"), "{0:C0}", dto.montoTotal));
                            dtDoc.SetValue("co_Estado", IndexMatrix, FuncionesComunes.ObtenerEstadoSII(dto.estadoSii));

                            if (dto.formaDePago != null)
                            {
                                dtDoc.SetValue("co_Pago", IndexMatrix, FuncionesComunes.ObtenerFormaPago(dto.formaDePago.ToString()));
                            }
                            else
                            {
                                dtDoc.SetValue("co_Pago", IndexMatrix, "No informado");
                            }
                            String ProveedorNR;
                            CardCode = FuncionesComunes.ObtenerCardcode(dto.rutEmisor, out ProveedorNR);

                            if (!String.IsNullOrEmpty(CardCode))
                            {
                                dtDoc.SetValue("co_Exissn", IndexMatrix, "Y");
                                dtDoc.SetValue("co_Card", IndexMatrix, CardCode);
                                dtDoc.SetValue("co_ProvNR", IndexMatrix, ProveedorNR);
                            }
                            else
                            {
                                dtDoc.SetValue("co_Exissn", IndexMatrix, "N");
                            }

                            // Descargar XML
                            var clientGetXML = new RestClient();
                            var requestGetXML = new RestRequest(String.Format(ConfigurationManager.AppSettings["GetXML"], dto.febosId), Method.GET);

                            requestGetXML.RequestFormat = DataFormat.Json;
                            requestGetXML.AddHeader("token", FuncionesComunes.ObtenerToken());
                            requestGetXML.AddHeader("empresa", FuncionesComunes.ObtenerRut());
                            requestGetXML.AddHeader("xml", "si");
                            requestGetXML.AddHeader("xmlFirmado", "si");
                            requestGetXML.AddHeader("incrustar", "si");

                            IRestResponse responseGetXML = clientGetXML.Execute(requestGetXML);

                            if (responseGetXML.StatusDescription.Equals("OK"))
                            {
                                RootObjectGetXML x = JsonConvert.DeserializeObject<RootObjectGetXML>(responseGetXML.Content);
                                byte[] datos = Convert.FromBase64String(x.xmlData);
                                Encoding iso = Encoding.GetEncoding("ISO-8859-1");
                                String DecodeString = iso.GetString(datos);

                                ResultMessage result = FuncionesComunes.ObtenerDTE(DecodeString);

                                if (result.Success)
                                {
                                    DTE objDTE = (DTE)result.DTE;                                                                        
                                    DTEMatrix objDteMatrix = new DTEMatrix();
                                    objDteMatrix.FebosID = dto.febosId;
                                    objDteMatrix.objDTE = objDTE;
                                    ListaDTEMatrix.ListaDTE.Add(objDteMatrix);
                                }
                            }

                            IndexMatrix++;
                        }
                    }

                    oMatrix.LoadFromDataSource();
                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Información descargada correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Obtiene el PDF desde febos de los documentos seleccionados en grilla y levante el PDF en programa predeterminado.
        /// </summary>
        public static ResultMessage GestionarVisualizar(SAPbouiCOM.Form oForm, Int32 Row)
        {
            ResultMessage result = new ResultMessage();

            try
            {
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("oMtx").Specific;
                String FebosId = oMatrix.Columns.Item("co_FebId").Cells.Item(Row).Specific.Value;

                // Descargar PDF
                var clientGetPDF = new RestClient();
                var requestGetPDF = new RestRequest(String.Format(ConfigurationManager.AppSettings["GetXML"], FebosId), Method.GET);

                requestGetPDF.RequestFormat = DataFormat.Json;
                requestGetPDF.AddHeader("token", FuncionesComunes.ObtenerToken());
                requestGetPDF.AddHeader("empresa", FuncionesComunes.ObtenerRut());
                requestGetPDF.AddHeader("imagen", "si");
                requestGetPDF.AddHeader("tipoImagen", "0");

                IRestResponse responseGetPDF = clientGetPDF.Execute(requestGetPDF);

                if (responseGetPDF.StatusDescription.Equals("OK"))
                {
                    RootObjectGetXML x = JsonConvert.DeserializeObject<RootObjectGetXML>(responseGetPDF.Content);

                    String link = x.imagenLink;

                    ProcessStartInfo info = new ProcessStartInfo();
                    info.FileName = link;
                    info.Verb = "open";
                    info.WindowStyle = ProcessWindowStyle.Maximized;
                    Process.Start(info);
                }
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Mensaje = ex.Message;
                return result;
            }
        }

        /// <summary>
        /// Asigna el rut del socio de negocio al campo socio de negocio, luego de seleccionarlo en el CFL
        /// </summary>
        public static void GestionarSeleccionCFL(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
            String CflID = oCFLEvento.ChooseFromListUID;
            SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(CflID);

            SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;
            if (oDataTable != null)
            {
                oForm.DataSources.UserDataSources.Item("dsSN").Value = String.Empty;
                for (Int32 Row = 0; Row <= (oDataTable.Rows.Count - 1); Row++)
                {                    
                    String LicTradNum = String.Empty;
                    LicTradNum = oDataTable.GetValue(23, Row).ToString();
                    oForm.DataSources.UserDataSources.Item("dsSN").Value += LicTradNum;
                    if (Row < (oDataTable.Rows.Count - 1))
                    {
                        oForm.DataSources.UserDataSources.Item("dsSN").Value += ";";
                    }
                }
            }
        }
                
        /// <summary>
        /// Integra el documento de proveedor seleccionado a SAP, ya sea por articulo o servicio.
        /// TipoIntegracion: 1 -> integrar 2 -> integrar error contado
        /// </summary>
        public static void GestionarIntegracion(SAPbouiCOM.Form oForm, Int32 Row, Int32 TipoIntegracion)
        {

            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("oMtx").Specific;                       
            String RutEmisor = String.Empty;
            String RznSocial = String.Empty;
            String Tipo = String.Empty;
            String Folio = String.Empty;
            String FebId = String.Empty;
            String CardCode = String.Empty;
                         
            DTE objDTE = null;

            RutEmisor = oMatrix.Columns.Item("co_Rut").Cells.Item(Row).Specific.Value;
            RznSocial = oMatrix.Columns.Item("co_RznSoc").Cells.Item(Row).Specific.Value;
            Tipo = FuncionesComunes.ObtenerTipoDocumentoNumero(oMatrix.Columns.Item("co_Tipo").Cells.Item(Row).Specific.Value);
            Folio = oMatrix.Columns.Item("co_Folio").Cells.Item(Row).Specific.Value;
            FebId = oMatrix.Columns.Item("co_FebId").Cells.Item(Row).Specific.Value;
            CardCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("co_Card").Cells.Item(Row).Specific).Value;

            if (FebId.Equals(ListaFilas.ListFilas[0].FebosId))
            {
                ResultMessage rslt = FuncionesComunes.ValidacionDTEIntegrado(RutEmisor, Int32.Parse(Tipo), Int64.Parse(Folio));

                if (rslt.Success)
                {
                    objDTE = ListaDTEMatrix.ListaDTE.Where(i => i.FebosID == FebId).Select(i => i.objDTE).SingleOrDefault();

                    if (objDTE != null)
                    {
                        try
                        {
                            Conexion_SBO.m_SBO_Appl.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_PurchaseInvoice, String.Empty, String.Empty);
                            SAPbouiCOM.Form oFormPI = null;
                            oFormPI = Conexion_SBO.m_SBO_Appl.Forms.ActiveForm;
                            oFormPI.Freeze(true);
                            oFormPI.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            oFormPI.Items.Item("4").Specific.Value = CardCode;
                            oFormPI.Items.Item("208").Specific.Value = objDTE.IdDoc.TipoDTE;
                            oFormPI.Items.Item("211").Specific.Value = objDTE.IdDoc.Folio;
                            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oFormPI.Items.Item("3").Specific;
                            oCombo.Select("S", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oFormPI.Items.Item("10").Specific.Value = objDTE.IdDoc.FchEmis.Replace("-", "");
                            oFormPI.Items.Item("46").Specific.Value = objDTE.IdDoc.FchEmis.Replace("-", "");
                            if (!String.IsNullOrEmpty(objDTE.IdDoc.FchVenc))
                            {
                                oFormPI.Items.Item("12").Specific.Value = objDTE.IdDoc.FchVenc.Replace("-", "");
                            }
                                                        
                            SAPbouiCOM.Form oFormUF = null;
                            for (Int32 index = 0; index <= Conexion_SBO.m_SBO_Appl.Forms.Count; index++)
                            {
                                oFormUF = Conexion_SBO.m_SBO_Appl.Forms.Item(index);
                                if (oFormUF.TypeEx.Equals("-141") && oFormUF.UniqueID.Equals(oFormPI.UDFFormUID))
                                {
                                    break;
                                }
                            }
                            
                            SAPbouiCOM.EditText oEditText = oFormUF.Items.Item("U_SEI_FEBOSID").Specific;
                            oEditText.Value = FebId;
                            
                            if (TipoIntegracion.Equals(2))
                            {
                                SAPbouiCOM.ComboBox oComboBox = oFormUF.Items.Item("U_SEI_CONTADO").Specific;
                                oComboBox.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            }                            

                            SAPbouiCOM.Matrix oMtx = oFormPI.Items.Item("39").Specific;
                            Int32 Index = 1;
                            foreach (Detalle det in objDTE.Detalle)
                            {
                                if (det.IndExe.Equals(2) || det.IndExe.Equals(3) || det.IndExe.Equals(4) || det.IndExe.Equals(5) || det.IndExe.Equals(6))
                                {
                                    oMtx.Columns.Item("1").Cells.Item(Index).Specific.Value = det.NmbItem + det.DscItem + det.MontoItem.ToString();
                                    oMtx.Columns.Item("12").Cells.Item(Index).Specific.Value = 0;
                                    oMtx.Columns.Item("5").Cells.Item(Index).Specific.Value = 0;
                                }
                                else
                                {
                                    if (det.NmbItem.Equals(det.DscItem) || String.IsNullOrEmpty(det.DscItem))
                                    {
                                        oMtx.Columns.Item("1").Cells.Item(Index).Specific.Value = det.NmbItem;
                                    }
                                    else
                                    {
                                        if ((det.NmbItem + det.DscItem).Length > 100)
                                        {
                                            oMtx.Columns.Item("1").Cells.Item(Index).Specific.Value = (det.NmbItem + det.DscItem).Substring(0, 100);
                                        }
                                        else
                                        {
                                            oMtx.Columns.Item("1").Cells.Item(Index).Specific.Value = (det.NmbItem + det.DscItem);
                                        }
                                    }

                                    oMtx.Columns.Item("2").Cells.Item(Index).Specific.Value = ListaFilas.ListFilas.First(d => d.LineNum == det.NroLinDet).Cuenta;
                                    oMtx.Columns.Item("2004").Cells.Item(Index).Specific.Value = ListaFilas.ListFilas.First(d => d.LineNum == det.NroLinDet).Especie;
                                    oMtx.Columns.Item("2003").Cells.Item(Index).Specific.Value = ListaFilas.ListFilas.First(d => d.LineNum == det.NroLinDet).Variedad;
                                    oMtx.Columns.Item("2002").Cells.Item(Index).Specific.Value = ListaFilas.ListFilas.First(d => d.LineNum == det.NroLinDet).Category;
                                    oMtx.Columns.Item("2001").Cells.Item(Index).Specific.Value = ListaFilas.ListFilas.First(d => d.LineNum == det.NroLinDet).CatPacking;
                                    oMtx.Columns.Item("2000").Cells.Item(Index).Specific.Value = ListaFilas.ListFilas.First(d => d.LineNum == det.NroLinDet).RolPrivado;
                                    oMtx.Columns.Item("12").Cells.Item(Index).Specific.Value = det.MontoItem;
                                    oMtx.Columns.Item("5").Cells.Item(Index).Specific.Value = det.PrcItem * det.QtyItem;
                                    oMtx.Columns.Item("6").Cells.Item(Index).Specific.Value = det.DescuentoPct;


                                    if (det.CodImpAdic.Count > 0)
                                    {
                                        String QueryImp = null;

                                        switch (Conexion_SBO.m_oCompany.DbServerType)
                                        {
                                            case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                                                QueryImp = "SELECT \"Code\" FROM \"OSTA\" WHERE \"U_SEI_COIM\" = '" + det.CodImpAdic[0].sCodImpAdic + "'";
                                                break;
                                            default:
                                                QueryImp = "SELECT Code FROM OSTA WHERE U_SEI_COIM = '" + det.CodImpAdic[0].sCodImpAdic + "'";
                                                break;
                                        }

                                        SAPbobsCOM.Recordset oRec = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        oRec.DoQuery(QueryImp);
                                        oMtx.Columns.Item("95").Cells.Item(Index).Specific.Value = oRec.Fields.Item(0).Value;
                                    }
                                    else
                                    {
                                        if (det.IndExe.Equals(1))
                                        {
                                            oMtx.Columns.Item("95").Cells.Item(Index).Specific.Value = "IVA_EXE";
                                        }
                                        else
                                        {
                                            oMtx.Columns.Item("95").Cells.Item(Index).Specific.Value = "IVA";
                                        }
                                    }
                                    if (Index < objDTE.Detalle.Count)
                                    {
                                        oMtx.AddRow();
                                    }
                                    Index++;
                                }
                            }

                            // Descuentos y/o recargos
                            if (objDTE.DscRcgGlobal.Count > 0)
                            {
                                foreach (DscRcgGlobal objDRG in objDTE.DscRcgGlobal)
                                {
                                    // Descuento
                                    if (objDRG.TpoMov.Equals("D"))
                                    {
                                        // Por porcentaje
                                        if (objDRG.TpoValor.Equals("%"))
                                        {
                                            oFormPI.Items.Item("24").Specific.Value = objDRG.ValorDR;
                                        }
                                        // Por valor
                                        else
                                        {
                                            Double SumMontoNeto = objDTE.Detalle.Where(i => i.IndExe == 0).Sum(i => i.MontoItem);
                                            Double Porcentaje = Math.Round((objDRG.ValorDR * 100 / SumMontoNeto), 2);
                                            String sPorcentaje = Porcentaje.ToString().Replace(',', '.');
                                            oFormPI.Items.Item("24").Specific.Value = sPorcentaje;
                                        }
                                    }
                                    // Recargo
                                    else
                                    {
                                        // Por porcentaje
                                        if (objDRG.TpoValor.Equals("%"))
                                        {
                                            oFormPI.Items.Item("24").Specific.Value = (objDRG.ValorDR * -1);
                                        }
                                        // Por valor
                                        else
                                        {
                                            Double SumMontoNeto = objDTE.Detalle.Where(i => i.IndExe == 0).Sum(i => i.MontoItem);
                                            Double Porcentaje = Math.Round((objDRG.ValorDR * 100 / SumMontoNeto), 2);
                                            String sPorcentaje = Porcentaje.ToString().Replace(',', '.');
                                            oFormPI.Items.Item("24").Specific.Value = "-" + sPorcentaje;
                                        }
                                    }


                                }
                            }

                            if (!objDTE.Totales.MontoNF.Equals(0))
                            {
                                oFormPI.Items.Item("105").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oFormPI.Items.Item("103").Specific.Value = objDTE.Totales.MontoNF;

                            }

                            oFormPI.Items.Item("4").Specific.Value = CardCode;
                            oFormPI.Freeze(false);
                            Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Factura de proveedores llenada de forma correcta. Complete la información.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                        catch (Exception ex)
                        {
                            Conexion_SBO.m_SBO_Appl.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                    }//
                }
                else
                {
                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText(String.Format("Error: El DTE {0} Tipo {1} de {2}-{3} ya se encuentra integrado.", Folio, Tipo, RznSocial, RutEmisor), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }
            else
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Documento seleccionado no coincide con documento asignado", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }
        
        /// <summary>
        /// Crear un socio de negocio en SAP del documento seleccionado en la grilla y levanta el maestro de SN creado.
        /// </summary>
        public static void LLenarSocioNegocio(SAPbouiCOM.Form oForm, Int32 Row)
        {
            Conexion_SBO.m_SBO_Appl.StatusBar.SetText("LLenando Socio de Negocio. Espere unos momentos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            SAPbouiCOM.Matrix oMatrix = null;
            Boolean checkSN = false;
            oMatrix = oForm.Items.Item("oMtx").Specific;

            checkSN = ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Exissn").Cells.Item(Row).Specific).Checked;
            if (!checkSN)
            {
                String FebosId = oMatrix.Columns.Item("co_FebId").Cells.Item(Row).Specific.Value;

                // Descargar XML
                var clientGetXML = new RestClient();
                var requestGetXML = new RestRequest(String.Format(ConfigurationManager.AppSettings["GetXML"], FebosId), Method.GET);

                requestGetXML.RequestFormat = DataFormat.Json;
                requestGetXML.AddHeader("token", FuncionesComunes.ObtenerToken());
                requestGetXML.AddHeader("empresa", FuncionesComunes.ObtenerRut());
                requestGetXML.AddHeader("xml", "si");
                requestGetXML.AddHeader("xmlFirmado", "si");
                requestGetXML.AddHeader("incrustar", "si");

                IRestResponse responseGetXML = clientGetXML.Execute(requestGetXML);

                if (responseGetXML.StatusDescription.Equals("OK"))
                {
                    RootObjectGetXML x = JsonConvert.DeserializeObject<RootObjectGetXML>(responseGetXML.Content);
                    byte[] datos = Convert.FromBase64String(x.xmlData);
                    Encoding iso = Encoding.GetEncoding("ISO-8859-1");
                    String DecodeString = iso.GetString(datos);

                    ResultMessage result = FuncionesComunes.ObtenerDTE(DecodeString);

                    if (result.Success)
                    {
                        DTE objDTE = (DTE)result.DTE;
                        Conexion_SBO.m_SBO_Appl.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_BusinessPartner, String.Empty, String.Empty);
                        SAPbouiCOM.Form oFormSN = Conexion_SBO.m_SBO_Appl.Forms.ActiveForm;
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                        String CardCode = String.Format("P{0}", objDTE.Emisor.RUTEmisor);

                        oFormSN.Items.Item("7").Specific.Value = objDTE.Emisor.RznSoc.ToLower();
                        oFormSN.Items.Item("128").Specific.Value = objDTE.Emisor.RznSoc.ToLower();
                        SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oFormSN.Items.Item("40").Specific;
                        oCombo.Select("S", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        oCombo = (SAPbouiCOM.ComboBox)oFormSN.Items.Item("16").Specific;
                        oCombo.Select("101", SAPbouiCOM.BoSearchKey.psk_ByValue);

                        if (objDTE.Emisor.GiroEmis.Length > 40)
                        {
                            oFormSN.Items.Item("U_SEI_GNRP").Specific.Value = objDTE.Emisor.GiroEmis.Substring(0, 40);
                        }
                        else
                        {
                            oFormSN.Items.Item("U_SEI_GNRP").Specific.Value = objDTE.Emisor.GiroEmis;
                        }

                        oFormSN.Items.Item("41").Specific.Value = objDTE.Emisor.RUTEmisor;
                        oFormSN.Items.Item("43").Specific.Value = objDTE.Emisor.Telefono;
                        oFormSN.Items.Item("60").Specific.Value = objDTE.Emisor.CorreoEmisor;
                        oFormSN.Items.Item("5").Specific.Value = CardCode;
                        Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Socio de Negocios llenado de forma correcta. Complete la información.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                }
                else
                {
                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText(responseGetXML.ErrorMessage, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }
            else
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("El Socio de negocio del documento seleccionado ya existe", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }
                
        /// <summary>
        /// Funcion que detecta los detalles de la factura, y llena clase generica para luego mostrar por pantalla a usuario
        /// quien por cada linea de detalle llena cuenta, especie, variedad y category.
        /// </summary>
        public static void listarFilas(Int32 Row, SAPbouiCOM.Matrix oMatrix)
        {
            ListaFilas.ListFilas.Clear();
            String FebosId = oMatrix.Columns.Item("co_FebId").Cells.Item(Row).Specific.Value;

            var clientGetXML = new RestClient();
            var requestGetXML = new RestRequest(String.Format(ConfigurationManager.AppSettings["GetXML"], FebosId), Method.GET);

            requestGetXML.RequestFormat = DataFormat.Json;
            requestGetXML.AddHeader("token", FuncionesComunes.ObtenerToken());
            requestGetXML.AddHeader("empresa", FuncionesComunes.ObtenerRut());
            requestGetXML.AddHeader("xml", "si");
            requestGetXML.AddHeader("xmlFirmado", "si");
            requestGetXML.AddHeader("incrustar", "si");

            IRestResponse responseGetXML = clientGetXML.Execute(requestGetXML);

            if (responseGetXML.StatusDescription.Equals("OK"))
            {
                RootObjectGetXML x = JsonConvert.DeserializeObject<RootObjectGetXML>(responseGetXML.Content);
                byte[] datos = Convert.FromBase64String(x.xmlData);
                Encoding iso = Encoding.GetEncoding("ISO-8859-1");
                String DecodeString = iso.GetString(datos);

                ResultMessage result = FuncionesComunes.ObtenerDTE(DecodeString);

                if (result.Success)
                {
                    DTE objDTE = (DTE)result.DTE;

                    foreach (Detalle det in objDTE.Detalle)
                    {
                        Filas fila = new Filas();
                        fila.FebosId = FebosId;
                        fila.LineNum = det.NroLinDet;
                        fila.Servicio = det.NmbItem;
                        fila.Total = det.MontoItem;
                        if (det.CodImpAdic.Count > 0)
                        {
                            String QueryImp = null;

                            switch (Conexion_SBO.m_oCompany.DbServerType)
                            {
                                case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                                    QueryImp = "SELECT \"Code\" FROM \"OSTA\" WHERE \"U_SEI_COIM\" = '" + det.CodImpAdic[0].sCodImpAdic + "'";
                                    break;
                                default:
                                    QueryImp = "SELECT Code FROM OSTA WHERE U_SEI_COIM = '" + det.CodImpAdic[0].sCodImpAdic + "'";
                                    break;
                            }

                            SAPbobsCOM.Recordset oRec = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            oRec.DoQuery(QueryImp);
                            fila.Impuesto = oRec.Fields.Item(0).Value;
                        }
                        else
                        {
                            if (det.IndExe.Equals(1))
                            {
                                fila.Impuesto = "IVA_EXE";
                            }
                            else
                            {
                                fila.Impuesto = "IVA";
                            }
                        }
                        ListaFilas.ListFilas.Add(fila);
                    }
                }
            }
        }

        
    }


}
