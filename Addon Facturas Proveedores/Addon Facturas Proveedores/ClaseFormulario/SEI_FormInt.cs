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
    public class SEI_FormIntegracion
    {
        public static SAPbouiCOM.DataTable dt = null;
        public SEI_FormIntegracion(String FormUID)
        {
            CargarXML();
            CargarFormulario(FormUID);
        }


        /*
        private static void CargarXML(String FormUID)
        {
            XmlDocument oXmlDoc = null;
            SAPbouiCOM.FormCreationParams creationPackage = null;
            Boolean bFormAbierto = false;
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Form oFormAbierto = null;
            Int32 index;
            String xml = String.Empty;

            try
            {
                for (index = 0; index < Conexion_SBO.m_SBO_Appl.Forms.Count; index++)
                {
                    oFormAbierto = Conexion_SBO.m_SBO_Appl.Forms.Item(index);
                    if (oFormAbierto.UniqueID.Equals(FormUID))
                    {
                        bFormAbierto = true;
                        break;
                    }
                }

                // Si el form no esta abierto se abre
                if (!bFormAbierto)
                {
                    oXmlDoc = new XmlDocument();
                    xml = System.Windows.Forms.Application.StartupPath.ToString();
                    //xml += "\\";
                    xml = xml + "\\Formularios\\" + "FormInt.srf";
                    oXmlDoc.Load(xml);

                    xml = oXmlDoc.InnerXml.ToString();

                    creationPackage = (SAPbouiCOM.FormCreationParams)Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                    creationPackage.UniqueID = FormUID;
                    creationPackage.FormType = FormUID;
                    //creationPackage.Modality = SAPbouiCOM.BoFormModality.fm_Modal;
                    creationPackage.XmlData = xml;
                    oForm = Conexion_SBO.m_SBO_Appl.Forms.AddEx(creationPackage);
                    //oForm.Title = String.Format(oForm.Title);
                }
                // Si no traer al frente
                else
                {
                    Conexion_SBO.m_SBO_Appl.Forms.Item(index).Select();
                }
            }
            catch (Exception ex)
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("CargarXML > SEI_FormIntegracion.cs " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oForm);
                FuncionesComunes.LiberarObjetoGenerico(oFormAbierto);
            }
        }

        private static void CargarFormulario(String FormUID)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Matrix oMatrix = null;
            Int32 index = 1;
            try
            {
                oForm = Conexion_SBO.m_SBO_Appl.Forms.Item(FormUID);
                oForm.Freeze(true);
                oForm.Freeze(false);

            }
            catch
            {
                oForm.Freeze(false);

            }
            finally
            {
                oForm.Freeze(false);
                FuncionesComunes.LiberarObjetoGenerico(oForm);
                FuncionesComunes.LiberarObjetoGenerico(oMatrix);
            }
        }
        */



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
                    if (oFormAbierto.UniqueID.Equals("SEI_INT"))
                    {
                        bFormAbierto = true;
                        break;
                    }
                }

                // Si el form no esta abierto se abre
                if (!bFormAbierto)
                {
                    ////oForm= CargarFormulario("", Conexion_SBO.m_SBO_Appl, SAPbouiCOM.BoFormModality.fm_None);
                    oXmlDoc = new XmlDocument();

                    String xml = "";
                    xml = Application.StartupPath.ToString();
                    //xml += "\\";
                    xml += "\\Formularios\\" + "FormInt.srf";
                    oXmlDoc.Load(xml);

                    String sXML = oXmlDoc.InnerXml.ToString();
                    SAPbouiCOM.FormCreationParams creationPackage = null;
                    creationPackage = (SAPbouiCOM.FormCreationParams)Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                    creationPackage.UniqueID = "SEI_INT";
                    creationPackage.FormType = "SEI_INT";
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
                result = Msj_Appl.Errores(14, "CargarXML > SEI_FormInt.cs " + ex.Message);
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
      
        public static SAPbouiCOM.Form CargarFormulario(string pathXML, SAPbouiCOM.Application oAplicacion, SAPbouiCOM.BoFormModality modalidad)
        {

            string xml;
            string retVal = "";
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.FormCreationParams oCreationParams = null;

            try
            {
                oCreationParams = (SAPbouiCOM.FormCreationParams)oAplicacion.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                xml = Application.StartupPath.ToString();
                xml += "\\";
                xml = xml + "Formularios\\" + "FormInt.srf";
                oCreationParams.XmlData = xml;
                oCreationParams.Modality = modalidad;
                oCreationParams.UniqueID = "SEI_INT";
                oCreationParams.FormType = "SEI_INT";
                oForm = oAplicacion.Forms.AddEx(oCreationParams);
                retVal = oAplicacion.GetLastBatchResults();

            }
            catch (Exception e)
            {
                throw new Exception(retVal + " -- " + e.Message);
            }
            return oForm;
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
            oItem.Width = 150;
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
            //oComboBox.ValidValues.Add("", "");
            oComboBox.ValidValues.Add("33", "Factura Electrónica");
            oComboBox.ValidValues.Add("34", "Factura Exenta de IVA Electrónica");
            oComboBox.ValidValues.Add("56", "Nota de Débito Electrónica");
            oComboBox.ValidValues.Add("61", "Nota de Crédito Electrónica");
            oComboBox.Select("33", SAPbouiCOM.BoSearchKey.psk_ByValue);
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;


            Top += 20;
            // Matrix para mostrar documentos
            oItem = oForm.Items.Add("oMtx", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)oItem.Specific);
            oMatrix.Item.Top = Top;
            oMatrix.Item.Left = Left;
            oMatrix.Item.Height = 340;
            oMatrix.Item.Width = 1320;
            //Lucindo 08-02-2018
            // Columnas para Rut Emisor, Check Razón social emisor, Tipo Doc, folio, monto total, estado y Razon de reparo y/o rechazo
            SAPbouiCOM.Columns oColumns;
            oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;

            oMatrix.Layout = SAPbouiCOM.BoMatrixLayoutType.mlt_Normal;
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

            Top += 330;
            // static Razón Rechazo
            oItem = oForm.Items.Add("stRR", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Visible = false;
            itemST = ((SAPbouiCOM.StaticText)(oItem.Specific));
            itemST.Item.Top = Top;
            itemST.Item.Width = 80;
            itemST.Item.Left = Left;
            itemST.Caption = "Razón Rechazo:";

            // Campo Socio de Negocio
            oItem = oForm.Items.Add("etRR", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Top = Top;
            oItem.Left = Left + oForm.Items.Item("stRR").Width;
            oItem.Width = 300;
            oItem.Visible = false;
            oItem.AffectsFormMode = false;
            oEditText = ((SAPbouiCOM.EditText)oItem.Specific);

            // Datasource para blindear matrix
            oForm.DataSources.DataTables.Add("DOCUMENTOS");
            dt = oForm.DataSources.DataTables.Item("DOCUMENTOS");

            dt.Columns.Add("co_FebId", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Rut", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_RznSoc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Tipo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Folio", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Fecha", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_FechaR", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Total", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Estado", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Plazo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Pago", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Refoc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Exisoc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Foliooc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Totaloc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Exisocs", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Exisem", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Folioem", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Totalem", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Exissn", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_ProvNR", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Cat", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Card", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Base", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_CkeckSe", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

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

            oColumn = oColumns.Add("co_CkeckSe", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Seleccionado";
            oColumn.Editable = true;
            oColumn.Visible = true;
            oColumn.Width = 70;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_CkeckSe");

            oColumn = oColumns.Add("co_Exissn", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Existe SN";
            oColumn.Editable = false;
            oColumn.ValOn = "Y";
            oColumn.ValOff = "N";
            oColumn.Width = 70;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Exissn");

            oColumn = oColumns.Add("co_ProvNR", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Proveedor NR";
            oColumn.Editable = false;
            oColumn.ValOn = "Y";
            oColumn.ValOff = "N";
            oColumn.Width = 70;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_ProvNR");

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

            oColumn = oColumns.Add("co_Plazo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Plazo Rechazo";
            oColumn.Editable = false;
            oColumn.Width = 80;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Plazo");

            oColumn = oColumns.Add("co_Pago", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Pago";
            oColumn.Editable = false;
            oColumn.Width = 60;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Pago");

            oColumn = oColumns.Add("co_Refoc", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Ref. OC-FP";
            oColumn.Editable = false;
            oColumn.ValOn = "Y";
            oColumn.ValOff = "N";
            oColumn.Width = 70;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Refoc");

            oColumn = oColumns.Add("co_Exisoc", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Existe OC-FP";
            oColumn.Editable = false;
            oColumn.ValOn = "Y";
            oColumn.ValOff = "N";
            oColumn.Width = 70;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Exisoc");

            oColumn = oColumns.Add("co_Foliooc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Folio OC";
            oColumn.Editable = false;
            oColumn.Width = 80;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Foliooc");

            oColumn = oColumns.Add("co_Totaloc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Monto Total OC";
            oColumn.Editable = false;
            oColumn.Width = 80;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Totaloc");

            /*
            oColumn = oColumns.Add("co_Exisocs", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Existe OC-FP Serv";
            oColumn.Editable = false;
            oColumn.ValOn = "Y";
            oColumn.ValOff = "N";
            oColumn.Width = 90;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Exisocs");
            */

            oColumn = oColumns.Add("co_Exisem", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Existe EM";
            oColumn.Editable = false;
            oColumn.ValOn = "Y";
            oColumn.ValOff = "N";
            oColumn.Width = 70;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Exisem");

            oColumn = oColumns.Add("co_Folioem", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Folio EM";
            oColumn.Editable = false;
            oColumn.Width = 80;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Folioem");

            oColumn = oColumns.Add("co_Totalem", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Monto Total EM";
            oColumn.Editable = false;
            oColumn.Width = 80;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Totalem");

            /*
            oColumn = oColumns.Add("co_Cat", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Catalogo Art.";
            oColumn.Editable = false;
            oColumn.ValOn = "Y";
            oColumn.ValOff = "N";
            oColumn.Width = 70;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Cat");
            */

            oColumn = oColumns.Add("co_Card", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Cardcode";
            oColumn.Visible = false;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Card");

            oColumn = oColumns.Add("co_Base", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Base";
            oColumn.Visible = false;
            oColumn.DataBind.Bind("DOCUMENTOS", "co_Base");

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
            Int32 RespuestaMsg = 0;
            String MensajeMsg = String.Empty;
            SAPbouiCOM.ComboBox oCombo = null;

            try
            {
                if (pVal.BeforeAction)
                {

                    if (pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CLICK) && pVal.ItemUID.Equals("bt_Proc"))
                    {
                        oCombo = oForm.Items.Item("cb_Options").Specific;
                        if (oCombo.Selected != null)
                        {
                            #region Integrar

                            if (oCombo.Selected.Value.Equals("1"))
                            {
                                oMatrix = oForm.Items.Item("oMtx").Specific;
                                Row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
                                if (Row != -1)
                                {
                                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Integrando 1 documento. Espere un momento.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                    //result = GestionarVisualizar(oForm, Row);

                                    GestionarIntegracion(oForm, Row);

                                    //if (!result.Success) Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                }
                                else
                                {                                   
                                        BubbleEvent = false;
                                        Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Seleccione un documento para integrar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                }
                            }
                            else if (oCombo.Selected.Value.Equals("4"))
                            {

                                GestionarIntegracionS(oForm);
                            }
                            else if (oCombo.Selected.Value.Equals("5"))
                            {
                                GestionarIntegracionSV(oForm);
                            }

                            #endregion

                            #region Rechazar
                            else if (oCombo.Selected.Value.Equals("2"))
                            {
                                oMatrix = oForm.Items.Item("oMtx").Specific;
                                Row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
                                if (Row != -1)
                                {
                                    String RazonRechazo = oForm.Items.Item("etRR").Specific.Value;
                                    if (!String.IsNullOrEmpty(RazonRechazo))
                                    {
                                        MensajeMsg = String.Format("Se rechazará documento. ¿ Desea continuar ?");
                                        RespuestaMsg = Conexion_SBO.m_SBO_Appl.MessageBox(MensajeMsg, 1, "Continuar", "Cancelar");

                                        if (RespuestaMsg.Equals(1))
                                        {
                                            Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Rechazando documento. Espere un momento.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                            result = GestionarRechazo(oForm, RazonRechazo, Row);
                                            if (result.Success)
                                            {
                                                RazonRechazoVisible(oForm, false);
                                                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                                dt.Rows.Remove(Row - 1);
                                                oMatrix.LoadFromDataSource();
                                            }
                                            else
                                            {
                                                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        BubbleEvent = false;
                                        Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Ingrese razón de rechazo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                    }
                                }
                                else
                                {
                                    BubbleEvent = false;
                                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Seleccione un documento para rechazar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
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
                                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Prellenando SN de 1 documento. Espere un momento.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                    LLenarSocioNegocio(oForm, Row);
                                }
                                else
                                {
                                    BubbleEvent = false;
                                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Seleccione un documento para llenar Socio de negocio", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                } 
                            }
                            #endregion

                            #region LLenar OC
                            else if (oCombo.Selected.Value.Equals("5"))
                            {
                                oMatrix = oForm.Items.Item("oMtx").Specific;
                                Row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
                                if (Row != -1)
                                {                                
                                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Prellenando orden de compra de 1 documento. Espere un momento.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                    LLenarOrdenCompra(oForm, Row);
                                }                            
                                else
                                {
                                    BubbleEvent = false;
                                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Seleccione un documento para llenar Orden de Compra", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
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
                        //GestionarIntegracion(oForm);
                    }               
                    if (pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CLICK) && pVal.ItemUID.Equals("bt_Salir"))
                    {
                        oForm.Close();
                    }
                }
                if (!pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST))
                {
                    GestionarSeleccionCFL(oForm, pVal);
                }
                if (!pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) && pVal.ItemUID.Equals("cb_Options") && pVal.ItemChanged)
                {
                    oCombo = oForm.Items.Item("cb_Options").Specific;
                    if (oCombo.Selected != null)
                    {
                        if (oCombo.Selected.Value.Equals("2"))
                        {
                            RazonRechazoVisible(oForm, true);
                        }                        
                        else
                        {
                            RazonRechazoVisible(oForm, false);
                        }
                    }
                    else
                    {
                        RazonRechazoVisible(oForm, false);
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
            string strMensaje = "";
            List<Referencia> RefOC = null;
            List<Referencia> RefEM = null;

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
                Filtros = String.Format("rutReceptor:{0}|estadoComercial:0|incompleto:N", FuncionesComunes.ObtenerRut());

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
                    Filtros += String.Format("|fechaEmision:{0}--{1}", FechaInicial, FechaFinal);
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
                    Filtros += String.Format("|fechaEmision:{0}--{1}", FechaInicial, FechaFinal);
                    //Filtros += String.Format("|fechaRecepcion:{0}--{1}", FechaInicial, FechaFinal);
                }

                var clientDescarga = new RestClient();
                var requestDescarga = new RestRequest(ConfigurationManager.AppSettings["Descarga"], Method.GET);

                String nroRegistros = ConfigurationManager.AppSettings["NroRegistros"];

                requestDescarga.RequestFormat = DataFormat.Json;
                requestDescarga.AddHeader("token", FuncionesComunes.ObtenerToken());
                requestDescarga.AddHeader("empresa", FuncionesComunes.ObtenerRut());
                requestDescarga.AddHeader("pagina", "1");
                requestDescarga.AddHeader("itemsPorPagina", nroRegistros);
                requestDescarga.AddHeader("campos", "tipoDocumento,folio,fechaEmision,fechaRecepcion,formaDePago,rutEmisor,razonSocialEmisor,montoTotal,estadoSii,plazo,fechaRecepcionSii");

                requestDescarga.AddHeader("filtros", Filtros);
                requestDescarga.AddHeader("orden", "+fechaRecepcion");

                IRestResponse responseDescarga = clientDescarga.Execute(requestDescarga);

                if (responseDescarga.StatusDescription.Equals("OK"))
                {
                    RootObjectDescarga d = JsonConvert.DeserializeObject<RootObjectDescarga>(responseDescarga.Content);
                    Int32 IndexMatrix = 0;
                    String CardCode = String.Empty;
                    String DocEntryBase = String.Empty;
                    String Query = String.Empty;
                    Boolean existe = false;
                    String TotalEM = String.Empty;
                    String TotalOC = String.Empty;
                    String FoliosEMSAP = String.Empty;

                    strMensaje = d.mensaje;

                    foreach (Comunes.Documento dto in d.documentos)
                    {
                        if (!BuscarFacturasProveedores(dto.rutEmisor, dto.folio.ToString()))
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
                                dtDoc.SetValue("co_Plazo", IndexMatrix, dto.plazo);

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

                                        if (objDTE.Referencia.Count > 0)
                                        {

                                            RefOC = new List<Referencia>();
                                            RefEM = new List<Referencia>();

                                            RefOC = objDTE.Referencia.Where(i => i.TpoDocRef == "801").ToList();
                                            RefEM = objDTE.Referencia.Where(i => i.TpoDocRef == "52").ToList();

                                            if (RefEM.Count > 0)
                                            {
                                                //dtDoc.SetValue("co_Refoc", IndexMatrix, "Y");
                                                string FoliosEM = string.Join(", ", RefEM.Select(z => z.FolioRef));
                                                existe = FuncionesComunes.ExisteEntradaMercancia("52", FoliosEM, CardCode, ref TotalEM, ref FoliosEMSAP);

                                                if (existe)
                                                {
                                                    //dtDoc.SetValue("co_Base", IndexMatrix, DocEntryBase);
                                                    dtDoc.SetValue("co_Exisem", IndexMatrix, "Y");
                                                    dtDoc.SetValue("co_Folioem", IndexMatrix, FoliosEMSAP);
                                                    dtDoc.SetValue("co_Totalem", IndexMatrix, String.Format(System.Globalization.CultureInfo.GetCultureInfo("es-CL"), "{0:C0}", Convert.ToDouble(TotalEM)));
                                                }
                                                else
                                                {
                                                    //dtDoc.SetValue("co_Base", IndexMatrix, DocEntryBase);
                                                    dtDoc.SetValue("co_Exisem", IndexMatrix, "N");
                                                }

                                            }
                                            else if (RefOC.Count > 0)
                                            {
                                                string FoliosOC = string.Join(", ", RefOC.Select(z => z.FolioRef));
                                                existe = FuncionesComunes.ExisteEntradaMercancia("801", FoliosOC, CardCode, ref TotalEM, ref FoliosEMSAP);

                                                if (existe)
                                                {
                                                    //dtDoc.SetValue("co_Base", IndexMatrix, DocEntryBase);
                                                    dtDoc.SetValue("co_Exisem", IndexMatrix, "Y");
                                                    dtDoc.SetValue("co_Folioem", IndexMatrix, FoliosEMSAP);
                                                    dtDoc.SetValue("co_Totalem", IndexMatrix, String.Format(System.Globalization.CultureInfo.GetCultureInfo("es-CL"), "{0:C0}", Convert.ToDouble(TotalEM)));

                                                }
                                                else
                                                {
                                                    //dtDoc.SetValue("co_Base", IndexMatrix, DocEntryBase);
                                                    dtDoc.SetValue("co_Exisem", IndexMatrix, "N");
                                                }
                                            }
                                            if (RefOC.Count > 0)
                                            {
                                                dtDoc.SetValue("co_Refoc", IndexMatrix, "Y");

                                                string FoliosOC = string.Join(", ", RefOC.Select(z => z.FolioRef));
                                                dtDoc.SetValue("co_Foliooc", IndexMatrix, FoliosOC);
                                                existe = FuncionesComunes.ExisteReferencia("801", FoliosOC, CardCode, ref TotalOC, String.Empty);
                                                if (existe)
                                                {
                                                    //dtDoc.SetValue("co_Base", IndexMatrix, DocEntryBase);
                                                    dtDoc.SetValue("co_Exisoc", IndexMatrix, "Y");
                                                    dtDoc.SetValue("co_Totaloc", IndexMatrix, String.Format(System.Globalization.CultureInfo.GetCultureInfo("es-CL"), "{0:C0}", Convert.ToDouble(TotalOC)));
                                                }
                                                else
                                                {
                                                    dtDoc.SetValue("co_Exisoc", IndexMatrix, "N");
                                                }
                                            }

                                        }
                                        else
                                        {
                                            dtDoc.SetValue("co_Refoc", IndexMatrix, "N");
                                            dtDoc.SetValue("co_Exisoc", IndexMatrix, "N");
                                            dtDoc.SetValue("co_Exisem", IndexMatrix, "N");
                                        }

                                        DTEMatrix objDteMatrix = new DTEMatrix();
                                        objDteMatrix.FebosID = dto.febosId;
                                        objDteMatrix.objDTE = objDTE;
                                        ListaDTEMatrix.ListaDTE.Add(objDteMatrix);
                                    }
                                }

                                IndexMatrix++;
                            }
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
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(strMensaje + " " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {

                GC.Collect();
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
        /// Envia un rechazo comercial al sii del documento seleccionado en grilla.
        /// </summary>
        public static ResultMessage GestionarRechazo(SAPbouiCOM.Form oForm, String RazonRechazo, Int32 Row)
        {            
            ResultMessage result = new ResultMessage();
            
            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("oMtx").Specific;
            String FebosId = oMatrix.Columns.Item("co_FebId").Cells.Item(Row).Specific.Value;
            result = FuncionesComunes.EnviarRespuestaComercial(FebosId, "RCD", RazonRechazo, String.Empty);            
            return result;            
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
        /// </summary>
        public static void GestionarIntegracion(SAPbouiCOM.Form oForm)
        {
            Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Gestionando integración. Espere unos momentos", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("oMtx").Specific;
            String Path = System.IO.Path.GetTempPath() + "Logs.txt";
            String RutEmisor = String.Empty;
            String RznSocial = String.Empty;
            String Tipo = String.Empty;
            String Folio = String.Empty;
            String FebId = String.Empty;
            DateTime dtFechaVenc = DateTime.Now;
            String Plazo = String.Empty;
            Int32 iPlazo = 0;
            
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(Path, false))
            {                
                DTE objDTE = null;

                for (Int32 index = 1; index <= oMatrix.RowCount; index++)
                {                 
                    RutEmisor = oMatrix.Columns.Item("co_Rut").Cells.Item(index).Specific.Value;
                    RznSocial = oMatrix.Columns.Item("co_RznSoc").Cells.Item(index).Specific.Value;
                    Tipo = FuncionesComunes.ObtenerTipoDocumentoNumero(oMatrix.Columns.Item("co_Tipo").Cells.Item(index).Specific.Value);
                    Folio = oMatrix.Columns.Item("co_Folio").Cells.Item(index).Specific.Value;
                    FebId = oMatrix.Columns.Item("co_FebId").Cells.Item(index).Specific.Value;
                    Plazo = oMatrix.Columns.Item("co_Plazo").Cells.Item(index).Specific.Value;
                    Boolean esNumero = Int32.TryParse(Plazo, out iPlazo);

                    ResultMessage rslt = FuncionesComunes.ValidacionDTEIntegrado(RutEmisor, Int32.Parse(Tipo), Int64.Parse(Folio));

                    if (rslt.Success)
                    {
                        // Integración por articulo
                        if (((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Refoc").Cells.Item(index).Specific).Checked && ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Exisoc").Cells.Item(index).Specific).Checked && ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Exissn").Cells.Item(index).Specific).Checked && ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Cat").Cells.Item(index).Specific).Checked)
                        {
                            objDTE = ListaDTEMatrix.ListaDTE.Where(i => i.FebosID == FebId).Select(i => i.objDTE).SingleOrDefault();
                                                        
                            if (objDTE != null)
                            {
                                // Integrar
                                SAPbobsCOM.Documents oDoc = null;
                                switch (Tipo)
                                {
                                    case "33":
                                        oDoc = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                        break;
                                    case "34":
                                        oDoc = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                        break;
                                    case "56":
                                        oDoc = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                        break;
                                    case "61":
                                        oDoc = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
                                        break;
                                }

                                // Encabezado
                                oDoc.CardCode = oMatrix.Columns.Item("co_Card").Cells.Item(index).Specific.Value;
                                oDoc.CardName = FuncionesComunes.ObtenerCardName(RutEmisor, oDoc.CardCode);
                                switch (Tipo)
                                {
                                    case "33":
                                        oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                                        break;
                                    case "34":
                                        oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                                        break;
                                    case "52":
                                        oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes;
                                        break;
                                    case "56":
                                        oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                                        break;
                                    case "61":
                                        oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes;
                                        break;
                                    default:
                                        oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                                        break;
                                }

                                oDoc.DocDate = Convert.ToDateTime(objDTE.IdDoc.FchEmis);
                                oDoc.DocDueDate = Convert.ToDateTime(objDTE.IdDoc.FchVenc);
                                oDoc.DocTotal = objDTE.Totales.MntTotal;
                                oDoc.FolioNumber = Convert.ToInt32(Folio);
                                oDoc.FolioPrefixString = Tipo;
                                oDoc.Indicator = Tipo;
                                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                                oDoc.UserFields.Fields.Item("U_SEI_FEBOSID").Value = FebId;

                                // Detalle
                                Int32 indexDet = 0;

                                foreach (Detalle det in objDTE.Detalle)
                                {
                                    if (indexDet != 0)
                                    {
                                        oDoc.Lines.Add();
                                    }

                                    oDoc.Lines.SupplierCatNum = det.CdgItem[0].VlrCodigo;
                                    //oDoc.Lines.ItemCode = det.ItemCode;
                                    oDoc.Lines.ItemDescription = det.NmbItem;
                                    oDoc.Lines.LineTotal = det.MontoItem;
                                    oDoc.Lines.Quantity = det.QtyItem;
                                    oDoc.Lines.Price = det.PrcItem;
                                    oDoc.Lines.DiscountPercent = det.DescuentoPct;
                                    oDoc.Lines.BaseEntry = Int32.Parse(oMatrix.Columns.Item("co_Base").Cells.Item(index).Specific.Value);
                                    oDoc.Lines.BaseLine = det.LineNumBase;
                                    if (objDTE.IdDoc.TipoDTE.Equals("61"))
                                    {
                                        oDoc.Lines.BaseType = 18;
                                    }
                                    else
                                    {
                                        oDoc.Lines.BaseType = 22;
                                    }

                                    if (det.IndExe.Equals(1))
                                    {
                                        oDoc.Lines.TaxCode = "IVA_EXE";
                                    }
                                    indexDet++;
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
                                                oDoc.DiscountPercent = objDRG.ValorDR;
                                            }
                                            // Por valor
                                            else
                                            {
                                                Double SumMontoNeto = objDTE.Detalle.Where(i => i.IndExe == 0).Sum(i => i.MontoItem);
                                                Double Porcentaje = Math.Round((objDRG.ValorDR * 100 / SumMontoNeto), 2);
                                                oDoc.DiscountPercent = Porcentaje;
                                            }
                                        }
                                        // Recargo
                                        else
                                        {
                                            Int32 ExpensesCode = 0;
                                            String Query = null; 

                                            switch (Conexion_SBO.m_oCompany.DbServerType)
                                            {
                                                case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                                                    Query = "SELECT \"ExpnsCode\" FROM OEXD WHERE \"ExpnsName\" = 'RcgGlobal'";
                                                    break;
                                                default:
                                                    Query = "SELECT ExpnsCode FROM OEXD WHERE ExpnsName = 'RcgGlobal'";
                                                    break;
                                            }

                                            SAPbobsCOM.Recordset oRec = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                            oRec.DoQuery(Query);
                                            if (!oRec.EoF)
                                            {
                                                ExpensesCode = oRec.Fields.Item(0).Value;
                                            }

                                            // Por porcentaje
                                            if (objDRG.TpoValor.Equals("%"))
                                            {
                                                Double SumMontoNeto = objDTE.Detalle.Where(i => i.IndExe == 0).Sum(i => i.MontoItem);
                                                Double Monto = Math.Round((objDRG.ValorDR * SumMontoNeto / 100), 2);
                                                oDoc.Expenses.ExpenseCode = ExpensesCode;
                                                oDoc.Expenses.LineTotal = Monto;
                                                oDoc.Expenses.Add();
                                            }
                                            // Por valor
                                            else
                                            {
                                                oDoc.Expenses.ExpenseCode = ExpensesCode;
                                                oDoc.Expenses.LineTotal = objDRG.ValorDR;
                                                oDoc.Expenses.Add();
                                            }
                                        }
                                    }
                                }

                                // Referencia
                                if (objDTE.IdDoc.TipoDTE.Equals("61"))
                                {
                                    Referencia Ref = null;
                                    Ref = objDTE.Referencia.Where(i => i.TpoDocRef == "33").SingleOrDefault();
                                    if (Ref == null)
                                    {
                                        Ref = objDTE.Referencia.Where(i => i.TpoDocRef == "34").SingleOrDefault();
                                    }

                                    oDoc.UserFields.Fields.Item("U_SEI_INREF").Value = Ref.TpoDocRef;
                                    oDoc.UserFields.Fields.Item("U_SEI_FOREF").Value = Ref.FolioRef;
                                    oDoc.UserFields.Fields.Item("U_SEI_FEREF").Value = Convert.ToDateTime(Ref.FchRef);
                                    oDoc.UserFields.Fields.Item("U_SEI_CREF").Value = Ref.CodRef.ToString();
                                    oDoc.UserFields.Fields.Item("U_SEI_RAREF").Value = Ref.RazonRef;
                                }

                                Int32 RetVal = oDoc.Add();
                                String Mensaje = String.Empty;
                                if (RetVal.Equals(0))
                                {
                                    rslt = FuncionesComunes.EnviarRespuestaComercial(FebId, "ACD", String.Empty, String.Empty);
                                    if (rslt.Success)
                                    {
                                        file.WriteLine(String.Format("Exito: El DTE {0} Tipo {1} de {2}-{3} Se integro.", objDTE.IdDoc.Folio, objDTE.IdDoc.TipoDTE, objDTE.Emisor.RznSoc, objDTE.Emisor.RUTEmisor));
                                    }
                                    else
                                    {
                                        file.WriteLine(String.Format("Reparo: El DTE {0} Tipo {1} de {2}-{3} Se integro, pero no se completo proceso de intercambio.", objDTE.IdDoc.Folio, objDTE.IdDoc.TipoDTE, objDTE.Emisor.RznSoc, objDTE.Emisor.RUTEmisor));
                                    }
                                }
                                else
                                {
                                    Int32 ErrCode = 0;
                                    String ErrMsj = String.Empty;
                                    Conexion_SBO.m_oCompany.GetLastError(out ErrCode, out ErrMsj);
                                    file.WriteLine(String.Format("Error: El DTE {0} Tipo {1} de {2}-{3} No se integro. {4}", objDTE.IdDoc.Folio, objDTE.IdDoc.TipoDTE, objDTE.Emisor.RznSoc, objDTE.Emisor.RUTEmisor, ErrMsj));
                                }
                            }
                                                     
                        }
                        // Integración por servicio
                        else if (((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Exisocs").Cells.Item(index).Specific).Checked && ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Exissn").Cells.Item(index).Specific).Checked)
                        {
                            objDTE = ListaDTEMatrix.ListaDTE.Where(i => i.FebosID == FebId).Select(i => i.objDTE).SingleOrDefault();
                                                                                                                
                            if (objDTE != null)
                            {
                                // Integrar con documento base
                                Int32 BaseEntry = Int32.Parse(oMatrix.Columns.Item("co_Base").Cells.Item(index).Specific.Value);
                                SAPbobsCOM.Documents oDoc = null;
                                SAPbobsCOM.Documents oDocBase = null;
                                switch (Tipo)
                                {
                                    case "33":
                                        oDoc = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);                                            
                                        oDocBase = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
                                        break;
                                    case "34":
                                        oDoc = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                        oDocBase = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
                                        break;
                                    case "56":
                                        oDoc = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                        oDoc.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_PurchaseDebitMemo;
                                        oDocBase = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                        break;
                                    case "61":
                                        oDoc = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes);
                                        oDocBase = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                        break;
                                }

                                if (oDocBase.GetByKey(BaseEntry))
                                {
                                    oDoc.CardCode = oDocBase.CardCode;
                                    oDoc.CardName = oDocBase.CardName;
                                                                           
                                    oDoc.DocDate = Convert.ToDateTime(objDTE.IdDoc.FchEmis);
                                    oDoc.DocDueDate = Convert.ToDateTime(objDTE.IdDoc.FchVenc);
                                    oDoc.UserFields.Fields.Item("U_SEI_FCHV").Value = dtFechaVenc.AddDays(iPlazo);
                                    oDoc.DocTotal = oDocBase.DocTotal;
                                    oDoc.DiscountPercent = oDocBase.DiscountPercent;
                                    oDoc.Rounding = oDocBase.Rounding;
                                    oDoc.RoundingDiffAmount = oDocBase.RoundingDiffAmount;
                                    oDoc.FolioNumber = Convert.ToInt32(Folio);
                                    oDoc.FolioPrefixString = Tipo;
                                    oDoc.Indicator = Tipo;
                                    oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                                    oDoc.UserFields.Fields.Item("U_SEI_FEBOSID").Value = FebId;


                                    // Gastos adicionales
                                    Int32 Expenses = oDocBase.Expenses.Count;
                                    if (Expenses > 1)
                                    {
                                        for (Int32 i = 0; i < oDocBase.Expenses.Count; i++)
                                        {
                                            oDocBase.Expenses.SetCurrentLine(i);
                                            if (i != 0)
                                            {
                                                oDoc.Expenses.Add();
                                            }
                                            oDoc.Expenses.ExpenseCode = oDocBase.Expenses.ExpenseCode;
                                            oDoc.Expenses.BaseDocType = Int32.Parse(oDocBase.DocObjectCodeEx);
                                            oDoc.Expenses.BaseDocLine = oDocBase.Expenses.LineNum;
                                            oDoc.Expenses.BaseDocEntry = oDocBase.DocEntry;
                                        }
                                    }

                                    // Detalle
                                    for (Int32 i = 0; i < oDocBase.Lines.Count; i++)
                                    {
                                        oDocBase.Lines.SetCurrentLine(i);
                                        if (i != 0)
                                        {
                                            oDoc.Lines.Add();
                                        }
                                                                                    
                                        oDoc.Lines.BaseEntry = oDocBase.DocEntry;
                                        oDoc.Lines.BaseLine = oDocBase.Lines.LineNum;
                                        oDoc.Lines.BaseType = Int32.Parse(oDocBase.DocObjectCodeEx);
                                    }

                                    if (objDTE.IdDoc.TipoDTE.Equals("61"))
                                    {
                                        Referencia Ref = null;
                                        Ref = objDTE.Referencia.Where(i => i.TpoDocRef == "33").SingleOrDefault();
                                        if (Ref == null)
                                        {
                                            Ref = objDTE.Referencia.Where(i => i.TpoDocRef == "34").SingleOrDefault();
                                        }

                                        oDoc.UserFields.Fields.Item("U_SEI_INREF").Value = Ref.TpoDocRef;
                                        oDoc.UserFields.Fields.Item("U_SEI_FOREF").Value = Ref.FolioRef;
                                        oDoc.UserFields.Fields.Item("U_SEI_FEREF").Value = Convert.ToDateTime(Ref.FchRef);
                                        oDoc.UserFields.Fields.Item("U_SEI_CREF").Value = Ref.CodRef.ToString();
                                        oDoc.UserFields.Fields.Item("U_SEI_RAREF").Value = Ref.RazonRef;
                                        FuncionesComunes.CambiarStatusFacturaEstadoDePagoErroneo(Ref, oMatrix.Columns.Item("co_Card").Cells.Item(index).Specific.Value);
                                    }

                                    Int32 RetVal = oDoc.Add();
                                    String Mensaje = String.Empty;
                                    if (RetVal.Equals(0))
                                    {
                                        rslt = FuncionesComunes.EnviarRespuestaComercial(FebId, "ACD", String.Empty, String.Empty);
                                        if (rslt.Success)
                                        {
                                            file.WriteLine(String.Format("Exito: El DTE {0} Tipo {1} de {2}-{3} Se integro.", objDTE.IdDoc.Folio, objDTE.IdDoc.TipoDTE, objDTE.Emisor.RznSoc, objDTE.Emisor.RUTEmisor));
                                        }
                                        else
                                        {
                                            file.WriteLine(String.Format("Reparo: El DTE {0} Tipo {1} de {2}-{3} Se integro, pero no se completo proceso de intercambio.", objDTE.IdDoc.Folio, objDTE.IdDoc.TipoDTE, objDTE.Emisor.RznSoc, objDTE.Emisor.RUTEmisor));
                                        }
                                    }
                                    else
                                    {
                                        Int32 ErrCode = 0;
                                        String ErrMsj = String.Empty;
                                        Conexion_SBO.m_oCompany.GetLastError(out ErrCode, out ErrMsj);
                                        file.WriteLine(String.Format("Error: El DTE {0} Tipo {1} de {2}-{3} No se integro. {4}", objDTE.IdDoc.Folio, objDTE.IdDoc.TipoDTE, objDTE.Emisor.RznSoc, objDTE.Emisor.RUTEmisor, ErrMsj));
                                    }

                                }
                                else
                                {
                                    file.WriteLine(String.Format("Error: DTE {0} Tipo {1} de {2}-{3} no se encuentra documento base.", Folio, Tipo, RznSocial, RutEmisor));
                                }                                    
                            }
                                                    
                        }
                        else
                        {
                            file.WriteLine(String.Format("Error: El DTE {0} Tipo {1} de {2}-{3} No cumple con las validaciones.", Folio, Tipo, RznSocial, RutEmisor));
                        }
                    }
                    else
                    {
                        file.WriteLine(String.Format("Error: El DTE {0} Tipo {1} de {2}-{3} ya se encuentra integrado.", Folio, Tipo, RznSocial, RutEmisor));
                    }
                    
                }
            }
            System.Diagnostics.Process.Start(Path);
            Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Proceso de integración completo. Revise log de resultados", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }

        /// <summary>
        /// Integra el documento de proveedor seleccionado a SAP, ya sea por articulo o servicio.
        /// </summary>
        public static void GestionarIntegracion(SAPbouiCOM.Form oForm, Int32 index)
        {

            String mensaje = string.Empty;
            bool noRechazable = false;

            try
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Gestionando integración. Espere unos momentos", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                if (GestionarValidacionesFactura(oForm, index, out mensaje, out noRechazable))
                {
                    if (noRechazable)
                    {
                        //IntegracionPorDocumentoServicioMas(oForm);
                        IntegracionPorDocumentoServicio(oForm, index);
                    }
                    else
                    {
                        IntegracionPorDocumentoBase(oForm, index);
                    }
                }
                else
                {
                
                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Validacion: " + mensaje, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                
            }
            catch (Exception ex)
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {

            }
        }

        /// <summary>
        /// Integra el documento de proveedor seleccionado a SAP, ya sea por articulo o servicio.
        /// </summary>
        public static void GestionarIntegracionSV(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = null;

            String RutEmisor = String.Empty;
            String RznSocial = String.Empty;
            String Tipo = String.Empty;
            String Folio = String.Empty;
            String FebId = String.Empty;
            String CardCode = String.Empty;
            SAPbouiCOM.CheckBox oCheckBox = null;
            try
            {
                oMatrix = oForm.Items.Item("oMtx").Specific;
                Documents doc = new Documents();
                List<Documents> ListDoc = new List<Documents>();
                DTE objDTE = null;
                for (int index = 1; index <= oMatrix.VisualRowCount; index++)
                {
                    oCheckBox = oMatrix.Columns.Item("co_CkeckSe").Cells.Item(index).Specific;
                    if (oCheckBox.Checked)
                    {
                        
                        doc.FebId = oMatrix.Columns.Item("co_FebId").Cells.Item(index).Specific.Value;
                        FebId = doc.FebId;
                        objDTE = ListaDTEMatrix.ListaDTE.Where(x => x.FebosID == FebId).Select(x => x.objDTE).SingleOrDefault();
                        doc.RutEmisor = oMatrix.Columns.Item("co_Rut").Cells.Item(index).Specific.Value;
                        //doc.RznSocial = oMatrix.Columns.Item("co_RznSoc").Cells.Item(index).Specific.Value;
                        doc.Tipo = FuncionesComunes.ObtenerTipoDocumentoNumero(oMatrix.Columns.Item("co_Tipo").Cells.Item(index).Specific.Value);
                        doc.Folio = oMatrix.Columns.Item("co_Folio").Cells.Item(index).Specific.Value;

                        doc.CardCode = oMatrix.Columns.Item("co_Card").Cells.Item(index).Specific.Value;
                        doc.FchEmis = Convert.ToDateTime(objDTE.IdDoc.FchEmis).ToString();
                        doc.FchVenc = Convert.ToDateTime(objDTE.IdDoc.FchVenc).ToString();
                        doc.MntTotal = objDTE.Totales.MntTotal.ToString();
                        doc.IVA = objDTE.Totales.IVA.ToString();
                        ListDoc.Add(doc);
                    }
                }
                
            }
            catch (Exception ex)
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oMatrix);
            }
        }

        public static void GestionarIntegracionS(SAPbouiCOM.Form oForm)
        {
            try
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Gestionando integración. Espere unos momentos", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                IntegracionPorDocumentoServicioMas(oForm);

            }
            catch (Exception ex)
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {

            }
        }
        private static bool GestionarValidacionesFactura(SAPbouiCOM.Form oForm, Int32 index, out String mensaje, out bool noRechazable)
        {
            String Referencia = string.Empty;
            String ExisteOC = string.Empty;
            String ExisteEM = string.Empty;
            String TotalFactura = string.Empty;
            String TotalEM = string.Empty;
            String ExisteSN = string.Empty;
            String ProvNR = string.Empty;
            noRechazable = false;

            SAPbouiCOM.Matrix oMatrix = null;
            try
            {
                oMatrix = oForm.Items.Item("oMtx").Specific;
                Referencia = oMatrix.Columns.Item("co_Refoc").Cells.Item(index).Specific.Checked == true ? "Y" : "N";
                ExisteOC = oMatrix.Columns.Item("co_Exisoc").Cells.Item(index).Specific.Checked == true ? "Y" : "N";
                ExisteEM = oMatrix.Columns.Item("co_Exisem").Cells.Item(index).Specific.Checked == true ? "Y" : "N";
                TotalFactura = oMatrix.Columns.Item("co_Total").Cells.Item(index).Specific.Value;
                TotalEM = oMatrix.Columns.Item("co_Totalem").Cells.Item(index).Specific.Value;
                ExisteSN = oMatrix.Columns.Item("co_Exissn").Cells.Item(index).Specific.Checked == true ? "Y" : "N";
                ProvNR = oMatrix.Columns.Item("co_ProvNR").Cells.Item(index).Specific.Checked == true ? "Y" : "N";

                if (Referencia.Equals("N"))
                {
                    if (ProvNR.Equals("Y"))
                    {
                        noRechazable = true;
                        mensaje = "OK";
                        return true;
                    }
                    else
                    {
                        mensaje = "La factura no trae un OC en sus referencias.";
                        return false;
                    }
                }
                else if (ExisteSN.Equals("N"))
                {
                    mensaje = "Socio de negocios no creado en SAP";
                    return false;
                }
                else if (ExisteEM.Equals("N"))
                {
                    mensaje = "No existen Entradas de Mercancia en SAP para el documento";
                    return false;
                }
                else if (!TotalFactura.Equals(TotalEM))
                {
                    mensaje = "Total de Factura no es igual a el total de las entradas de Mercancia";
                    return false;
                }
                else
                {
                    mensaje = "OK";
                    return true;
                }

            }
            catch (Exception ex)
            {
                mensaje = "Se produjo un error en GestionarValidacionesFactura " + ex.Message;
                return false;

            }
            finally
            {

            }
        }

        private static String QueryLineasEntradaMercancia(String CardCode, String Folios)
        {
            String Query = string.Empty;

            try
            {
                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        Query = " SELECT T1.\"DocEntry\", T1.\"LineNum\", T1.\"ObjType\" ";
                        Query += " FROM OPDN T0 ";
                        Query += " INNER JOIN PDN1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" ";
                        Query += " WHERE T0.\"CardCode\" = '" + CardCode + "' AND T0.\"FolioNum\" IN (" + Folios + ") ";
                        break;
                    default:
                        Query = " SELECT T1.DocEntry, T1.LineNum, T1.ObjType ";
                        Query += " FROM OPDN T0 ";
                        Query += " INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry ";
                        Query += " WHERE T0.CardCode = '" + CardCode + "' AND T0.FolioNum IN (" + Folios + ") ";
                        break;
                }


                
                return Query;

            }
            catch (Exception)
            {

                return "";
            }
            finally
            {

            }
        }
        
        /// <summary>
        /// 
        /// </summary>
        private static void IntegracionPorDocumentoBase(SAPbouiCOM.Form oForm, Int32 index)
        {
            String RutEmisor = String.Empty;
            String RznSocial = String.Empty;
            String Tipo = String.Empty;
            String Folio = String.Empty;
            String FebId = String.Empty;
            DateTime dtFechaVenc = DateTime.Now;
            String Plazo = String.Empty;
            String Path = String.Empty;
            Int32 iPlazo = 0;
            String FoliosEM = string.Empty;

            SAPbouiCOM.Matrix oMatrix = null;
            DTE objDTE = null;
            SAPbobsCOM.Documents oDoc = null;
            SAPbobsCOM.Recordset oRec = null;
            try
            {
                Path = System.IO.Path.GetTempPath() + "Logs.txt";
                oMatrix = oForm.Items.Item("oMtx").Specific;

                RutEmisor = oMatrix.Columns.Item("co_Rut").Cells.Item(index).Specific.Value;
                RznSocial = oMatrix.Columns.Item("co_RznSoc").Cells.Item(index).Specific.Value;
                Tipo = FuncionesComunes.ObtenerTipoDocumentoNumero(oMatrix.Columns.Item("co_Tipo").Cells.Item(index).Specific.Value);
                Folio = oMatrix.Columns.Item("co_Folio").Cells.Item(index).Specific.Value;
                FebId = oMatrix.Columns.Item("co_FebId").Cells.Item(index).Specific.Value;
                Plazo = oMatrix.Columns.Item("co_Plazo").Cells.Item(index).Specific.Value;
                FoliosEM = oMatrix.Columns.Item("co_Folioem").Cells.Item(index).Specific.Value;
                Boolean esNumero = Int32.TryParse(Plazo, out iPlazo);

                ResultMessage rslt = FuncionesComunes.ValidacionDTEIntegrado(RutEmisor, Int32.Parse(Tipo), Int64.Parse(Folio));
                if (rslt.Success)
                {
                    objDTE = ListaDTEMatrix.ListaDTE.Where(i => i.FebosID == FebId).Select(i => i.objDTE).SingleOrDefault();
                    if (objDTE != null)
                    {
                        switch (Tipo)
                        {
                            case "33":
                                oDoc = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                break;
                            case "34":
                                oDoc = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                break;
                            case "56":
                                oDoc = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                break;
                            case "61":
                                oDoc = (SAPbobsCOM.Documents)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
                                break;
                        }

                        oDoc.CardCode = oMatrix.Columns.Item("co_Card").Cells.Item(index).Specific.Value;
                        //oDoc.CardName = FuncionesComunes.ObtenerCardName(RutEmisor, oDoc.CardCode);
                        switch (Tipo)
                        {
                            case "33":
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                                break;
                            case "34":
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                                break;
                            case "52":
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes;
                                break;
                            case "56":
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                                oDoc.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_DebitMemo;
                                break;
                            case "61":
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes;
                                break;
                            default:
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                                break;
                        }

                        oDoc.DocDate = Convert.ToDateTime(objDTE.IdDoc.FchEmis);
                        oDoc.DocDueDate = Convert.ToDateTime(objDTE.IdDoc.FchVenc);
                        oDoc.DocTotal = objDTE.Totales.MntTotal;
                        oDoc.FolioNumber = Convert.ToInt32(Folio);
                        oDoc.FolioPrefixString = Tipo;
                        oDoc.Indicator = Tipo;
                        oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                        oDoc.UserFields.Fields.Item("U_SEI_FEBOSID").Value = FebId;

                        oRec = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRec.DoQuery(QueryLineasEntradaMercancia(oDoc.CardCode, FoliosEM));

                        Int32 indexDet = 0;

                        while (!oRec.EoF)
                        {
                            if (indexDet != 0)
                            {
                                oDoc.Lines.Add();
                            }

                            switch (Tipo)
                            {
                                case "34": // Factura IVA Exento
                                    oDoc.Lines.BaseEntry = oRec.Fields.Item("DocEntry").Value;
                                    oDoc.Lines.BaseLine = oRec.Fields.Item("LineNum").Value;
                                    oDoc.Lines.BaseType = Int32.Parse(oRec.Fields.Item("ObjType").Value);
                                    oDoc.Lines.TaxCode = "IVA_EXE";
                                    break;
                                default:
                                    oDoc.Lines.BaseEntry = oRec.Fields.Item("DocEntry").Value;
                                    oDoc.Lines.BaseLine = oRec.Fields.Item("LineNum").Value;
                                    oDoc.Lines.BaseType = Int32.Parse(oRec.Fields.Item("ObjType").Value);
                                    break;
                            }                         
                            oRec.MoveNext();
                        }

                        Int32 RetVal = oDoc.Add();
                        String Mensaje = String.Empty;
                        if (RetVal.Equals(0))
                        {
                            rslt = FuncionesComunes.EnviarRespuestaComercial(FebId, "ACD", String.Empty, String.Empty);
                            if (rslt.Success)
                            {
                                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(String.Format("Exito: El DTE {0} Tipo {1} de {2}-{3} Se integro.", objDTE.IdDoc.Folio, objDTE.IdDoc.TipoDTE, objDTE.Emisor.RznSoc, objDTE.Emisor.RUTEmisor), SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                //file.WriteLine(String.Format("Exito: El DTE {0} Tipo {1} de {2}-{3} Se integro.", objDTE.IdDoc.Folio, objDTE.IdDoc.TipoDTE, objDTE.Emisor.RznSoc, objDTE.Emisor.RUTEmisor));
                            }
                            else
                            {
                                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(String.Format("Reparo: El DTE {0} Tipo {1} de {2}-{3} Se integro, pero no se completo proceso de intercambio.", objDTE.IdDoc.Folio, objDTE.IdDoc.TipoDTE, objDTE.Emisor.RznSoc, objDTE.Emisor.RUTEmisor), SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                //file.WriteLine(String.Format("Reparo: El DTE {0} Tipo {1} de {2}-{3} Se integro, pero no se completo proceso de intercambio.", objDTE.IdDoc.Folio, objDTE.IdDoc.TipoDTE, objDTE.Emisor.RznSoc, objDTE.Emisor.RUTEmisor));
                            }
                        }
                        else
                        {
                            Int32 ErrCode = 0;
                            String ErrMsj = String.Empty;
                            Conexion_SBO.m_oCompany.GetLastError(out ErrCode, out ErrMsj);
                            Conexion_SBO.m_SBO_Appl.StatusBar.SetText(String.Format("Error: El DTE {0} Tipo {1} de {2}-{3} No se integro. {4}", objDTE.IdDoc.Folio, objDTE.IdDoc.TipoDTE, objDTE.Emisor.RznSoc, objDTE.Emisor.RUTEmisor, ErrMsj), SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            //file.WriteLine(String.Format("Error: El DTE {0} Tipo {1} de {2}-{3} No se integro. {4}", objDTE.IdDoc.Folio, objDTE.IdDoc.TipoDTE, objDTE.Emisor.RznSoc, objDTE.Emisor.RUTEmisor, ErrMsj));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void IntegracionPorDocumentoServicio(SAPbouiCOM.Form oForm, Int32 index)
        {
            SAPbouiCOM.Matrix oMatrix = null;

            String RutEmisor = String.Empty;
            String RznSocial = String.Empty;
            String Tipo = String.Empty;
            String Folio = String.Empty;
            String FebId = String.Empty;
            String CardCode = String.Empty;

            try
            {
                oMatrix = oForm.Items.Item("oMtx").Specific;

                RutEmisor = oMatrix.Columns.Item("co_Rut").Cells.Item(index).Specific.Value;
                RznSocial = oMatrix.Columns.Item("co_RznSoc").Cells.Item(index).Specific.Value;
                Tipo = FuncionesComunes.ObtenerTipoDocumentoNumero(oMatrix.Columns.Item("co_Tipo").Cells.Item(index).Specific.Value);
                Folio = oMatrix.Columns.Item("co_Folio").Cells.Item(index).Specific.Value;
                FebId = oMatrix.Columns.Item("co_FebId").Cells.Item(index).Specific.Value;
                CardCode = oMatrix.Columns.Item("co_Card").Cells.Item(index).Specific.Value;

                SEI_FormDocS oFormDocS = new SEI_FormDocS(index, RutEmisor, Tipo, Folio, FebId, CardCode);
            }
            catch (Exception ex)
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oMatrix);
            }
        }

        private static void IntegracionPorDocumentoServicioMas(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = null;

            String RutEmisor = String.Empty;
            String RznSocial = String.Empty;
            String Tipo = String.Empty;
            String Folio = String.Empty;
            String FebId = String.Empty;
            String CardCode = String.Empty;
            SAPbouiCOM.CheckBox oCheckBox = null;
            try
            {
                oMatrix = oForm.Items.Item("oMtx").Specific;
                Documents doc = new Documents();
                List<Documents> ListDoc = new List<Documents>();
                DTE objDTE = null;
                for (int index=1;index<= oMatrix.VisualRowCount;index++)
                {
                    oCheckBox = oMatrix.Columns.Item("co_CkeckSe").Cells.Item(index).Specific;
                    if (oCheckBox.Checked)
                    {
                        doc.FebId = oMatrix.Columns.Item("co_FebId").Cells.Item(index).Specific.Value;
                        FebId = doc.FebId;
                        objDTE = ListaDTEMatrix.ListaDTE.Where(x => x.FebosID == FebId).Select(x => x.objDTE).SingleOrDefault();
                        doc.RutEmisor = oMatrix.Columns.Item("co_Rut").Cells.Item(index).Specific.Value;
                        //doc.RznSocial = oMatrix.Columns.Item("co_RznSoc").Cells.Item(index).Specific.Value;
                        doc.Tipo = FuncionesComunes.ObtenerTipoDocumentoNumero(oMatrix.Columns.Item("co_Tipo").Cells.Item(index).Specific.Value);
                        doc.Folio = oMatrix.Columns.Item("co_Folio").Cells.Item(index).Specific.Value;
                      
                        doc.CardCode = oMatrix.Columns.Item("co_Card").Cells.Item(index).Specific.Value;
                        doc.FchEmis = Convert.ToDateTime(objDTE.IdDoc.FchEmis).ToString();
                        doc.FchVenc = Convert.ToDateTime(objDTE.IdDoc.FchVenc).ToString();
                        doc.MntTotal = objDTE.Totales.MntTotal.ToString();
                        doc.IVA = objDTE.Totales.IVA.ToString();
                        ListDoc.Add(doc);
                    }
                }               

                //if (ListDoc.Count > 0)
                    SEI_FormDocSMas oFormDocS = new SEI_FormDocSMas(ListDoc);
               // else
                   // Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Seleccione un documento para integrar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            }
            catch (Exception ex)
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oMatrix);
            }
        }


        /// <summary>
        /// Funcion que muestra o esconde la razon de rechazo.
        /// </summary>
        public static void RazonRechazoVisible(SAPbouiCOM.Form oForm ,Boolean valor)
        {
            oForm.Items.Item("etDesde").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            oForm.Items.Item("stRR").Visible = valor;
            oForm.Items.Item("etRR").Visible = valor;
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
        /// Crea la OC del documento seleccionado en la grilla y levanta la OC creada.
        /// </summary>
        public static void LLenarOrdenCompra(SAPbouiCOM.Form oForm, Int32 Row)
        {
            SAPbouiCOM.Form oFormOC = null;

            try
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("LLenando Orden de Compra. Espere unos momentos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                SAPbouiCOM.Matrix oMatrix = null;
                Boolean checkOC = false;
                Boolean checkSN = false;
                oMatrix = oForm.Items.Item("oMtx").Specific;

                checkOC = ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Exisocs").Cells.Item(Row).Specific).Checked;
                if (!checkOC)
                {
                    checkSN = ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Exissn").Cells.Item(Row).Specific).Checked;
                    if (checkSN)
                    {
                        String FebosId = oMatrix.Columns.Item("co_FebId").Cells.Item(Row).Specific.Value;
                        String CardCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("co_Card").Cells.Item(Row).Specific).Value;

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
                                
                                Conexion_SBO.m_SBO_Appl.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_PurchaseOrder, String.Empty, String.Empty);
                                oFormOC = Conexion_SBO.m_SBO_Appl.Forms.ActiveForm;
                                oFormOC.Freeze(true);
                                oFormOC.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                                oFormOC.Items.Item("4").Specific.Value = CardCode;
                                oFormOC.Items.Item("14").Specific.Value = objDTE.IdDoc.Folio;                             
                                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oFormOC.Items.Item("3").Specific;
                                oCombo.Select("S", SAPbouiCOM.BoSearchKey.psk_ByValue);

                                SAPbouiCOM.Matrix oMtx = oFormOC.Items.Item("39").Specific;
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
                                                oFormOC.Items.Item("24").Specific.Value = objDRG.ValorDR;
                                            }
                                            // Por valor
                                            else
                                            {
                                                Double SumMontoNeto = objDTE.Detalle.Where(i => i.IndExe == 0).Sum(i => i.MontoItem);
                                                Double Porcentaje = Math.Round((objDRG.ValorDR * 100 / SumMontoNeto), 2);
                                                String sPorcentaje = Porcentaje.ToString().Replace(',', '.');
                                                oFormOC.Items.Item("24").Specific.Value = sPorcentaje;
                                            }
                                        }
                                        // Recargo
                                        else
                                        {
                                            // Por porcentaje
                                            if (objDRG.TpoValor.Equals("%"))
                                            {
                                                oFormOC.Items.Item("24").Specific.Value = (objDRG.ValorDR * -1);
                                            }
                                            // Por valor
                                            else
                                            {
                                                Double SumMontoNeto = objDTE.Detalle.Where(i => i.IndExe == 0).Sum(i => i.MontoItem);
                                                Double Porcentaje = Math.Round((objDRG.ValorDR * 100 / SumMontoNeto), 2);
                                                String sPorcentaje = Porcentaje.ToString().Replace(',', '.');
                                                oFormOC.Items.Item("24").Specific.Value = "-" + sPorcentaje;
                                            }
                                        }


                                    }
                                }

                                if (!objDTE.Totales.MontoNF.Equals(0))
                                {
                                    oFormOC.Items.Item("105").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oFormOC.Items.Item("103").Specific.Value = objDTE.Totales.MontoNF;

                                }

                                oFormOC.Items.Item("4").Specific.Value = CardCode;
                                oFormOC.Freeze(false);
                                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Orden de Compra llenada de forma correcta. Complete la información.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                            }
                        }
                        else
                        {
                            Conexion_SBO.m_SBO_Appl.StatusBar.SetText(responseGetXML.ErrorMessage, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                    }
                    else
                    {
                        Conexion_SBO.m_SBO_Appl.StatusBar.SetText("El Socio de Negocio del documento seleccionado no existe", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                }
                else
                {
                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText("La Orden de Compra del documento seleccionado ya existe", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }
            catch (Exception ex)
            {
                oFormOC.Freeze(false);
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private static bool BuscarFacturasProveedores(String CardCode, String Folios)
        {
            String Query = string.Empty;
            SAPbobsCOM.Recordset oRecordset = null;
            try
            {
                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        Query = " SELECT T0.\"DocEntry\" ";
                        Query += " FROM OPCH T0 ";
                        Query += " WHERE T0.\"LicTradNum\" = '" + CardCode + "' AND T0.\"FolioNum\" IN (" + Folios + ") ";
                        break;
                    default:
                        Query = " SELECT T0.DocEntry ";
                        Query += " FROM OPCH T0 ";
                        Query += " WHERE T0.LicTradNum = '" + CardCode + "' AND T0.FolioNum IN (" + Folios + ") ";
                        break;
                }

                oRecordset = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(Query);
                if (oRecordset.RecordCount == 0)
                    return false;
                else
                    return true;
            }
            catch (Exception)
            {

                return false;
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oRecordset);
                GC.Collect();
            }
        }

        private static bool IntegrarMaxivo(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.CheckBox oCheckBox = null;
            bool blPaso = false;
            int Row = 0;
            try
            {
                oMatrix = oForm.Items.Item("oMtx").Specific;
                for (int i = 1; i < oMatrix.VisualRowCount; i++)
                {
                    oCheckBox = oMatrix.Columns.Item("co_CkeckSe").Cells.Item(i).Specific;
                    if (oCheckBox.Checked)
                    {
                        Row = i; 
                        GestionarIntegracion(oForm, Row);
                        blPaso = true;
                    }
                }
                return blPaso;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    }


}
