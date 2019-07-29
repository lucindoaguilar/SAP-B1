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
using Newtonsoft.Json;
using RestSharp;

namespace Addon_Facturas_Proveedores.ClaseFormulario
{
    public class SEI_FormErr
    {
        public SEI_FormErr(String FormUID)
        {
            CargarXML();
            CargarFormulario(FormUID);
        }

        /// <summary>
        /// Carga el XML del formulario
        /// </summary>
        private void CargarXML()
        {
            XmlDocument oXmlDoc = null;
            ResultMessage result = new ResultMessage();

            try
            {
                oXmlDoc = new XmlDocument();
                
                String xml = "";
                xml = Application.StartupPath.ToString();
                xml += "\\";
                xml = xml + "Formularios\\" + "FormErr.srf";
                oXmlDoc.Load(xml);
                
                String sXML = oXmlDoc.InnerXml.ToString();
                Conexion_SBO.m_SBO_Appl.LoadBatchActions(ref sXML);                
            }
            catch (Exception ex)
            {
                result = Msj_Appl.Errores(14, "CargarXML > SEI_FormErr.cs " + ex.Message);
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
        
            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("Item_0").Specific;
            oMatrix.Layout = SAPbouiCOM.BoMatrixLayoutType.mlt_Normal;
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                        
            // Datasource para blindear matrix
            oForm.DataSources.DataTables.Add("DOCUMENTOS");
            SAPbouiCOM.DataTable dt = oForm.DataSources.DataTables.Item("DOCUMENTOS");

            dt.Columns.Add("col_DocEnt", SAPbouiCOM.BoFieldsType.ft_Integer);
            dt.Columns.Add("col_DocNum", SAPbouiCOM.BoFieldsType.ft_Integer);
            dt.Columns.Add("col_Tipo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("col_Prov", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("col_Folio", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("col_Fecha", SAPbouiCOM.BoFieldsType.ft_Date);
            dt.Columns.Add("col_Monto", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("col_FebId", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

            oForm.DataSources.UserDataSources.Add("col_chk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oMatrix.Columns.Item("col_chk").DataBind.SetBound(true, "", "col_chk");
            oMatrix.Columns.Item("col_DocEnt").DataBind.Bind("DOCUMENTOS", "col_DocEnt");
            oMatrix.Columns.Item("col_DocNum").DataBind.Bind("DOCUMENTOS", "col_DocNum");
            oMatrix.Columns.Item("col_Tipo").DataBind.Bind("DOCUMENTOS", "col_Tipo");
            oMatrix.Columns.Item("col_Prov").DataBind.Bind("DOCUMENTOS", "col_Prov");
            oMatrix.Columns.Item("col_Folio").DataBind.Bind("DOCUMENTOS", "col_Folio");
            oMatrix.Columns.Item("col_Fecha").DataBind.Bind("DOCUMENTOS", "col_Fecha");
            oMatrix.Columns.Item("col_Monto").DataBind.Bind("DOCUMENTOS", "col_Monto");
            oMatrix.Columns.Item("col_FebId").DataBind.Bind("DOCUMENTOS", "col_FebId");

            BindMatrix(oMatrix, dt);
            
            
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
            ResultMessage result = new ResultMessage();
            SAPbouiCOM.Matrix oMatrix = null;
            Int32 Row = -1;

            try
            {
                if (pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CLICK) && pVal.ItemUID.Equals("bt_Salir"))
                {
                    oForm.Close();
                }
                if (pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CLICK) && pVal.ItemUID.Equals("bt_Ver"))
                {
                    oMatrix = oForm.Items.Item("Item_0").Specific;
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
                
            }            
            catch(Exception ex)
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Enlaza matrix con datos de la BD
        /// </summary>
        public static void BindMatrix(SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DataTable dt)
        {
            String Query = null;
            switch (Conexion_SBO.m_oCompany.DbServerType)
            {                
                case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                    Query = "SELECT \"DocEntry\",  \"DocNum\",  CASE \"FolioPref\" WHEN '33' THEN 'Factura Electrónica' WHEN '34' THEN 'Factura Exenta Electrónica' END AS \"FolioPref\", \"CardName\", \"FolioNum\", \"DocDate\", \"DocTotal\", \"U_SEI_FEBOSID\" FROM OPCH WHERE \"U_SEI_CONTADO\" = 'N'";
                    break;
                default:
                    Query = "SELECT DocEntry, DocNum,  CASE FolioPref WHEN '33' THEN 'Factura Electrónica' WHEN '34' THEN 'Factura Exenta Electrónica' END AS FolioPref, CardName, FolioNum, DocDate, DocTotal, U_SEI_FEBOSID FROM OPCH WHERE U_SEI_CONTADO = 'N'";
                    break;
            }
            dt.ExecuteQuery(Query);
            oMatrix.LoadFromDataSource();
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
                String FebosId = oMatrix.Columns.Item("col_FebId").Cells.Item(Row).Specific.Value;

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

    }


}
