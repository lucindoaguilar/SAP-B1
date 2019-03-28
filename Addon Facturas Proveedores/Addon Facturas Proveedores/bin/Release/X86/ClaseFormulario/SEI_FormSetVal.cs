using Addon_Facturas_Proveedores.Comunes;
using Addon_Facturas_Proveedores.Conexiones;
using Addon_Facturas_Proveedores.Documento;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace Addon_Facturas_Proveedores.ClaseFormulario
{
    public class SEI_FormSetVal
    {
        
        /// <summary>
        /// Constructor
        /// </summary>
        public SEI_FormSetVal()
        {
            try
            {
                CargarXML();

            }
            catch (Exception ex)
            { }

        }

        /// <summary>
        /// funcion carga el XML del formulario
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
                    if (oFormAbierto.UniqueID.Equals("SET_VAL"))
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
                    xml = xml + "Formularios\\" + "FormSetVal.srf";
                    oXmlDoc.Load(xml);

                    String sXML = oXmlDoc.InnerXml.ToString();
                    SAPbouiCOM.FormCreationParams creationPackage = null;
                    creationPackage = (SAPbouiCOM.FormCreationParams)Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                    creationPackage.UniqueID = "SET_VAL";
                    creationPackage.FormType = "SET_VAL";
                    creationPackage.Modality = SAPbouiCOM.BoFormModality.fm_Modal;
                    creationPackage.XmlData = sXML;
                    oForm = Conexion_SBO.m_SBO_Appl.Forms.AddEx(creationPackage);
                    oForm.Visible = true;
                    LoadInfo(oForm);
                }
                // Si no traer al frente
                else
                {
                    Conexion_SBO.m_SBO_Appl.Forms.Item(index).Select();
                }
            }
            catch (Exception ex)
            {
                result = Msj_Appl.Errores(14, "CargarXML > SEI_FormSetVal.cs " + ex.Message);
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static void ItemEventEventHandler(String FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = null;

            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        break;
                }

            }
            catch (Exception ex)
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oForm);
            }
        }


        private static void LoadInfo(SAPbouiCOM.Form oForm)
        {
            SAPbobsCOM.Recordset oRecordset = null;
            SAPbouiCOM.Matrix oMatrix = null;
            try
            {
                oForm.Freeze(true);
                oRecordset = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oMatrix = oForm.Items.Item("Mtx1").Specific;
                oRecordset.DoQuery("SELECT \"DocEntry\", \"Code\" FROM \"@SEI_SETVALH\" WHERE \"Code\" = 1001 ");
                if (oRecordset.RecordCount > 0 )
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    oForm.DataSources.DBDataSources.Item("@SEI_SETVALH").SetValue("Code", 0, oRecordset.Fields.Item("Code").Value);
                    //lines
                    string sQuery = "SELECT \"U_Dscr\",\"U_Validar\" FROM \"@SEI_SETVALL\" WHERE \"Code\" = '1001' ";
                    oRecordset.DoQuery(sQuery);
                    string val = null;
                    for (int i =1; i <= oRecordset.RecordCount;i++)
                    {
                        val = (oRecordset.Fields.Item("U_Validar").Value);

                        oMatrix.AddRow();
                        ((SAPbouiCOM.EditText)(oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific)).Value = oRecordset.Fields.Item("U_Dscr").Value;
                        ((SAPbouiCOM.CheckBox)(oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific)).Checked = false;

                        //oForm.DataSources.DBDataSources.Item("@SEI_SETVALL").SetValue("U_Dscr", i, oRecordset.Fields.Item("U_Dscr").Value);
                        //oForm.DataSources.DBDataSources.Item("@SEI_SETVALL").SetValue("U_Validar", i, oRecordset.Fields.Item("U_Validar").Value);
                        oRecordset.MoveNext();
                    }
                }
                else
                {

                    oMatrix.AddRow(2);
                    //head
                    oForm.DataSources.DBDataSources.Item("@SEI_SETVALH").SetValue("Code", 0, "1001");
                    //lines
                    ((SAPbouiCOM.EditText)(oMatrix.Columns.Item("Col_0").Cells.Item(1).Specific)).Value = "Entrada de mercancias";
                    ((SAPbouiCOM.CheckBox)(oMatrix.Columns.Item("Col_1").Cells.Item(1).Specific)).Checked = false;
                    ((SAPbouiCOM.EditText)(oMatrix.Columns.Item("Col_0").Cells.Item(2).Specific)).Value = "Oferta de compra";
                    ((SAPbouiCOM.CheckBox)(oMatrix.Columns.Item("Col_1").Cells.Item(2).Specific)).Checked = false;
                    //oForm.DataSources.DBDataSources.Item("@SEI_SETVALL").SetValue("U_Dscr", 1, "Entrada de mercancias");
                    //oForm.DataSources.DBDataSources.Item("@SEI_SETVALL").SetValue("U_Validar", 1, "N");
                    //oForm.DataSources.DBDataSources.Item("@SEI_SETVALL").SetValue("U_Dscr", 2, "Oferta de compra");
                    //oForm.DataSources.DBDataSources.Item("@SEI_SETVALL").SetValue("U_Validar", 2, "N");
                    //oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                }
                oForm.Freeze(false);

            }
            catch(Exception ex)
            {
                oForm.Freeze(false);
            }
        }

        public void Consul()
        {



            //HEAD
            //SELECT T0."Code", T0."Name", T0."DocEntry", T0."Canceled", T0."Object", T0."LogInst", T0."UserSign", 
            //T0."Transfered", T0."CreateDate", T0."CreateTime", T0."UpdateDate", T0."UpdateTime", T0."DataSource" 
            //FROM "ESIGN_TEST"."@SEI_SETVALH"  T0
            //lINEA
            //SELECT T0."Code", T0."LineId", T0."Object", T0."LogInst", T0."U_Valdar" FROM "ESIGN_TEST"."@SEI_SETVALL"  T0
        }
    }
}
