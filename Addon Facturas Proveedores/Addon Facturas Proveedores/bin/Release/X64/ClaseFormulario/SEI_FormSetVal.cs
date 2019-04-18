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
                oForm = Conexion_SBO.m_SBO_Appl.Forms.Item(FormUID);
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                       switch(pVal.ItemUID)
                        {
                            case "1":
                                
                                break;

                        }
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
                oRecordset.DoQuery("SELECT \"DocNum\" FROM \"@SEI_SETVALH\" ");
                if (oRecordset.RecordCount > 0 )
                {
                    SAPbouiCOM.EditText oedit = null;
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    oedit = oForm.Items.Item("DocNum").Specific;
                    oedit.Value = "1";
                    oForm.EnableMenu("1281", false);// 1281 --> Buscar
                    oForm.EnableMenu("1282", false);// 1282 --> Crear 
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("DocNum").Visible = false;

                }
                else
                {

                    oMatrix.AddRow(2);
                    //lines
                    ((SAPbouiCOM.EditText)(oMatrix.Columns.Item("Col_0").Cells.Item(1).Specific)).Value = "Entrada de mercancias";
                    ((SAPbouiCOM.CheckBox)(oMatrix.Columns.Item("Col_1").Cells.Item(1).Specific)).Checked = true;
                    ((SAPbouiCOM.EditText)(oMatrix.Columns.Item("Col_2").Cells.Item(1).Specific)).Value = "1";
                    ((SAPbouiCOM.EditText)(oMatrix.Columns.Item("Col_0").Cells.Item(2).Specific)).Value = "Oferta de compra";
                    ((SAPbouiCOM.CheckBox)(oMatrix.Columns.Item("Col_1").Cells.Item(2).Specific)).Checked = true;
                    ((SAPbouiCOM.EditText)(oMatrix.Columns.Item("Col_2").Cells.Item(1).Specific)).Value = "2";

                    ((SAPbouiCOM.CheckBox)(oForm.Items.Item("cbSentAcep").Specific)).Checked = true;
                }
                oForm.Items.Item("Mtx1").Visible = false;
                oForm.Freeze(false);

            }
            catch(Exception ex)
            {
                oForm.Freeze(false);
            }
        }

        private static void UpdateInfo(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbobsCOM.GeneralService oDocGeneralService = null;
            SAPbobsCOM.CompanyService oCompService = null;
            SAPbobsCOM.GeneralData oDocGeneralData = null;
            SAPbobsCOM.GeneralDataCollection oDocLinesCollection = null;
            SAPbobsCOM.GeneralData oDocLineGeneralData = null;
            string sDocEntry = null,sCode=null, sDscrp = null;
            bool Check = false;

            try
            {
                #region fm_UPDATE_MODE Validaciones
                //MODIFICAR Setting Validaciones

                sCode = oForm.DataSources.DBDataSources.Item("@SEI_SETVALH").GetValue("Code", 0);
                sDocEntry = oForm.DataSources.DBDataSources.Item("@SEI_SETVALH").GetValue("DocEntry", 0);


                SAPbobsCOM.GeneralDataParams oGeneralDataParams = null;
                Conexion_SBO.m_oCompany.StartTransaction();
                oCompService = Conexion_SBO.m_oCompany.GetCompanyService();
                oDocGeneralService = oCompService.GetGeneralService("SEI_SETVAL");
                oGeneralDataParams = oDocGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralDataParams.SetProperty("Code", sCode);
                oDocGeneralData = oDocGeneralService.GetByParams(oGeneralDataParams);


                //oDocGeneralData.SetProperty("U_FechaParada", dtFechP.ToString("yyyy-MM-dd"));
                //oDocGeneralData.SetProperty("U_Turno", strTurP);

                #region Lines validaciones
                oDocLinesCollection = oDocGeneralData.Child("SEI_SETVALL");
                oMatrix = oForm.Items.Item("Mtx1").Specific;
                for (int reg = 1; reg <= oMatrix.RowCount; reg++)
                {
                    #region Asignacion de Valores
                    sDscrp = ((SAPbouiCOM.EditText)(oMatrix.Columns.Item("Col_0").Cells.Item(reg).Specific)).Value;
                    Check = ((SAPbouiCOM.CheckBox)(oMatrix.Columns.Item("Col_1").Cells.Item(reg).Specific)).Checked;
                    #endregion Asignacion de Valores

                    #region Crear Lista de validaciones
                    oDocLineGeneralData = oDocLinesCollection.Item(reg-1);
                    oDocLineGeneralData.SetProperty("U_Dscr", sDscrp);
                    oDocLineGeneralData.SetProperty("U_Validar", (Check ==  true ?"Y" : "N"));
                    #endregion Crear Lista de validaciones                                    
                }
                //oDocGeneralService.Add(oDocGeneralData);
                #endregion Lines validaciones

                oDocGeneralService.Update(oDocGeneralData);

                if (Conexion_SBO.m_oCompany.InTransaction)
                {
                    Conexion_SBO.m_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }

                oCompService = null;
                GC.Collect();
                
                #endregion fm_UPDATE_MODE Validaciones
            }
            catch (Exception ex)
            {
                if (Conexion_SBO.m_oCompany.InTransaction)
                {
                    Conexion_SBO.m_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }
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
