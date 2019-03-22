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
    public class SEI_FormDocSMas
    {
        /// <summary>
        /// obtiene una lista de documentos sap
        /// </summary>
        /// <param name="Listdocuments">Lista de documentos</param>
        public SEI_FormDocSMas(List<Documents> Listdocuments)
        {
            CargarXML();

            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.Matrix oMatrix = null;
            int cantRow = 1;
            try
            {
                oForm = Conexion_SBO.m_SBO_Appl.Forms.Item("FormDocSMas");
                oMatrix = oForm.Items.Item("MatrixIM").Specific;
                foreach(Documents doc in Listdocuments)
                {
                    oMatrix.AddRow();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_7").Cells.Item(cantRow).Specific).Value = doc.RutEmisor;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_8").Cells.Item(cantRow).Specific).Value = doc.Folio;

                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1-1").Cells.Item(cantRow).Specific).Value = doc.DocRef;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1-2").Cells.Item(cantRow).Specific).Value = doc.NroRef;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1-3").Cells.Item(cantRow).Specific).Value = doc.FecRef;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1-4").Cells.Item(cantRow).Specific).Value = doc.RzRef;

                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_9").Cells.Item(cantRow).Specific).Value  = doc.Tipo;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_10").Cells.Item(cantRow).Specific).Value = doc.FebId;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_11").Cells.Item(cantRow).Specific).Value = doc.CardCode;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_FEmis").Cells.Item(cantRow).Specific).Value = doc.FchEmis;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_FVenc").Cells.Item(cantRow).Specific).Value = doc.FchVenc;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_Total").Cells.Item(cantRow).Specific).Value = doc.MntTotal;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_IVA").Cells.Item(cantRow).Specific).Value = doc.IVA;
                    ++cantRow;
                }
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
                    if (oFormAbierto.UniqueID.Equals("FormDocSMas"))
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
                    xml = xml + "Formularios\\" + "FormDocSMas.srf";
                    oXmlDoc.Load(xml);

                    String sXML = oXmlDoc.InnerXml.ToString();
                    SAPbouiCOM.FormCreationParams creationPackage = null;
                    creationPackage = (SAPbouiCOM.FormCreationParams)Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                    creationPackage.UniqueID = "FormDocSMas";
                    creationPackage.FormType = "FormDocSMas";
                    creationPackage.Modality = SAPbouiCOM.BoFormModality.fm_Modal;
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
                result = Msj_Appl.Errores(14, "CargarXML > SEI_FormDocSMas.cs " + ex.Message);
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// funcion controla los eventos del formulario
        /// </summary>
        /// <param name="FormUID">identificador unico del formulario</param>
        /// <param name="pVal">Item even</param>
        /// <param name="BubbleEvent">BobbleEvent</param>
        public static void m_SBO_Appl_ItemEvent(String FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = null;

            try
            {
                if (!pVal.BeforeAction)
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                            oForm = Conexion_SBO.m_SBO_Appl.Forms.Item(FormUID);
                            GestionarSeleccionCFL(oForm, pVal);
                            break;

                        case SAPbouiCOM.BoEventTypes.et_CLICK:
                            oForm = Conexion_SBO.m_SBO_Appl.Forms.Item(FormUID);
                            switch(pVal.ItemUID)
                            {
                                case "btnInt":
                                    IntegracionServicio(oForm);

                                    oForm.Close();
                                    break;
                                case "btnCompl":
                                    CopyInfo(oForm);                                        
                                    break;
                            }
                                                      
                            break;
                        default:
                            break;
                    }

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

        /// <summary>
        /// funcion gentiion el uso del ChooseFromList
        /// </summary>
        /// <param name="oForm">objeto del formulario</param>
        /// <param name="pVal">Item Even</param>
        private static void GestionarSeleccionCFL(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
            String CflID = string.Empty;
            SAPbouiCOM.ChooseFromList oCFL = null;

            SAPbouiCOM.DataTable oDataTable = null;
            try
            {
                oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                CflID = oCFLEvento.ChooseFromListUID;
                oCFL = oForm.ChooseFromLists.Item(CflID);

                oDataTable = oCFLEvento.SelectedObjects;

                if (oDataTable != null)
                {
                    SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("MatrixIM").Specific;
                    try
                    {
                        switch (CflID)
                        {
                            case "CFL_ACT":

                                oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("AcctCode", 0).ToString();
                                break;
                            case "CFL_dim1":
                                oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("OcrCode", 0).ToString();
                                break;
                            case "CFL_dim2":
                                oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("OcrCode", 0).ToString();
                                break;
                            case "CFL_dim3":
                                oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("OcrCode", 0).ToString();
                                break;
                            case "CFL_dim4":
                                oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("OcrCode", 0).ToString();
                                break;
                            case "CFL_dim5":
                                oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("OcrCode", 0).ToString();
                                break;

                            default:
                                break;
                        }
                    }
                    catch (Exception ex)
                    { }
                }
            }
            catch (Exception ex)
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oCFLEvento);
                FuncionesComunes.LiberarObjetoGenerico(oCFL);
                FuncionesComunes.LiberarObjetoGenerico(oDataTable);
            }

        }

        /// <summary>
        /// funcion integra facturas
        /// </summary>
        /// <param name="oForm"></param>
        private static void IntegracionServicio(SAPbouiCOM.Form oForm)
        {
            //DTE objDTE = null;
            SAPbobsCOM.Documents oDoc = null;
            SAPbouiCOM.Matrix oMatrix = null;
            string Descripcion = string.Empty;
            string Cuenta = string.Empty;
            string Dim1 = string.Empty;
            string Dim2 = string.Empty;
            string Dim3 = string.Empty;
            string Dim4 = string.Empty;
            string Dim5 = string.Empty;

            string RutEmisor = String.Empty;
            string Tipo = String.Empty;
            string Folio = String.Empty;
            string FebId = String.Empty;
            string CardCode = String.Empty;
            string FchEmis = String.Empty;
            string FchVenc = String.Empty;
            string MntTotal = String.Empty;
            string IVA = String.Empty;
            string NroRef = null;

            try
            {
                oMatrix = oForm.Items.Item("MatrixIM").Specific;
                for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                {

                    Descripcion = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value; // oForm.Items.Item("Item_1").Specific.Value.ToString();
                    Cuenta = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific).Value;
                    Dim1 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(i).Specific).Value;
                    Dim2 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_3").Cells.Item(i).Specific).Value;
                    Dim3 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_4").Cells.Item(i).Specific).Value;
                    Dim4 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_5").Cells.Item(i).Specific).Value;
                    Dim5 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_6").Cells.Item(i).Specific).Value;

                    RutEmisor = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_7").Cells.Item(i).Specific).Value;
                    Tipo = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_9").Cells.Item(i).Specific).Value;
                    Folio = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_8").Cells.Item(i).Specific).Value;
                    FebId = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_10").Cells.Item(i).Specific).Value;
                    CardCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_11").Cells.Item(i).Specific).Value;
                    FchEmis = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_FEmis").Cells.Item(i).Specific).Value;
                    FchVenc = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_FVenc").Cells.Item(i).Specific).Value;
                    MntTotal = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_Total").Cells.Item(i).Specific).Value;
                    IVA = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_IVA").Cells.Item(i).Specific).Value;

                    NroRef = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1-1").Cells.Item(i).Specific).Value;

                    ResultMessage rslt = FuncionesComunes.ValidacionDTEIntegrado(RutEmisor, Int32.Parse(Tipo), Int64.Parse(Folio));
                    rslt.Success = true;
                    if (rslt.Success)
                    {
                        //objDTE = ListaDTEMatrix.ListaDTE.Where(x => x.FebosID == FebId).Select(x => x.objDTE).SingleOrDefault();
                        //if (objDTE != null)
                        //{
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

                            oDoc.CardCode = CardCode;
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
                                    break;
                                case "61":
                                    oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes;
                                    break;
                                default:
                                    oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                                    break;
                            }

                            oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;

                            oDoc.DocDate = Convert.ToDateTime(FchEmis);
                            oDoc.DocDueDate = Convert.ToDateTime(FchVenc);
                            oDoc.DocTotal = Convert.ToDouble(MntTotal);
                            oDoc.FolioNumber = Convert.ToInt32(Folio);
                            oDoc.FolioPrefixString = Tipo;
                            oDoc.Indicator = Tipo;
                            oDoc.UserFields.Fields.Item("U_SEI_FEBOSID").Value = FebId;
                            oDoc.NumAtCard = NroRef;

                            oDoc.Lines.ItemDescription = Descripcion;
                            oDoc.Lines.AccountCode = Cuenta;
                            oDoc.Lines.LineTotal = Convert.ToDouble(Convert.ToDouble(MntTotal) - Convert.ToDouble(IVA));
                            switch (Tipo)
                            {
                                case "34":
                                    oDoc.Lines.TaxCode = "IVA_EXE";
                                    break;
                                default:

                                    break;
                            }

                            if (!string.IsNullOrEmpty(Dim1)) oDoc.Lines.CostingCode = Dim1;
                            if (!string.IsNullOrEmpty(Dim2)) oDoc.Lines.CostingCode2 = Dim2;
                            if (!string.IsNullOrEmpty(Dim3)) oDoc.Lines.CostingCode3 = Dim3;
                            if (!string.IsNullOrEmpty(Dim4)) oDoc.Lines.CostingCode4 = Dim4;
                            if (!string.IsNullOrEmpty(Dim5)) oDoc.Lines.CostingCode5 = Dim5;


                            Int32 RetVal = oDoc.Add();
                            String Mensaje = String.Empty;
                            if (RetVal.Equals(0))
                            {
                                rslt = FuncionesComunes.EnviarRespuestaComercial(FebId, "ACD", String.Empty, String.Empty);
                                if (rslt.Success)
                                {
                                        Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Exito: El DTE :" + Folio + " Tipo :" + Tipo + " de :" + RutEmisor, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                }
                                else
                                {
                                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Reparo: El DTE :" + Folio + " Tipo :" + Tipo + " de :" + RutEmisor + " Se integro, pero no se completo proceso de intercambio.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                    //file.WriteLine(String.Format("Reparo: El DTE {0} Tipo {1} de {2}-{3} Se integro, pero no se completo proceso de intercambio.", objDTE.IdDoc.Folio, objDTE.IdDoc.TipoDTE, objDTE.Emisor.RznSoc, objDTE.Emisor.RUTEmisor));
                                }
                            }
                            else
                            {
                                Int32 ErrCode = 0;
                                String ErrMsj = String.Empty;
                                Conexion_SBO.m_oCompany.GetLastError(out ErrCode, out ErrMsj);
                                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Error: El DTE :" + Folio + " Tipo :" + Tipo + " de :" + RutEmisor + " No se integro " + ErrMsj, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                //file.WriteLine(String.Format("Error: El DTE {0} Tipo {1} de {2}-{3} No se integro. {4}", objDTE.IdDoc.Folio, objDTE.IdDoc.TipoDTE, objDTE.Emisor.RznSoc, objDTE.Emisor.RUTEmisor, ErrMsj));
                            }
                        //}
                    }
                }
            }
            catch (Exception ex)
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
               // FuncionesComunes.LiberarObjetoGenerico(objDTE);
                FuncionesComunes.LiberarObjetoGenerico(oDoc);
            }

        }

        private static void CopyInfo(SAPbouiCOM.Form oForm)
        {
            string sDescripcion = null;
            string sCuenta = null;
            string sDim1 = null;
            string sDim2 = null;
            string sDim3 = null;
            string sDim4 = null;
            string sDim5 = null;                

            SAPbouiCOM.Matrix oMatrix = null;
            try
            {
                oMatrix = oForm.Items.Item("MatrixIM").Specific;
                oForm.Freeze(true);
                for(int i= 1; i <= oMatrix.RowCount;i++)
                {
                    if (i == 1)
                    {
                        sDescripcion = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value;
                        sCuenta = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific).Value;
                        sDim1 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(i).Specific).Value;
                        sDim2 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_3").Cells.Item(i).Specific).Value;
                        sDim3 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_4").Cells.Item(i).Specific).Value;
                        sDim4 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_5").Cells.Item(i).Specific).Value;
                        sDim5 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_6").Cells.Item(i).Specific).Value;
                    }
                    else
                    {
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value = sDescripcion;
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific).Value = sCuenta;
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(i).Specific).Value = sDim1;
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_3").Cells.Item(i).Specific).Value = sDim2;
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_4").Cells.Item(i).Specific).Value = sDim3;
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_5").Cells.Item(i).Specific).Value = sDim4;
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_6").Cells.Item(i).Specific).Value = sDim5;
                    }
                }
                

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
            }
        }

    }
}
