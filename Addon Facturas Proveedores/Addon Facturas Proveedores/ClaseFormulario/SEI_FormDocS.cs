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
    class SEI_FormDocS
    {

        private static String Descripcion = string.Empty;
        private static String Cuenta = string.Empty;
        private static String Dim1 = string.Empty;
        private static String Dim2 = string.Empty;
        private static String Dim3 = string.Empty;
        private static String Dim4 = string.Empty;
        private static String Dim5 = string.Empty;

        private static String RutEmisor = String.Empty;
        private static String Tipo = String.Empty;
        private static String Folio = String.Empty;
        private static String FebId = String.Empty;
        private static String CardCode = String.Empty;
        private static Int32 index = 0;



        public SEI_FormDocS(Int32 indexAnt, String RutEmisorAnt, String TipoAnt, String FolioAnt, String FebIdAnt, String CardCodeAnt)
        {
            index = indexAnt;
            RutEmisor = RutEmisorAnt;
            Tipo = TipoAnt;
            Folio = FolioAnt;
            FebId = FebIdAnt;
            CardCode = CardCodeAnt;


            Descripcion = string.Empty;
            Cuenta = string.Empty;
            Dim1 = string.Empty;
            Dim2 = string.Empty;
            Dim3 = string.Empty;
            Dim4 = string.Empty;
            Dim5 = string.Empty;

            CargarXML();
            CargarFormulario("FormDocS");
        }

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
                    if (oFormAbierto.UniqueID.Equals("FormDocS"))
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
                    xml = xml + "Formularios\\" + "FormDocS.srf";
                    oXmlDoc.Load(xml);

                    String sXML = oXmlDoc.InnerXml.ToString();
                    SAPbouiCOM.FormCreationParams creationPackage = null;
                    creationPackage = (SAPbouiCOM.FormCreationParams)Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                    creationPackage.UniqueID = "FormDocS";
                    creationPackage.FormType = "FormDocS";
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

        private static void CargarFormulario(String FormUID)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.EditText oEditText = null;

            try
            {
                oForm = Conexion_SBO.m_SBO_Appl.Forms.Item(FormUID);

                oEditText = oForm.Items.Item("Item_3").Specific;
                FuncionesComunes.CrearCFL_CuentaContable(oForm, oEditText);

                oEditText = oForm.Items.Item("Item_6").Specific;
                FuncionesComunes.CrearCFL_Dimensiones(oForm, oEditText, "1");

                oEditText = oForm.Items.Item("Item_8").Specific;
                FuncionesComunes.CrearCFL_Dimensiones(oForm, oEditText, "2");

                oEditText = oForm.Items.Item("Item_10").Specific;
                FuncionesComunes.CrearCFL_Dimensiones(oForm, oEditText, "3");

                oEditText = oForm.Items.Item("Item_12").Specific;
                FuncionesComunes.CrearCFL_Dimensiones(oForm, oEditText, "4");

                oEditText = oForm.Items.Item("Item_14").Specific;
                FuncionesComunes.CrearCFL_Dimensiones(oForm, oEditText, "5");
            }
            catch (Exception ex)
            {
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oForm);
                FuncionesComunes.LiberarObjetoGenerico(oEditText);
            }


        }

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

                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            oForm = Conexion_SBO.m_SBO_Appl.Forms.Item(FormUID);
                            if (pVal.ItemUID.Equals("Item_15"))
                            {
                                Descripcion = oForm.Items.Item("Item_1").Specific.Value.ToString();
                                Cuenta = oForm.Items.Item("Item_3").Specific.Value.ToString();
                                Dim1 = oForm.Items.Item("Item_6").Specific.Value.ToString();
                                Dim2 = oForm.Items.Item("Item_8").Specific.Value.ToString();
                                Dim3 = oForm.Items.Item("Item_10").Specific.Value.ToString();
                                Dim4 = oForm.Items.Item("Item_12").Specific.Value.ToString();
                                Dim5 = oForm.Items.Item("Item_14").Specific.Value.ToString();

                                IntegracionServicio();

                                oForm.Close();
                            }
                            if (pVal.ItemUID.Equals("Item_16"))
                            {
                                oForm.Close();
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

                    switch (CflID)
                    {
                        case "CflCC":
                            oForm.DataSources.UserDataSources.Item("ds_Acc").Value = oDataTable.GetValue("AcctCode", 0).ToString();
                            break;
                        case "CflDim1":
                            oForm.DataSources.UserDataSources.Item("ds_Dim1").Value = oDataTable.GetValue("OcrCode", 0).ToString();
                            break;
                        case "CflDim2":
                            oForm.DataSources.UserDataSources.Item("ds_Dim2").Value = oDataTable.GetValue("OcrCode", 0).ToString();
                            break;
                        case "CflDim3":
                            oForm.DataSources.UserDataSources.Item("ds_Dim3").Value = oDataTable.GetValue("OcrCode", 0).ToString();
                            break;
                        case "CflDim4":
                            oForm.DataSources.UserDataSources.Item("ds_Dim4").Value = oDataTable.GetValue("OcrCode", 0).ToString();
                            break;
                        case "CflDim5":
                            oForm.DataSources.UserDataSources.Item("ds_Dim5").Value = oDataTable.GetValue("OcrCode", 0).ToString();
                            break;

                        default:
                            break;
                    }
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

        private static void IntegracionServicio()
        {
            DTE objDTE = null;
            SAPbobsCOM.Documents oDoc = null;
            try
            {
                ResultMessage rslt = FuncionesComunes.ValidacionDTEIntegrado(RutEmisor, Int32.Parse(Tipo), Folio);// Int64.Parse(Folio));
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

                        oDoc.DocDate = Convert.ToDateTime(objDTE.IdDoc.FchEmis);
                        oDoc.DocDueDate = Convert.ToDateTime(objDTE.IdDoc.FchVenc);
                        oDoc.DocTotal = objDTE.Totales.MntTotal;
                        oDoc.FolioNumber = Convert.ToInt32(Folio);
                        oDoc.FolioPrefixString = Tipo;
                        oDoc.Indicator = Tipo;
                        oDoc.UserFields.Fields.Item("U_SEI_FEBOSID").Value = FebId;

                        oDoc.Lines.ItemDescription = Descripcion;
                        oDoc.Lines.AccountCode = Cuenta;
                        oDoc.Lines.LineTotal = (objDTE.Totales.MntTotal - objDTE.Totales.IVA);
                        switch(Tipo)
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
                            rslt.Success = true;
                            //rslt = FuncionesComunes.EnviarRespuestaComercial(FebId, "ACD", String.Empty, String.Empty);
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
                FuncionesComunes.LiberarObjetoGenerico(objDTE);
                FuncionesComunes.LiberarObjetoGenerico(oDoc);
            }
            
        }


    }
}
