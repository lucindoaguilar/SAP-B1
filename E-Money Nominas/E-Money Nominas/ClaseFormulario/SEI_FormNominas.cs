using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml;
using E_Money_Nominas.Comunes;
using E_Money_Nominas.Conexiones;

namespace E_Money_Nominas.ClaseFormulario
{
    public class SEI_FormNominas
    {
        public static SAPbouiCOM.Form oForm = null;
        public static ResultMessage result = new ResultMessage();

        public SEI_FormNominas(string FormUID)
        {
            CargarXML();
            DefinirFormulario(FormUID);
        }

        /// <summary>
        /// Carga los valores del form desde el XML (SRF)
        /// </summary>
        private void CargarXML()
        {
            XmlDocument oXmlDoc = null;
            try
            {
                oXmlDoc = new XmlDocument();

                String xml = "";
                xml = Application.StartupPath.ToString();
                xml += "\\";
                xml = xml + "Formularios\\" + "FormNominas.srf";
                oXmlDoc.Load(xml);

                string sXML = oXmlDoc.InnerXml.ToString();
                Conexion_SBO.m_SBO_Appl.LoadBatchActions(ref sXML);
                Conexion_SBO.m_SBO_Appl.Forms.Item("SEI_NOM").Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
            }

            catch (Exception ex)
            {

                result = Msj_Appl.Errores(14, "CargarXML > SEI_FormNominas.cs " + ex.Message);
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        /// <summary>
        /// Enlaza el combobox de pagos con los datos de pagos de la BD.
        /// </summary>
        /// <param name="FormUID"></param>
        private void DefinirFormulario(string FormUID)
        {
            // Form actual.
            oForm = Conexion_SBO.m_SBO_Appl.Forms.Item(FormUID);

            // Asociar Fecha
            SAPbouiCOM.Item oItem = oForm.Items.Item("6");
            SAPbouiCOM.EditText oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

            oForm.DataSources.UserDataSources.Add("udsFecha", SAPbouiCOM.BoDataType.dt_DATE, 100);
            oEditText.DataBind.SetBound(true, "", "udsFecha");

            // Crear Datasource para pagos
            oItem = oForm.Items.Item("4");
            SAPbouiCOM.ComboBox oComboBox = ((SAPbouiCOM.ComboBox)(oItem.Specific));
            oForm.DataSources.UserDataSources.Add("udsPagos", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100);
            oComboBox.DataBind.SetBound(true, "", "udsPagos");
        }

        /// <summary>
        /// Maneja los eventos de los item del formulario
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public static void m_SBO_Appl_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = null;
            string Pago = string.Empty;
            string Fecha = string.Empty;

            try
            {
                oForm = Conexion_SBO.m_SBO_Appl.Forms.Item(FormUID);
                if (pVal.BeforeAction == false)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
                    {                        
                        if (pVal.ItemUID.Equals("6"))
                        {
                            // Obtener Fecha
                            Fecha = ((SAPbouiCOM.EditText)oForm.Items.Item("6").Specific).String;

                            int Index = 0;

                            // Asociar Pagos
                            List<string> ListaPagos = PagosMasivos.ObtenerPagosExistentes(Fecha);

                            if (ListaPagos.Count > 0)
                            {
                                SAPbouiCOM.Item oItem = oForm.Items.Item("4");
                                SAPbouiCOM.ComboBox oComboBox = ((SAPbouiCOM.ComboBox)(oItem.Specific));
                                oComboBox.Item.DisplayDesc = true;

                                // Limpiar Combobox si tiene datos y eliminar userdatasource
                                if (oComboBox.ValidValues.Count > 0)
                                {
                                    FuncionesComunes.BorrarCombo(FormUID, "4");
                                    oComboBox.DataBind.UnBind();
                                }

                                foreach (string pago in ListaPagos)
                                {
                                    oComboBox.ValidValues.Add(Index.ToString(), pago);
                                    Index++;
                                }                                
                                oComboBox.DataBind.SetBound(true, "", "udsPagos");
                            }
                            else
                            {
                                FuncionesComunes.BorrarCombo(FormUID, "4");
                                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(string.Format("No se encontraron pagos en fecha: {0}", Fecha), SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            }
                        }
                    }
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        // boton Generar
                        if (pVal.ItemUID.Equals("1"))
                        {
                            // Obtener Fecha
                            Fecha = ((SAPbouiCOM.EditText)oForm.Items.Item("6").Specific).String;

                            // Obtener pago
                            Pago = ((SAPbouiCOM.ComboBox)oForm.Items.Item("4").Specific).Value;

                            // Validar ingreso de pago
                            result = ValidarCampos(Fecha, Pago);
                            if (result.Success)
                            {
                                Pago = ((SAPbouiCOM.ComboBox)oForm.Items.Item("4").Specific).Selected.Description;
                                result = PagosMasivos.GestionarPagoMasivo(Pago);

                                if (result.Success)
                                {
                                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                }
                                else
                                {
                                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                }
                                        
                            }
                            else
                            {
                                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            }
                        }
                        // boton cancelar
                        else if (pVal.ItemUID.Equals("2"))
                        {
                            oForm.Close();
                        }
                    }                   
                }
            }
            catch
            {
            }
        }

        /// <summary>
        /// Valida el ingreso de datos al form. Devuelve True si esta correcto
        /// </summary>
        /// <param name="Fecha"></param>
        /// <param name="Pago"></param>
        /// <returns></returns>
        private static ResultMessage ValidarCampos(string Fecha, string Pago)
        {
            result = new ResultMessage();

            if (string.IsNullOrEmpty(Fecha))
            {
                result.Success = false;
                result.Mensaje = "Ingrese Fecha de pago";
            }
            else if (string.IsNullOrEmpty(Pago))
            {
                result.Success = false;
                result.Mensaje = "Seleccione Pago";
            }
            else
            {
                result.Success = true;
            }

            return result;
        }
    }
}
