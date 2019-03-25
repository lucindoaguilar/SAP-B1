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
                    //creationPackage.FormType = "FormDocSMas";
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
            SAPbouiCOM.Matrix oMatrix = null;
            try
            {
                oMatrix = oForm.Items.Item("").Specific;

            }
            catch(Exception ex)
            {
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
