using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Addon_Facturas_Proveedores.Comunes;
using Addon_Facturas_Proveedores.Conexiones;
using System.Xml;
using System.Windows.Forms;

namespace Addon_Facturas_Proveedores.ClaseFormulario
{
    class SEI_FormDat
    {       

        public SEI_FormDat(String FormUID)
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
                    if (oFormAbierto.UniqueID.Equals("SEI_DAT"))
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
                    xml = xml + "Formularios\\" + "FormDat.srf";
                    oXmlDoc.Load(xml);

                    String sXML = oXmlDoc.InnerXml.ToString();
                    SAPbouiCOM.FormCreationParams creationPackage = null;
                    creationPackage = (SAPbouiCOM.FormCreationParams)Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                    creationPackage.UniqueID = "SEI_DAT";
                    creationPackage.FormType = "SEI_DAT";
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
                result = Msj_Appl.Errores(14, "CargarXML > SEI_FormDat.cs " + ex.Message);
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Forma el formulario de documentos, con los campos necesarios.
        /// </summary>
        private void CargarFormulario(String FormUID)
        {
            SAPbouiCOM.Form oForm = Conexion_SBO.m_SBO_Appl.Forms.Item(FormUID);
            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("Item_0").Specific;
            oMatrix.Layout = SAPbouiCOM.BoMatrixLayoutType.mlt_Normal;
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
            
            // Datasource para blindear matrix
            oForm.DataSources.DataTables.Add("FILAS");
            SAPbouiCOM.DataTable dt = oForm.DataSources.DataTables.Item("FILAS");

            dt.Columns.Add("co_LineNum", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Serv", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Tot", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Imp", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Cta", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);            
            dt.Columns.Add("co_Esp", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Var", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Categ", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Catpa", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("co_Rolp", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

            Int32 Index = 0;

            foreach (Filas fil in ListaFilas.ListFilas)
            {
                dt.Rows.Add();
                dt.SetValue("co_LineNum", Index, fil.LineNum.ToString());
                dt.SetValue("co_Serv", Index, fil.Servicio);
                dt.SetValue("co_Tot", Index, String.Format(System.Globalization.CultureInfo.GetCultureInfo("es-CL"), "{0:C0}", fil.Total));
                dt.SetValue("co_Imp", Index, fil.Impuesto);
                dt.SetValue("co_Cta", Index, String.Empty);                
                dt.SetValue("co_Esp", Index, String.Empty);
                dt.SetValue("co_Var", Index, String.Empty);
                dt.SetValue("co_Categ", Index, String.Empty);
                dt.SetValue("co_Catpa", Index, String.Empty);
                dt.SetValue("co_Rolp", Index, String.Empty);
                Index++;
            }

            oMatrix.Columns.Item("col_lin").DataBind.Bind("FILAS", "co_LineNum");
            oMatrix.Columns.Item("col_Serv").DataBind.Bind("FILAS", "co_Serv");
            oMatrix.Columns.Item("col_Tot").DataBind.Bind("FILAS", "co_Tot");
            oMatrix.Columns.Item("col_Imp").DataBind.Bind("FILAS", "co_Imp");
            oMatrix.Columns.Item("col_Cta").DataBind.Bind("FILAS", "co_Cta");            
            oMatrix.Columns.Item("col_Esp").DataBind.Bind("FILAS", "co_Esp"); 
            oMatrix.Columns.Item("col_Var").DataBind.Bind("FILAS", "co_Var");
            oMatrix.Columns.Item("col_Cat").DataBind.Bind("FILAS", "co_Categ");
            oMatrix.Columns.Item("col_Catpa").DataBind.Bind("FILAS", "co_Catpa");
            oMatrix.Columns.Item("col_Rolp").DataBind.Bind("FILAS", "co_Rolp");

            oMatrix.LoadFromDataSource();
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

            try
            {
                if (pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CLICK) && pVal.ItemUID.Equals("btnCanc"))
                {
                    ListaFilas.ListFilas.Clear();
                    oForm.Close();
                }
                if (pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CLICK) && pVal.ItemUID.Equals("btnAcep"))
                {
                    result = ValidarDatos(oForm.Items.Item("Item_0").Specific);
                    if (!result.Success)
                    {
                        BubbleEvent = false;
                        Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                }
                if (!pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CLICK) && pVal.ItemUID.Equals("btnAcep"))
                {
                    AsignarDatos(oForm.Items.Item("Item_0").Specific);
                    oForm.Close();
                }
            }
            catch (Exception ex)
            {
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }        

        /// <summary>
        /// Valida el ingreso de datos por parte del usuario.
        /// </summary>
        public static ResultMessage ValidarDatos(SAPbouiCOM.Matrix oMatrix)
        {
            ResultMessage result = new ResultMessage();
            result.Success = true;
            String LineNum = String.Empty;
            String Servicio = String.Empty;
            String Cuenta = String.Empty;
            
            for (Int32 index = 1; index <= oMatrix.RowCount; index++)
            {
                LineNum = oMatrix.Columns.Item("col_lin").Cells.Item(index).Specific.Value;
                Servicio = oMatrix.Columns.Item("col_Serv").Cells.Item(index).Specific.Value;
                Cuenta = oMatrix.Columns.Item("col_Cta").Cells.Item(index).Specific.Value;

                if (String.IsNullOrEmpty(Cuenta))
                {
                    result.Success = false;
                    result.Mensaje = String.Format("Ingrese Cuenta en Linea: {0} Servicio: {1}", LineNum, Servicio);
                    return result;
                }                
            }
            return result; 
        }

        /// <summary>
        /// Asigna los datos ingresados por usuario
        /// </summary>
        public static void AsignarDatos(SAPbouiCOM.Matrix oMatrix)
        {
            for (Int32 index = 1; index <= oMatrix.RowCount; index++)
            {
                ListaFilas.ListFilas.First(d => d.LineNum == index).Cuenta = oMatrix.Columns.Item("col_Cta").Cells.Item(index).Specific.Value;
                ListaFilas.ListFilas.First(d => d.LineNum == index).Especie = oMatrix.Columns.Item("col_Esp").Cells.Item(index).Specific.Value;
                ListaFilas.ListFilas.First(d => d.LineNum == index).Variedad = oMatrix.Columns.Item("col_Var").Cells.Item(index).Specific.Value;
                ListaFilas.ListFilas.First(d => d.LineNum == index).Category = oMatrix.Columns.Item("col_Cat").Cells.Item(index).Specific.Value;
                ListaFilas.ListFilas.First(d => d.LineNum == index).CatPacking = oMatrix.Columns.Item("col_Catpa").Cells.Item(index).Specific.Value;
                ListaFilas.ListFilas.First(d => d.LineNum == index).RolPrivado = oMatrix.Columns.Item("col_Rolp").Cells.Item(index).Specific.Value;
            }
        }

        

        
    }

    
}
