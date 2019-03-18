using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Addon_Facturas_Proveedores.Conexiones;
using Addon_Facturas_Proveedores.Comunes;
using System.Xml;
using System.Windows.Forms;

namespace Addon_Facturas_Proveedores.ClaseFormulario
{
    class SEI_FormINGMH
    {
        static List<BaseFactura> listaBase = null;
        public static List<string> Campo = new List<string>();

        public SEI_FormINGMH(String FormUID, List<BaseFactura> BaseFactura, List<String> CodigoCampo)
        {
            Campo = CodigoCampo;
            CargarXML();
            CargarFormulario(FormUID, CodigoCampo);
            listaBase = new List<BaseFactura>();
            listaBase = BaseFactura;
            Addon_Facturas_Proveedores.Comunes.AsignacionMultiple.AsignadosHec.ListaAsignacionHec.Clear();            
        }

        /// <summary>
        /// Carga el formulario de ingreso múltiple a traves de su XML
        /// </summary>
        public static void CargarXML()
        {
            XmlDocument oXmlDoc = null;
            ResultMessage result = new ResultMessage();
            Boolean bFormAbierto = false;
            Int32 index;

            try
            {
                for (index = 0; index < Conexion_SBO.m_SBO_Appl.Forms.Count; index++)
                {
                    SAPbouiCOM.Form oFormAbierto = Conexion_SBO.m_SBO_Appl.Forms.Item(index);
                    if (oFormAbierto.UniqueID.Equals("SEI_INGMH"))
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
                    xml = xml + "Formularios\\" + "FormIngMultipleHec.srf";
                    oXmlDoc.Load(xml);
                    
                    String sXML = oXmlDoc.InnerXml.ToString();
                    Conexion_SBO.m_SBO_Appl.LoadBatchActions(ref sXML);
                }
                // Si no traer al frente
                else
                {
                    Conexion_SBO.m_SBO_Appl.Forms.Item(index).Select();
                }
            }
            catch (Exception ex)
            {
                result = Msj_Appl.Errores(14, "CargarXML > SEI_FormINGMH.cs " + ex.Message);
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText(result.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Carga en el formulario los objetos necesario para el ingreso múltiple
        /// Etiquetas de texto, matrices, etc.
        /// </summary>
        public static void CargarFormulario(string FormUID, List<string> CodigoCampo)
        {
            // Posicionamiento en form
            int Top = 10;
            int Left = 10;

            // Form actual
            SAPbouiCOM.Form oForm = Conexion_SBO.m_SBO_Appl.Forms.Item(FormUID);
            oForm.Freeze(true);

            #region MATRIX VARIEDADES

            //Etiqueta de texto informando: FERTILIZANTES
            SAPbouiCOM.Item oItemST = oForm.Items.Add("st_Var", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            SAPbouiCOM.StaticText oStaticText = ((SAPbouiCOM.StaticText)(oItemST.Specific));
            oStaticText.Caption = "VARIEDADES";
            oStaticText.Item.Top = Top;
            oStaticText.Item.Left = Left;
            oStaticText.Item.Width = 330;

            // Matrix de variedades
            Top += 20;
            SAPbouiCOM.Matrix oMatrix;
            SAPbouiCOM.Item oItem = oForm.Items.Add("matrixVar", SAPbouiCOM.BoFormItemTypes.it_MATRIX);

            oItem.Width = 470;
            oItem.Top = Top;
            oItem.Left = Left;
            oItem.Height = 230;
            oMatrix = oForm.Items.Item("matrixVar").Specific;
            SAPbouiCOM.Columns oColumns;
            oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;

            oMatrix.Layout = SAPbouiCOM.BoMatrixLayoutType.mlt_Normal;
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None;

            oColumn = oColumns.Add("co_Chk", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = string.Empty;
            oColumn.Editable = true;
            oColumn.Width = 22;
            oColumn.ValOff = "N";
            oColumn.ValOn = "Y";

            oColumn = oColumns.Add("co_Cod", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Código";
            oColumn.Editable = false;
            oColumn.Width = 70;

            oColumn = oColumns.Add("co_Desc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Descripción";
            oColumn.Editable = false;
            oColumn.Width = 220;

            oColumn = oColumns.Add("co_Esp", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Especie";
            oColumn.Editable = false;
            oColumn.Width = 70;

            oColumn = oColumns.Add("co_Cpo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Campo";
            oColumn.Editable = false;
            oColumn.Width = 70;   

            Top += 235;
            SAPbouiCOM.Button oButton;
            oItem = oForm.Items.Add("bt_chk", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = Left;
            oItem.Width = 60;
            oItem.Top = Top;
            oItem.Height = 16;
            oItem.FromPane = 0;
            oItem.ToPane = 0;
            oItem.Enabled = true;
            oButton = ((SAPbouiCOM.Button)(oItem.Specific));
            // set the caption
            oButton.Caption = "check";

            oForm.DataSources.DataTables.Add("VARIEDADES");
            oForm.DataSources.DataTables.Add("SEL");

            #endregion

            #region MATRIX CUARTELES Y BOTON CHECK

            //Etiqueta de texto informando: CUARTELES
            Top += 20;
            oItemST = oForm.Items.Add("st_Crt", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oStaticText = ((SAPbouiCOM.StaticText)(oItemST.Specific));
            oStaticText.Caption = "CUARTELES";
            oStaticText.Item.Top = Top;
            oStaticText.Item.Left = Left;
            oStaticText.Item.Width = 330;

            // Matrix de Cuarteles
            Top += 20;
            oItem = oForm.Items.Add("matrixCrt", SAPbouiCOM.BoFormItemTypes.it_MATRIX);

            oItem.Width = 460;
            oItem.Top = Top;
            oItem.Left = Left;
            oItem.Height = 110;
            oMatrix = oForm.Items.Item("matrixCrt").Specific;
            oColumns = oMatrix.Columns;

            oMatrix.Layout = SAPbouiCOM.BoMatrixLayoutType.mlt_Normal;
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None;

            oColumn = oColumns.Add("co_Chk", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = string.Empty;
            oColumn.Editable = true;
            oColumn.Width = 22;
            oColumn.ValOff = "N";
            oColumn.ValOn = "Y";
            oColumn.Editable = false;

            oColumn = oColumns.Add("co_Crt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Cuartel";
            oColumn.Editable = false;
            oColumn.Width = 80;

            oColumn = oColumns.Add("co_Desc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Variedad";
            oColumn.Editable = false;
            oColumn.Width = 210;

            oColumn = oColumns.Add("co_Hec", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Superficie";
            oColumn.Editable = false;
            oColumn.Width = 70;

            oColumn = oColumns.Add("co_Esp", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Especie";
            oColumn.Editable = false;
            oColumn.Width = 60;

            oColumn = oColumns.Add("co_Cpo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Campo";
            oColumn.Visible = false;

            oForm.DataSources.DataTables.Add("CUARTELES");            

            #endregion

            #region DATASOURCES

            string query = string.Empty;
            int index = 0;

            // Codigo para un campo por cuenta
            if (CodigoCampo.Count.Equals(1))
            {
                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        query = @"SELECT 'N' as ""check"", T0.""PrcCode"" as ""Code"", T0.""PrcName"" as ""Name"", T0.""U_SEI_Especie"" as ""Esp"", '" + CodigoCampo[0] + @"' as ""Campo"" FROM OPRC T0 INNER JOIN ""@SEI_DMDC"" T1 ON T0.""PrcCode"" = T1.""U_SEI_VAR"" WHERE T0.""DimCode""=2 AND T0.""Active""='Y' AND T1.""Code"" = '" + CodigoCampo[0] + @"' GROUP BY T0.""PrcCode"", T0.""PrcName"", T0.""U_SEI_Especie"" ORDER BY T0.""U_SEI_Especie"", T0.""PrcCode""";
                        break;
                    default:
                        query = @"SELECT 'N' as check, T0.PrcCode as Code, T0.PrcName as Name, T0.U_SEI_Especie as Esp, '" + CodigoCampo[0] + @"' as Campo FROM OPRC T0 INNER JOIN @SEI_DMDC T1 ON T0.PrcCode = T1.U_SEI_VAR WHERE T0.DimCode=2 AND T0.Active='Y' AND T1.Code = '" + CodigoCampo[0] + @"' GROUP BY T0.PrcCode, T0.PrcName, T0.U_SEI_Especie ORDER BY T0.U_SEI_Especie, T0.PrcCode";
                        break;
                }

                
            }
            // Codigo para mas de un campo
            else
            {
                query = "SELECT * FROM (";
                foreach (string sCampo in CodigoCampo)
                {
                    if (!index.Equals(0))
                    {
                        query += " UNION ";
                    }

                    switch (Conexion_SBO.m_oCompany.DbServerType)
                    {
                        case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                            query += @"SELECT 'N' as ""check"", T0.""PrcCode"" as ""Code"", T0.""PrcName"" as ""Name"", T0.""U_SEI_Especie"" as ""Esp"", '" + sCampo + @"' as ""Campo"" FROM OPRC T0 INNER JOIN ""@SEI_DMDC"" T1 ON T0.""PrcCode"" = T1.""U_SEI_VAR"" WHERE T0.""DimCode""=2 AND T0.""Active""='Y' AND T1.""Code"" = '" + sCampo + "'";
                            break;
                        default:
                            query += @"SELECT 'N' as check, T0.PrcCode as Code, T0.PrcName as Name, T0.U_SEI_Especie as Esp, '" + sCampo + @"' as Campo FROM OPRC T0 INNER JOIN @SEI_DMDC T1 ON T0.PrcCode = T1.U_SEI_VAR WHERE T0.DimCode = 2 AND T0.Active = 'Y' AND T1.Code = '" + sCampo + "'";
                            break;
                    }

                    
                    index++;
                }

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        query += @" ) GROUP BY ""check"", ""Code"", ""Name"", ""Esp"",""Campo"" ORDER BY ""Campo"",""Esp"", ""Code""";
                        break;
                    default:
                        query += @" ) GROUP BY check, Code, Name, Esp,Campo ORDER BY Campo,Esp , Code";
                        break;
                }
                
            }

            
            oForm.DataSources.DataTables.Item("VARIEDADES").ExecuteQuery(query);
            oMatrix = oForm.Items.Item("matrixVar").Specific;
            oMatrix.Columns.Item("co_Chk").DataBind.Bind("VARIEDADES", "check");
            oMatrix.Columns.Item("co_Cod").DataBind.Bind("VARIEDADES", "Code");
            oMatrix.Columns.Item("co_Desc").DataBind.Bind("VARIEDADES", "Name");
            oMatrix.Columns.Item("co_Esp").DataBind.Bind("VARIEDADES", "Esp");
            oMatrix.Columns.Item("co_Cpo").DataBind.Bind("VARIEDADES", "Campo");
            oMatrix.Clear();
            oMatrix.LoadFromDataSource();

            oForm.DataSources.DataTables.Item("SEL").Columns.Add("Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250);

            #endregion

            oForm.Freeze(false);
        }

        /// <summary>
        /// Interviene los eventos de los item del form.
        /// </summary>
        public static void m_SBO_Appl_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = Conexion_SBO.m_SBO_Appl.Forms.Item(FormUID);
            ResultMessage rslt = new ResultMessage();

            try
            {
                if (pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CLICK) && pVal.ItemUID.Equals("bt_chk"))
                {
                    EjecutarChecking(oForm);
                }
                if (pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CLICK) && pVal.ItemUID.Equals("bt_SLR"))
                {
                    //AsignadosHec.ListaAsignacionHec.Clear();
                    oForm.Close();
                }
                if (pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_CLICK) && pVal.ItemUID.Equals("bt_ASG"))
                {
                    rslt = ValidarDatosAsignados(oForm);

                    if (rslt.Success)
                    {
                        //rslt = AsignarDatos(oForm);
                        if (rslt.Success)
                        {
                            oForm.Close();
                            Conexion_SBO.m_SBO_Appl.StatusBar.SetText(rslt.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                        else
                        {
                            BubbleEvent = false;
                            Conexion_SBO.m_SBO_Appl.StatusBar.SetText(rslt.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                    }
                    else
                    {
                        BubbleEvent = false;
                        Conexion_SBO.m_SBO_Appl.StatusBar.SetText(rslt.Mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                }
                if (pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) && pVal.ItemUID.Equals("matrixVar") && pVal.ColUID.Equals("co_Chk"))
                {
                    DataBindCuarteles(oForm);
                }
                if (pVal.BeforeAction && pVal.EventType.Equals(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) && pVal.ItemUID.Equals("matrixCrt") && pVal.ColUID.Equals("co_Chk"))
                {
                    SetButtonCheckCuarteles(oForm);
                }                
            }
            catch (Exception ex)
            {
                Msj_Appl.Errores(14, "m_SBO_Appl_ItemEvent() > SEI_FormINGMH.cs " + ex.Message);
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("m_SBO_Appl_ItemEvent() > SEI_FormINGMH.cs " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Checkea y descheckea todas las columnas de la matrix variedades
        /// </summary>
        public static void EjecutarChecking(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Button oButton = oForm.Items.Item("bt_chk").Specific;
            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("matrixVar").Specific;
            SAPbouiCOM.Matrix oMatrixCrt = oForm.Items.Item("matrixCrt").Specific;            
            
            oForm.Freeze(true);

            // Check
            if (oButton.Caption.Equals("check"))
            {
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    SAPbouiCOM.CheckBox chk = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Chk").Cells.Item(i).Specific;
                    chk.Checked = true;                    
                }                
                DataBindCuarteles(oForm);
                oButton.Caption = "uncheck";
            }
            // Uncheck
            else
            {
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    SAPbouiCOM.CheckBox chk = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Chk").Cells.Item(i).Specific;
                    chk.Checked = false;
                }
                oMatrixCrt.Clear();
                oButton.Caption = "check";                
            }

            oForm.Freeze(false);
        }        

        /// <summary>
        /// Trae los datos para grilla de cuarteles segun variedades asignadas (check)
        /// </summary>
        public static void DataBindCuarteles(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("matrixVar").Specific;
            string query = string.Empty;
            List<string> Chekeados = new List<string>();
            string Desc = string.Empty;

            // Recorremos la matrix de variedades y obtenemos los valores con check
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                SAPbouiCOM.CheckBox chk = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Chk").Cells.Item(i).Specific;
                if (chk.Checked)
                {
                    Chekeados.Add(oMatrix.Columns.Item("co_Cod").Cells.Item(i).Specific.Value);
                    Desc = oMatrix.Columns.Item("co_Desc").Cells.Item(i).Specific.Value.ToString();                    
                }
            }

            // Si hay variedades chekeadas asignar cuarteles
            if (Chekeados.Count > 0)
            {
                int index = 0;
                foreach (string sCampo in Campo)
                {
                    foreach (string var in Chekeados)
                    {
                        if (!index.Equals(0))
                        {
                            query += " UNION ";
                        }

                        switch (Conexion_SBO.m_oCompany.DbServerType)
                        {
                            case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                                query += @"SELECT 'Y' as ""check"", ""U_SEI_CRT"" as ""Cuartel"", ""U_SEI_VAR"" as ""Variedad"", ""U_SEI_HEC"" as ""Hec"", ""U_SEI_ESP"" as ""Esp"", ""Code"" as ""Campo"" FROM ""@SEI_DMDC"" WHERE ""Code"" = '" + sCampo + @"' AND ""U_SEI_VAR"" = '" + var + @"' AND ""U_SEI_EST""='Activo'";
                                break;
                            default:
                                query += @"SELECT 'Y' as check, U_SEI_CRT as Cuartel, U_SEI_VAR as Variedad, U_SEI_HEC as Hec, U_SEI_ESP as Esp, Code as Campo FROM @SEI_DMDC WHERE Code = '" + sCampo + @"' AND U_SEI_VAR = '" + var + @"' AND U_SEI_EST='Activo'";
                                break;
                        }                        

                        index++;
                    }
                }

                oForm.DataSources.DataTables.Item("CUARTELES").ExecuteQuery(query);
                oMatrix = oForm.Items.Item("matrixCrt").Specific;
                oMatrix.Columns.Item("co_Chk").DataBind.Bind("CUARTELES", "check");
                oMatrix.Columns.Item("co_Crt").DataBind.Bind("CUARTELES", "Cuartel");
                oMatrix.Columns.Item("co_Desc").DataBind.Bind("CUARTELES", "Variedad");
                oMatrix.Columns.Item("co_Hec").DataBind.Bind("CUARTELES", "Hec");
                oMatrix.Columns.Item("co_Esp").DataBind.Bind("CUARTELES", "Esp");
                oMatrix.Columns.Item("co_Cpo").DataBind.Bind("CUARTELES", "Campo");
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();

                // Si hay variedades chekeades setear boton check variedades
                SAPbouiCOM.Button oButton = oForm.Items.Item("bt_chk").Specific;
                oButton.Caption = "uncheck";
            }
            else
            {
                oMatrix = oForm.Items.Item("matrixCrt").Specific;
                oMatrix.Clear();
                // Si no hay variedades chekeades setear boton check variedades y cuarteles
                SAPbouiCOM.Button oButton = oForm.Items.Item("bt_chk").Specific;
                oButton.Caption = "check";
            }
        }

        /// <summary>
        /// Revisa el estado de la matrix cuarteles para dar valor correcto al boton check cuarteles
        /// </summary>
        public static void SetButtonCheckCuarteles(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("matrixCrt").Specific;
            SAPbouiCOM.Button oButton = oForm.Items.Item("bt_chkC").Specific;
            bool Chekeados = false;

            // Recorremos la matrix de cuarteles para saber si hay chekeados
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                SAPbouiCOM.CheckBox chk = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Chk").Cells.Item(i).Specific;
                if (chk.Checked)
                {
                    Chekeados = true;
                    break;
                }
            }

            // si hay valores chekeados cambiar button
            if (Chekeados)
            {
                oButton.Caption = "uncheck";
            }
            else
            {
                oButton.Caption = "check";
            }
        }

        /// <summary>
        /// Valida los datos del formulario ingreso múltiple
        /// Retorna True si los datos son validos y se puede asignar
        /// Retorna False si faltan datos y no se puede asignar
        /// </summary>
        public static ResultMessage ValidarDatosAsignados(SAPbouiCOM.Form oForm)
        {
            ResultMessage rslt = new ResultMessage();
            SAPbouiCOM.Matrix oMatrix = null;
            string Valor = string.Empty;
            bool check = false;

            try
            {
                // Validar Seleccion de Matrix Variedades
                oMatrix = oForm.Items.Item("matrixVar").Specific;
                check = false;

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    check = ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Chk").Cells.Item(i).Specific).Checked;
                    if (check == true)
                    {
                        break;
                    }
                }

                if (check == false)
                {
                    rslt.Success = false;
                    rslt.Mensaje = "Seleccione una variedad a asignar.";
                    return rslt;
                }

                // Validar Seleccion de Matrix Cuarteles
                oMatrix = oForm.Items.Item("matrixCrt").Specific;
                check = false;

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    check = ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Chk").Cells.Item(i).Specific).Checked;
                    if (check == true)
                    {
                        break;
                    }
                }

                if (check == false)
                {
                    rslt.Success = false;
                    rslt.Mensaje = "Seleccione un cuartel a asignar.";
                    return rslt;
                }

                rslt.Success = true;
                return rslt;

            }
            catch (Exception ex)
            {
                rslt.Success = false;
                rslt.Mensaje = ex.Message;
                return rslt;
            }
        }

        /// <summary>
        /// Asigna los datos en pantalla a la lista asignacion
        /// Retorna true si los datos fueron asignados
        /// Retorna false si no fueron asignados
        /// </summary>
        /*
        public static ResultMessage AsignarDatos(SAPbouiCOM.Form oForm)
        {
            ResultMessage rslt = new ResultMessage();
            AsignacionHec asig = null;            
            SAPbouiCOM.Matrix oMatrixV = null;
            SAPbouiCOM.Matrix oMatrixC = null;
            string sVariedad = string.Empty;
            string sVariedadCrt = string.Empty;
            List<Variedad> variedades = new List<Variedad>();
            Variedad var = null;
            double totalHec = 0;
            double totalHecSextafrut = 0;

            try
            {
                Conexion_SBO.m_SBO_Appl.SetStatusBarMessage("Calculando asignacion por hecareaje. Espere un momento", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                
                // Obtener Datos
                oMatrixV = oForm.Items.Item("matrixVar").Specific;
                oMatrixC = oForm.Items.Item("matrixCrt").Specific;

                // Recorrer la matrix cuarteles y obtener sumatoria de hectareaje y variedad
                for (int i = 1; i <= oMatrixC.RowCount; i++)
                {
                    bool checkC = ((SAPbouiCOM.CheckBox)oMatrixC.Columns.Item("co_Chk").Cells.Item(i).Specific).Checked;
                    // Si el cuartel esta chekeado suma y asigna variedades
                    if (checkC == true)
                    {                        
                        var = new Variedad();
                        var.Codigo = oMatrixC.Columns.Item("co_Desc").Cells.Item(i).Specific.Value.ToString();
                        var.Hectareaje = Convert.ToDouble(oMatrixC.Columns.Item("co_Hec").Cells.Item(i).Specific.Value.ToString().Replace(".", ","));
                        totalHec += Convert.ToDouble(oMatrixC.Columns.Item("co_Hec").Cells.Item(i).Specific.Value.ToString().Replace(".", ","));
                        var.Especie = oMatrixC.Columns.Item("co_Esp").Cells.Item(i).Specific.Value.ToString();
                        var.Campo = oMatrixC.Columns.Item("co_Cpo").Cells.Item(i).Specific.Value.ToString();
                        variedades.Add(var);
                    }
                }
                                
                // Obtener hectareaje por variedad
                var query =
                            (from p in variedades                            
                            group p by new 
                            {
                                p.Codigo,
                                p.Especie,
                                p.Campo
                            }  into g    
                            select new { Variedad = g.Key.Codigo, Especie =  g.Key.Especie, Campo = g.Key.Campo, Total = g.Sum(p => p.Hectareaje) }).ToList();
                
                var query1 =
                           (from p in query
                            group p by new
                            {
                                p.Campo
                            } into g
                            select new { Campo = g.Key.Campo, TotalHectareasCampo = g.Sum(p => p.Total) }).ToList();

                // Para sextrafrut se debe sumar por los 2 campos
                foreach (var v in query1)
                {
                    if (v.Campo.Equals("020") || v.Campo.Equals("021"))
                    {
                        totalHecSextafrut += v.TotalHectareasCampo;
                    }
                }

                // recorrer las lineas de la base de factura
                foreach(BaseFactura bse in listaBase)
                {
                    // recorrer las variedades asignadas con su total de hecatreaje
                    foreach (var item in query)
                    {
                        foreach (string Campo in bse.Campo)
                        {
                            if (Campo == item.Campo)
                            {                                
                                asig = new AsignacionHec();
                                if (Campo.Equals("020") || Campo.Equals("021"))
                                {
                                    asig.Total = (bse.Total / totalHecSextafrut) * item.Total;
                                }
                                else
                                {
                                    asig.Total = (bse.Total / query1.Where(s => s.Campo == Campo).Select(s => s.TotalHectareasCampo).Single()) * item.Total;
                                }
                                //asig.Total = (bse.Total / totalHec) * item.Total;
                                asig.IndImpuesto = bse.IndImpuesto;
                                asig.Descripcion = bse.Descripcion;
                                asig.CuentaMayor = bse.CuentaMayor;
                                //asig.BaseEntry = bse.BaseEntry;
                                //asig.BaseLine = bse.BaseLine;
                                //asig.BaseRef = bse.BaseRef;
                                //asig.BaseType = bse.BaseType;
                                asig.Variedad = item.Variedad;
                                asig.Especie = item.Especie;
                                asig.Category = bse.Category;
                                asig.CodMaquinaria = bse.CodMaquinaria;
                                asig.CodMant = bse.CodMantencion;
                                asig.FechaMantencion = bse.FechaMantencion;
                                asig.Horometro = bse.Horometro;

                                //AsignacionMultiple.AsignacionHec.ListaAsignacionHec.Add(asig);
                            }
                        }
                    }
                }                

                rslt.Success = true;
                rslt.Mensaje = string.Format("Datos asignados correctamente. Total : {0}", AsignadosHec.ListaAsignacionHec.Count());
                return rslt;
            }
            catch (Exception ex)
            {
                rslt.Success = false;
                rslt.Mensaje = ex.Message;
                return rslt;
            }
        }    */    

    }
}
