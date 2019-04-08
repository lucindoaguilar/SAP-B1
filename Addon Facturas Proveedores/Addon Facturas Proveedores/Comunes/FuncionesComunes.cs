using System;
using System.Xml;
using Addon_Facturas_Proveedores.Conexiones;
using Addon_Facturas_Proveedores.Documento;
using System.Text.RegularExpressions;
using RestSharp;
using System.Configuration;

namespace Addon_Facturas_Proveedores.Comunes
{
    /// <summary>
    /// Clase que contiene un conjunto de funcionalidad compartidad por todas las clases de la aplicacion.
    /// </summary>
    public class FuncionesComunes
    {
        #region Metodos

        /// <summary>
        /// Metodo para liberar objetos COM de memoria
        /// </summary>
        /// <param name="objeto">Objeto que se desea liberar de memoria</param> n 
        public static void LiberarObjetoGenerico(Object objeto)
        {
            try
            {
                if (objeto != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto);//Desacoplamos el objeto de la aplicacion
                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
                Msj_Appl.Errores(14, "FucionesComunes.cs -> LiberarObjetoGenerico: " + ex.Message);//Mensaje de errors
            }
        }

        /// <summary>
        /// Función que obtiene el token desde OADM
        /// </summary>
        public static String ObtenerToken()
        {
            try
            {
                String Token = String.Empty;
                String query = String.Empty;
                SAPbobsCOM.Recordset oRec = null;

                oRec = (SAPbobsCOM.Recordset)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        query = "SELECT \"U_SEI_TOKEN\" FROM OADM";
                        break;
                    default:
                        query = "SELECT U_SEI_TOKEN FROM OADM";
                        break;
                }
                
                oRec.DoQuery(query);

                if (!oRec.EoF)
                {
                    Token = oRec.Fields.Item(0).Value.ToString();
                }

                FuncionesComunes.LiberarObjetoGenerico(oRec);
                return Token;
            }
            catch (Exception ex)
            {
                Comunes.Msj_Appl.Errores(14, "FuncionesComunes_ObtenerToken->" + ex.Message);
                return String.Empty;
            }
        }

        /// <summary>
        /// Función que obtiene el Rut de la sociedad de la BD
        /// </summary>
        public static String ObtenerRut()
        {
            try
            {
                String Rut = String.Empty;
                SAPbobsCOM.Recordset oRec = null;
                String query = String.Empty;

                oRec = (SAPbobsCOM.Recordset)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        query = "SELECT \"RevOffice\" FROM OADM";
                        break;
                    default:
                        query = "SELECT RevOffice FROM OADM";
                        break;
                }
                
                oRec.DoQuery(query);

                if (!oRec.EoF)
                {
                    Rut = ValidarRut(oRec.Fields.Item(0).Value.ToString());
                }

                FuncionesComunes.LiberarObjetoGenerico(oRec);
                return Rut;
            }
            catch(Exception ex)
            {
                Comunes.Msj_Appl.Errores(14, "FuncionesComunes_ObtenerRut->" + ex.Message);
                return String.Empty;
            }
        }

        /// <summary>
        /// Función que obtiene el Cardcode del socio de negocio de la BD
        /// </summary>
        public static String ObtenerCardcode(String Rut, out String ProveedorNR)
        {
            ProveedorNR = "N";
            try
            {
                Rut = Rut.Replace(".", String.Empty);
                String Cardcode = String.Empty;
                SAPbobsCOM.Recordset oRec = null;
                String query = String.Empty;

                oRec = (SAPbobsCOM.Recordset)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        query = "SELECT \"CardCode\", \"QryGroup1\" FROM OCRD WHERE \"CardType\"='S' AND \"LicTradNum\"='" + Rut + "' ";
                        break;
                    default:
                        query = "SELECT CardCode, QryGroup1 FROM OCRD WHERE CardType = 'S' AND LicTradNum = '" + Rut + "' ";
                        break;
                }
                
                oRec.DoQuery(query);

                if (!oRec.EoF)
                {
                    Cardcode = oRec.Fields.Item(0).Value.ToString();
                    ProveedorNR = oRec.Fields.Item(1).Value.ToString();
                }

                FuncionesComunes.LiberarObjetoGenerico(oRec);
                return Cardcode;                
            }
            catch(Exception ex)
            {
                Comunes.Msj_Appl.Errores(14, "FuncionesComunes_ObtenerCardcode->" + ex.Message);
                return String.Empty;
            }
        }

        /// <summary>
        /// Función que obtiene el recinto par acuse de recepcion de mercaderias de la BD
        /// </summary>
        public static String ObtenerRecinto()
        {
            try
            {
                String Recinto = String.Empty;
                SAPbobsCOM.Recordset oRec = null;
                String query = String.Empty;

                oRec = (SAPbobsCOM.Recordset)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        query = "SELECT \"U_SEI_RECINTO\" FROM OADM";
                        break;
                    default:
                        query = "SELECT U_SEI_RECINTO FROM OADM";
                        break;
                }
                
                oRec.DoQuery(query);

                if (!oRec.EoF)
                {
                    Recinto = oRec.Fields.Item(0).Value.ToString();
                }

                FuncionesComunes.LiberarObjetoGenerico(oRec);
                return Recinto;
            }
            catch (Exception ex)
            {
                Comunes.Msj_Appl.Errores(14, "FuncionesComunes_ObtenerRecinto->" + ex.Message);
                return String.Empty;
            }
        }

        /// <summary>
        /// Función que obtiene el CardName del socio de negocio de la BD
        /// </summary>
        public static String ObtenerCardName(String Rut, String CardCode)
        {
            try
            {
                Rut = Rut.Replace(".", String.Empty);
                String CardName = String.Empty;
                SAPbobsCOM.Recordset oRec = null;
                String query = String.Empty;

                oRec = (SAPbobsCOM.Recordset)Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        query = "SELECT \"CardName\" FROM OCRD WHERE \"LicTradNum\"='" + Rut + "' AND \"CardCode\" = '" + CardCode + "'";
                        break;
                    default:
                        query = "SELECT CardName FROM OCRD WHERE LicTradNum = '" + Rut + "' AND CardCode = '" + CardCode + "' ";
                        break;
                }
                
                oRec.DoQuery(query);

                if (!oRec.EoF)
                {
                    CardName = oRec.Fields.Item(0).Value.ToString();
                }

                FuncionesComunes.LiberarObjetoGenerico(oRec);
                return CardName;
            }
            catch (Exception ex)
            {
                Comunes.Msj_Appl.Errores(14, "FuncionesComunes_ObtenerCardName->" + ex.Message);
                return String.Empty;
            }
        }

        /// <summary>
        /// Funcion que devuelve el tipo de objeto por tipo de documento
        /// </summary>
        public static Int32 ObtenerObjType(String TipoDoc)
        {
            Int32 ObjType = 0;

            switch (TipoDoc)
            {
                case "33":
                    ObjType = 18;
                    break;
                case "34":
                    ObjType = 18;
                    break;
                case "56":
                    ObjType = 18;
                    break;
                case "52":
                    ObjType = 20;
                    break;
                case "61":
                    ObjType = 19;
                    break;
            }

            return ObjType;
        }

        /// <summary>
        /// Función que valida un rut y su dígito verificador.
        /// </summary>
        public static String ValidarRut(String rut)
        {
            String rutval = "";
            Int32 rutleng = 0;
            String digVerf = "";
            try
            {
                rutval = rut.Replace(".", "");
                rutval = rutval.Replace("-", "");
                rutleng = rutval.Length;
                digVerf = rutval.Substring(rutleng - 1, 1).ToUpper();
                rutval = rutval.Substring(0, rutleng - 1) + '-'; //+ rutval.Substring(rutleng - 1, 1).ToUpper();

                if (digVerf.Equals("K"))
                {
                    rutval = rutval + digVerf.ToUpper();
                }
                else
                {
                    rutval = rutval + digVerf;
                }

                return rutval;
            }
            catch (Exception ex)
            {
                Comunes.Msj_Appl.Errores(14, "FuncionesComunes_ValidarRut->" + ex.Message);
                return string.Empty;
            }
        }

        /// <summary>
        /// Obtiene el estado sii en palabras
        /// </summary>
        public static String ObtenerEstadoSII(String Estado)
        {
            switch (Estado)
            {
                case "1": Estado = "Pendiente de envio al sii";
                    break;
                case "2": Estado = "Enviado al sii";
                    break;
                case "3": Estado = "Error al enviar";
                    break;
                case "4": Estado = "Aceptado por el sii";
                    break;
                case "5": Estado = "Aceptado con reparos por el sii";
                    break;
                case "6": Estado = "Rechazado por el sii";
                    break;
                case "7": Estado = "Pendiente de consulta en el sii";
                    break;
                case "8": Estado = "Error al consultar en el sii";
                    break;
            }

            return Estado;
        }

        /// <summary>
        /// Obtiene la forma de pago en palabras
        /// </summary>
        public static String ObtenerFormaPago(String Pago)
        {
            if (Pago.Equals("-1") || Pago.Equals("0"))
            {
                Pago = "No Definido";
            }
            else if (Pago.Equals("1"))
            {
                Pago = "Contado";
            }
            else if (Pago.Equals("2"))
            {
                Pago = "Crédito";
            }
            else if (Pago.Equals("3"))
            {
                Pago = "Sin costo";
            }
            else
            {
                Pago = "No Definido";
            }

            return Pago;
        }

        /// Función que retorna un objeto de tipo DTE, con los campos que vienen en un documento.
        /// </summary>
        public static ResultMessage ObtenerDTE(String decodedStringXml)
        {
            ResultMessage rslt = new ResultMessage();
            DTE objDTE = new DTE();
            
            try
            {
                // crear documento xml para obtener datos
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(decodedStringXml);

                // namespace del S.I.I.
                XmlNamespaceManager ns = new XmlNamespaceManager(xmlDoc.NameTable);
                ns.AddNamespace("sii", "http://www.sii.cl/SiiDte");

                // NODO IDENTIFICACION DEL DOCUMENTO
                #region IDDOC

                String PathIdDoc = "sii:DTE/sii:Documento/sii:Encabezado/sii:IdDoc";
                XmlNode IdentificacionDoc = xmlDoc.SelectSingleNode(PathIdDoc, ns);

                foreach (XmlNode childNode in IdentificacionDoc)
                {
                    switch (childNode.Name)
                    {
                        case "TipoDTE":
                            objDTE.IdDoc.TipoDTE = childNode.InnerText;
                            break;
                        case "Folio":
                            objDTE.IdDoc.Folio = Int64.Parse(childNode.InnerText);
                            break;
                        case "FchEmis":
                            objDTE.IdDoc.FchEmis = childNode.InnerText;
                            break;
                        case "IndNoRebaja":
                            objDTE.IdDoc.IndNoRebaja = Int32.Parse(childNode.InnerText);
                            break;
                        case "TipoDespacho":
                            objDTE.IdDoc.TipoDespacho = Int32.Parse(childNode.InnerText);
                            break;
                        case "IndTraslado":
                            objDTE.IdDoc.IndTraslado = Int32.Parse(childNode.InnerText);
                            break;
                        case "TpoImpresion":
                            objDTE.IdDoc.TpoImpresion = childNode.InnerText;
                            break;
                        case "IndServicio":
                            objDTE.IdDoc.IndServicio = Int32.Parse(childNode.InnerText);
                            break;
                        case "MntBruto":
                            objDTE.IdDoc.MntBruto = Int32.Parse(childNode.InnerText);
                            break;
                        case "FmaPago":
                            objDTE.IdDoc.FmaPago = Int32.Parse(childNode.InnerText);
                            break;
                        case "FmaPagExp":
                            objDTE.IdDoc.FmaPagExp = Int32.Parse(childNode.InnerText);
                            break;
                        case "FchCancel":
                            objDTE.IdDoc.FchCancel = childNode.InnerText;
                            break;
                        case "MntCancel":
                            objDTE.IdDoc.MntCancel = Int64.Parse(childNode.InnerText);
                            break;
                        case "SaldoInsol":
                            objDTE.IdDoc.SaldoInsol = Int64.Parse(childNode.InnerText);
                            break;
                        case "MntPagos":
                            MntPagos mnt = new MntPagos();
                            foreach (XmlNode child in childNode.ChildNodes)
                            {                                
                                if (child.Name.Equals("FchPago")) { mnt.FchPago = child.InnerText; }
                                else if (child.Name.Equals("MntPago")) { mnt.MntPago = Int64.Parse(child.InnerText); }
                                else if (child.Name.Equals("GlosaPagos")) { mnt.GlosaPagos = child.InnerText; }                                
                            }
                            objDTE.IdDoc.MntPagos.Add(mnt);
                            break;
                        case "PeriodoDesde":
                            objDTE.IdDoc.PeriodoDesde = childNode.InnerText;
                            break;
                        case "PeriodoHasta":
                            objDTE.IdDoc.PeriodoHasta = childNode.InnerText;
                            break;
                        case "MedioPago":
                            objDTE.IdDoc.MedioPago = childNode.InnerText;
                            break;
                        case "TipoCtaPago":
                            objDTE.IdDoc.TipoCtaPago = childNode.InnerText;
                            break;
                        case "NumCtaPago":
                            objDTE.IdDoc.NumCtaPago = childNode.InnerText;
                            break;
                        case "BcoPago":
                            objDTE.IdDoc.BcoPago = childNode.InnerText;
                            break;
                        case "TermPagoCdg":
                            objDTE.IdDoc.TermPagoCdg = childNode.InnerText;
                            break;
                        case "TermPagoGlosa":
                            objDTE.IdDoc.TermPagoGlosa = childNode.InnerText;
                            break;
                        case "TermPagoDias":
                            objDTE.IdDoc.TermPagoDias = childNode.InnerText;
                            break;
                        case "FchVenc":
                            objDTE.IdDoc.FchVenc = childNode.InnerText;
                            break;
                    }
                }

                #endregion

                // NODO EMISOR DEL DOCUMENTO
                #region EMISOR

                String PathEmisor = "sii:DTE/sii:Documento/sii:Encabezado/sii:Emisor";
                XmlNode Emisor = xmlDoc.SelectSingleNode(PathEmisor, ns);

                foreach (XmlNode childNode in Emisor)
                {
                    switch (childNode.Name)
                    {
                        case "RUTEmisor":
                            objDTE.Emisor.RUTEmisor = childNode.InnerText;
                            break;
                        case "RznSoc":
                            objDTE.Emisor.RznSoc = childNode.InnerText;
                            break;
                        case "GiroEmis":
                            objDTE.Emisor.GiroEmis = childNode.InnerText;
                            break;
                        case "Telefono":
                            objDTE.Emisor.Telefono = childNode.InnerText;
                            break;
                        case "CorreoEmisor":
                            objDTE.Emisor.CorreoEmisor = childNode.InnerText;
                            break;
                        case "Acteco":
                            objDTE.Emisor.Acteco = childNode.InnerText;
                            break;
                        case "GuiaExport":
                            foreach (XmlNode child in childNode.ChildNodes)
                            {
                                if (child.Name.Equals("CdgTraslado")) { objDTE.Emisor.CdgTraslado = Int32.Parse(child.InnerText); }
                                else if (child.Name.Equals("FolioAut")) { objDTE.Emisor.FolioAut = Int32.Parse(child.InnerText); }
                                else if (child.Name.Equals("FchAut")) { objDTE.Emisor.FchAut = child.InnerText; }
                            }
                            break;
                        case "Sucursal":
                            objDTE.Emisor.Sucursal = childNode.InnerText;
                            break;
                        case "CdgSIISucur":
                            objDTE.Emisor.CdgSIISucur = childNode.InnerText;
                            break;
                        case "DirOrigen":
                            objDTE.Emisor.DirOrigen = childNode.InnerText;
                            break;
                        case "CmnaOrigen":
                            objDTE.Emisor.CmnaOrigen = childNode.InnerText;
                            break;
                        case "CiudadOrigen":
                            objDTE.Emisor.CiudadOrigen = childNode.InnerText;
                            break;
                        case "CdgVendedor":
                            objDTE.Emisor.CdgVendedor = childNode.InnerText;
                            break;
                        case "IdAdicEmisor":
                            objDTE.Emisor.IdAdicEmisor = childNode.InnerText;
                            break;
                    }
                }

                #endregion

                // NODO RECEPTOR DEL DOCUMENTO
                #region RECEPTOR

                String PathReceptor = "sii:DTE/sii:Documento/sii:Encabezado/sii:Receptor";
                XmlNode Receptor = xmlDoc.SelectSingleNode(PathReceptor, ns);

                foreach (XmlNode childNode in Receptor)
                {
                    switch (childNode.Name)
                    {
                        case "RUTRecep":
                            objDTE.Receptor.RUTRecep = childNode.InnerText;
                            break;
                        case "CdgIntRecep":
                            objDTE.Receptor.CdgIntRecep = childNode.InnerText;
                            break;
                        case "RznSocRecep":
                            objDTE.Receptor.RznSocRecep = childNode.InnerText;
                            break;
                        case "Extranjero":
                            foreach (XmlNode child in childNode.ChildNodes)
                            {
                                if (child.Name.Equals("NumId")) { objDTE.Receptor.NumId = child.InnerText; }
                                else if (child.Name.Equals("Nacionalidad")) { objDTE.Receptor.Nacionalidad = child.InnerText; }
                                else if (child.Name.Equals("IdAdicRecep")) { objDTE.Receptor.IdAdicRecep = child.InnerText; }
                            }
                            break;
                        case "GiroRecep":
                            objDTE.Receptor.GiroRecep = childNode.InnerText;
                            break;
                        case "Contacto":
                            objDTE.Receptor.Contacto = childNode.InnerText;
                            break;
                        case "CorreoRecep":
                            objDTE.Receptor.CorreoRecep = childNode.InnerText;
                            break;
                        case "DirRecep":
                            objDTE.Receptor.DirRecep = childNode.InnerText;
                            break;
                        case "CmnaRecep":
                            objDTE.Receptor.CmnaRecep = childNode.InnerText;
                            break;
                        case "CiudadRecep":
                            objDTE.Receptor.CiudadRecep = childNode.InnerText;
                            break;
                        case "DirPostal":
                            objDTE.Receptor.DirPostal = childNode.InnerText;
                            break;
                        case "CmnaPostal":
                            objDTE.Receptor.CmnaPostal = childNode.InnerText;
                            break;
                        case "CiudadPostal":
                            objDTE.Receptor.CiudadPostal = childNode.InnerText;
                            break;
                    }
                }


                #endregion

                // NODO TRANSPORTE DEL DOCUMENTO
                #region TRANSPORTE
                
                String PathTransporte = "sii:DTE/sii:Documento/sii:Encabezado/sii:Transporte";
                XmlNode Transporte = xmlDoc.SelectSingleNode(PathTransporte, ns);

                if (Transporte != null)
                {
                    foreach (XmlNode childNode in Transporte)
                    {
                        switch (childNode.Name)
                        {
                            case "Patente":
                                objDTE.Transporte.Patente = childNode.InnerText;
                                break;
                            case "RUTTrans":
                                objDTE.Transporte.RUTTrans = childNode.InnerText;
                                break;
                            case "Chofer":
                                foreach (XmlNode child in childNode.ChildNodes)
                                {
                                    if (child.Name.Equals("RUTChofer")) { objDTE.Transporte.RUTChofer = child.InnerText; }
                                    else if (child.Name.Equals("NombreChofer")) { objDTE.Transporte.NombreChofer = child.InnerText; }
                                }
                                break;
                            case "DirDest":
                                objDTE.Transporte.DirDest = childNode.InnerText;
                                break;
                            case "CmnaDest":
                                objDTE.Transporte.CmnaDest = childNode.InnerText;
                                break;
                            case "CiudadDest":
                                objDTE.Transporte.CiudadDest = childNode.InnerText;
                                break;
                            case "Aduana":
                                foreach (XmlNode childNodeAduana in childNode)
                                {
                                    switch (childNodeAduana.Name)
                                    {
                                        case "CodModVenta":
                                            objDTE.Transporte.Aduana.CodModVenta = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "CodClauVenta":
                                            objDTE.Transporte.Aduana.CodClauVenta = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "TotClauVenta":
                                            objDTE.Transporte.Aduana.TotClauVenta = Double.Parse(childNodeAduana.InnerText.Replace(".", ","));
                                            break;
                                        case "CodViaTransp":
                                            objDTE.Transporte.Aduana.CodViaTransp = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "NombreTransp":
                                            objDTE.Transporte.Aduana.NombreTransp = childNodeAduana.InnerText;
                                            break;
                                        case "RUTCiaTransp":
                                            objDTE.Transporte.Aduana.RUTCiaTransp = childNodeAduana.InnerText;
                                            break;
                                        case "NomCiaTransp":
                                            objDTE.Transporte.Aduana.NomCiaTransp = childNodeAduana.InnerText;
                                            break;
                                        case "IdAdicTransp":
                                            objDTE.Transporte.Aduana.IdAdicTransp = childNodeAduana.InnerText;
                                            break;
                                        case "Booking":
                                            objDTE.Transporte.Aduana.Booking = childNodeAduana.InnerText;
                                            break;
                                        case "Operador":
                                            objDTE.Transporte.Aduana.Operador = childNodeAduana.InnerText;
                                            break;
                                        case "CodPtoEmbarque":
                                            objDTE.Transporte.Aduana.CodPtoEmbarque = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "IdAdicPtoEmb":
                                            objDTE.Transporte.Aduana.IdAdicPtoEmb = childNodeAduana.InnerText;
                                            break;
                                        case "CodPtoDesemb":
                                            objDTE.Transporte.Aduana.CodPtoDesemb = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "IdAdicPtoDesemb":
                                            objDTE.Transporte.Aduana.IdAdicPtoDesemb = childNodeAduana.InnerText;
                                            break;
                                        case "Tara":
                                            objDTE.Transporte.Aduana.Tara = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "CodUnidMedTara":
                                            objDTE.Transporte.Aduana.CodUnidMedTara = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "PesoBruto":
                                            objDTE.Transporte.Aduana.PesoBruto = Double.Parse(childNodeAduana.InnerText.Replace(".", ","));
                                            break;
                                        case "CodUnidPesoBruto":
                                            objDTE.Transporte.Aduana.CodUnidPesoBruto = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "PesoNeto":
                                            objDTE.Transporte.Aduana.PesoNeto = Double.Parse(childNodeAduana.InnerText.Replace(".", ","));
                                            break;
                                        case "CodUnidPesoNeto":
                                            objDTE.Transporte.Aduana.CodUnidPesoNeto = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "TotItems":
                                            objDTE.Transporte.Aduana.TotItems = Int64.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "TotBultos":
                                            objDTE.Transporte.Aduana.TotBultos = Int64.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "TipoBultos":
                                            TipoBultos tip = new TipoBultos();
                                            foreach (XmlNode child in childNodeAduana)
                                            {                                                
                                                if (child.Name.Equals("CodTpoBultos")) { tip.CodTpoBultos = Int32.Parse(child.InnerText); }
                                                else if (child.Name.Equals("CantBultos")) { tip.CantBultos = Int64.Parse(child.InnerText); }
                                                else if (child.Name.Equals("Marcas")) { tip.Marcas = child.InnerText; }
                                                else if (child.Name.Equals("IdContainer")) { tip.IdContainer = child.InnerText; }
                                                else if (child.Name.Equals("Sello")) { tip.Sello = child.InnerText; }
                                                else if (child.Name.Equals("EmisorSello")) { tip.EmisorSello = child.InnerText; }                                                
                                            }
                                            objDTE.Transporte.Aduana.TipoBultos.Add(tip);                                            
                                            break;
                                        case "MntFlete":
                                            objDTE.Transporte.Aduana.MntFlete = Double.Parse(childNodeAduana.InnerText.Replace(".", ","));
                                            break;
                                        case "MntSeguro":
                                            objDTE.Transporte.Aduana.MntSeguro = Double.Parse(childNodeAduana.InnerText.Replace(".", ","));
                                            break;
                                        case "CodPaisRecep":
                                            objDTE.Transporte.Aduana.CodPaisRecep = childNodeAduana.InnerText;
                                            break;
                                        case "CodPaisDestin":
                                            objDTE.Transporte.Aduana.CodPaisDestin = childNodeAduana.InnerText;
                                            break;
                                    }
                                }
                                break;
                        }
                    }
                }

                #endregion

                // NODO TOTALES
                #region TOTALES
                                
                String PathTotales = "sii:DTE/sii:Documento/sii:Encabezado/sii:Totales";
                XmlNode Totales = xmlDoc.SelectSingleNode(PathTotales, ns);
                
                foreach (XmlNode childNode in Totales)
                {
                    switch (childNode.Name)
                    {
                        case "MntNeto":
                            objDTE.Totales.MntNeto = Int64.Parse(childNode.InnerText);
                            break;
                        case "MntExe":
                            objDTE.Totales.MntExe = Int64.Parse(childNode.InnerText);
                            break;
                        case "MntBase":
                            objDTE.Totales.MntBase = Int64.Parse(childNode.InnerText);
                            break;
                        case "MntMargenCom":
                            objDTE.Totales.MntMargenCom = Int64.Parse(childNode.InnerText);
                            break;
                        case "TasaIVA":
                            objDTE.Totales.TasaIVA = Double.Parse(childNode.InnerText.Replace(".", ","));
                            break;
                        case "IVA":
                            objDTE.Totales.IVA = Int64.Parse(childNode.InnerText);
                            break;
                        case "IVAProp":
                            objDTE.Totales.IVAProp = Int64.Parse(childNode.InnerText);
                            break;
                        case "IVATerc":
                            objDTE.Totales.IVATerc = Int64.Parse(childNode.InnerText);
                            break;
                        case "ImptoReten":
                            ImptoReten imp = new ImptoReten();
                            foreach (XmlNode child in childNode.ChildNodes)
                            {                                
                                if (child.Name.Equals("TipoImp")) { imp.TipoImp = child.InnerText; }
                                else if (child.Name.Equals("TasaImp")) { imp.TasaImp = Double.Parse(child.InnerText.Replace(".", ",")); }
                                else if (child.Name.Equals("MontoImp")) { imp.MontoImp = Int64.Parse(child.InnerText); }                                
                            }
                            objDTE.Totales.ImptoReten.Add(imp);
                            break;
                        case "IVANoRet":
                            objDTE.Totales.IVANoRet = Int64.Parse(childNode.InnerText);
                            break;
                        case "CredEC":
                            objDTE.Totales.CredEC = Int64.Parse(childNode.InnerText);
                            break;
                        case "GrntDep":
                            objDTE.Totales.GrntDep = Int64.Parse(childNode.InnerText);
                            break;
                        case "Comisiones":
                            foreach (XmlNode child in childNode.ChildNodes)
                            {
                                if (child.Name.Equals("ValComNeto")) { objDTE.Totales.ComisionesTotal.ValComNeto = Int64.Parse(child.InnerText); }
                                else if (child.Name.Equals("ValComExe")) { objDTE.Totales.ComisionesTotal.ValComExe = Int64.Parse(child.InnerText); }
                                else if (child.Name.Equals("ValComIVA")) { objDTE.Totales.ComisionesTotal.ValComIVA = Int64.Parse(child.InnerText); }                                
                            }
                            break;
                        case "MntTotal":
                            objDTE.Totales.MntTotal = Int64.Parse(childNode.InnerText);
                            break;
                        case "MontoNF":
                            objDTE.Totales.MontoNF = Int64.Parse(childNode.InnerText);
                            break;
                        case "MontoPeriodo":
                            objDTE.Totales.MontoPeriodo = Int64.Parse(childNode.InnerText);
                            break;
                        case "SaldoAnterior":
                            objDTE.Totales.SaldoAnterior = Int64.Parse(childNode.InnerText);
                            break;
                        case "VlrPagar":
                            objDTE.Totales.VlrPagar = Int64.Parse(childNode.InnerText);
                            break;
                    }
                }

                #endregion

                // NODO OTRA MONEDA
                #region OTRAMONEDA

                String PathOtraMoneda = "sii:DTE/sii:Documento/sii:Encabezado/sii:OtraMoneda";
                XmlNode OtraMoneda = xmlDoc.SelectSingleNode(PathOtraMoneda, ns);
                
                if (OtraMoneda != null)
                {
                    foreach (XmlNode childNode in OtraMoneda)
                    {
                        switch (childNode.Name)
                        {
                            case "TpoMoneda":
                                objDTE.OtraMoneda.TpoMoneda = childNode.InnerText;
                                break;
                            case "TpoCambio":
                                objDTE.OtraMoneda.TpoCambio = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                            case "MntNetoOtrMnda":
                                objDTE.OtraMoneda.MntNetoOtrMnda = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                            case "MntExeOtrMnda":
                                objDTE.OtraMoneda.MntExeOtrMnda = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                            case "MntFaeCarneOtrMnda":
                                objDTE.OtraMoneda.MntFaeCarneOtrMnda = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                            case "MntMargComOtrMnda":
                                objDTE.OtraMoneda.MntMargComOtrMnda = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                            case "IVAOtrMnda":
                                objDTE.OtraMoneda.IVAOtrMnda = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                            case "ImpRetOtrMnda":
                                ImpRetOtrMnda imp = new ImpRetOtrMnda();
                                foreach (XmlNode child in childNode.ChildNodes)
                                {                                    
                                    if (child.Name.Equals("TipoImpOtrMnda")) { imp.TipoImpOtrMnda = child.InnerText; }
                                    else if (child.Name.Equals("TasaImpOtrMnda")) { imp.TasaImpOtrMnda = Double.Parse(child.InnerText.Replace(".", ",")); }
                                    else if (child.Name.Equals("VlrImpOtrMnda")) { imp.VlrImpOtrMnda = Int64.Parse(child.InnerText); }                                    
                                }
                                objDTE.OtraMoneda.ImpRetOtrMnda.Add(imp);
                                break;
                            case "IVANoRetOtrMnda":
                                objDTE.OtraMoneda.IVANoRetOtrMnda = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                            case "MntTotOtrMnda":
                                objDTE.OtraMoneda.MntTotOtrMnda = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                        }
                    }
                }

                #endregion

                // NODO DETALLE
                #region DETALLE

                String PathDetalle = "sii:DTE/sii:Documento/sii:Detalle";
                XmlNodeList Detalle = xmlDoc.SelectNodes(PathDetalle, ns);

                foreach (XmlNode childNode in Detalle)
                {
                    Detalle objDetalle = new Detalle();
                    foreach (XmlNode child in childNode)
                    {
                        switch (child.Name)
                        {
                            case "NroLinDet":
                                objDetalle.NroLinDet = Int32.Parse(child.InnerText);
                                break;
                            case "CdgItem":
                                CdgItem cdg = new CdgItem();
                                foreach (XmlNode child2 in child.ChildNodes)
                                {                                    
                                    if (child2.Name.Equals("TpoCodigo")) { cdg.TpoCodigo = child2.InnerText; }
                                    else if (child2.Name.Equals("VlrCodigo")) { cdg.VlrCodigo = child2.InnerText; }  
                                }
                                objDetalle.CdgItem.Add(cdg);
                                break;
                            case "TpoDocLiq":
                                objDetalle.TpoDocLiq = child.InnerText;
                                break;
                            case "IndExe":
                                objDetalle.IndExe = Int32.Parse(child.InnerText);
                                break;
                            case "Retenedor":
                                foreach (XmlNode child2 in child.ChildNodes)
                                {                                    
                                    if (child2.Name.Equals("IndAgente")) {  objDetalle.IndAgente = child2.InnerText; }
                                    else if (child2.Name.Equals("MntBaseFaena")) { objDetalle.MntBaseFaenaRet = Int64.Parse(child2.InnerText); }
                                    else if (child2.Name.Equals("MntMargComer")) { objDetalle.MntMargComer = Int64.Parse(child2.InnerText); }
                                    else if (child2.Name.Equals("PrcConsFinal")) { objDetalle.PrcConsFinal = Int64.Parse(child2.InnerText); }
                                }
                                break;
                            case "NmbItem":
                                objDetalle.NmbItem = child.InnerText;
                                break;
                            case "DscItem":
                                objDetalle.DscItem = child.InnerText;
                                break;
                            case "QtyRef":
                                objDetalle.QtyRef = Double.Parse(child.InnerText.Replace(".", ","));
                                break;
                            case "UnmdRe":
                                objDetalle.UnmdRe = child.InnerText;
                                break;
                            case "PrcRef":
                                objDetalle.PrcRef = Double.Parse(child.InnerText.Replace(".", ","));
                                break;
                            case "QtyItem":
                                objDetalle.QtyItem = Double.Parse(child.InnerText.Replace(".", ","));
                                break;
                            case "Subcantidad":
                                Subcantidad subCant = new Subcantidad();
                                foreach (XmlNode child2 in child.ChildNodes)
                                {
                                    if (child2.Name.Equals("SubQty")) { subCant.SubQty = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                    else if (child2.Name.Equals("SubCod")) { subCant.SubCod = child2.InnerText; }
                                    else if (child2.Name.Equals("SubQty")) { subCant.SubQty = Double.Parse(child2.InnerText.Replace(".", ",")); }                                    
                                }
                                objDetalle.Subcantidad.Add(subCant);
                                break;
                            case "FchElabor":
                                objDetalle.FchElabor = child.InnerText;
                                break;
                            case "FchVencim":
                                objDetalle.FchVencim = child.InnerText;
                                break;
                            case "UnmdItem":
                                objDetalle.UnmdItem = child.InnerText;
                                break;
                            case "PrcItem":
                                objDetalle.PrcItem = Convert.ToDouble(child.InnerText.Replace(".", ","));
                                break;
                            case "OtrMnda":
                                OtrMnda otr = new OtrMnda();
                                foreach (XmlNode child2 in child.ChildNodes)
                                {                                    
                                    if (child2.Name.Equals("PrcOtrMon")) { otr.PrcOtrMon = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                    else if (child2.Name.Equals("Moneda")) { otr.Moneda = child2.InnerText; }
                                    else if (child2.Name.Equals("FctConv")) { otr.FctConv = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                    else if (child2.Name.Equals("DctoOtrMnda")) { otr.DctoOtrMnda = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                    else if (child2.Name.Equals("RecargoOtrMnda")) { otr.RecargoOtrMnda = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                    else if (child2.Name.Equals("MontoItemOtrMnda")) { otr.MontoItemOtrMnda = Double.Parse(child2.InnerText.Replace(".", ",")); }                                    
                                }
                                objDetalle.OtrMnda.Add(otr);
                                break;
                            case "DescuentoPct":
                                objDetalle.DescuentoPct = Convert.ToDouble(child.InnerText.Replace(".", ","));
                                break;
                            case "DescuentoMonto":
                                objDetalle.DescuentoMonto = Convert.ToInt64(child.InnerText);
                                break;
                            case "SubDscto":
                                SubDscto subDes = new SubDscto();
                                foreach (XmlNode child2 in child.ChildNodes)
                                {

                                    if (child2.Name.Equals("TipoDscto")) { subDes.TipoDscto = child2.InnerText; }
                                    else if (child2.Name.Equals("ValorDscto")) { subDes.ValorDscto = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                }
                                objDetalle.SubDscto.Add(subDes);
                                break;
                            case "RecargoPct":
                                objDetalle.RecargoPct = Convert.ToDouble(child.InnerText.Replace(".", ","));
                                break;
                            case "RecargoMonto":
                                objDetalle.RecargoMonto = Convert.ToInt64(child.InnerText.Replace(".", ","));
                                break;
                            case "SubRecargo":
                                SubRecargo subRec = new SubRecargo();
                                foreach (XmlNode child2 in child.ChildNodes)
                                {
                                    if (child2.Name.Equals("TipoRecargo")) { subRec.TipoRecargo = child2.InnerText; }
                                    else if (child2.Name.Equals("ValorRecargo")) { subRec.ValorRecargo = Double.Parse(child2.InnerText.Replace(".", ",")); }                                    
                                }
                                objDetalle.SubRecargo.Add(subRec);
                                break;
                            case "CodImpAdic":
                                CodImpAdic cod = new CodImpAdic();
                                foreach (XmlNode child2 in child.ChildNodes)
                                {
                                    if (child2.Name.Equals("#text")) { cod.sCodImpAdic = child2.InnerText; }                                    
                                }
                                objDetalle.CodImpAdic.Add(cod);
                                break;
                            case "MontoItem":
                                objDetalle.MontoItem = Convert.ToInt64(child.InnerText.Replace(".", ","));
                                break;
                        }
                    }
                    objDTE.Detalle.Add(objDetalle);
                }

                #endregion

                // NODO SUBTOTALES INFORMATIVOS
                #region SUBTOTALES INFORMATIVOS

                String PathSubtotales = "sii:DTE/sii:Documento/sii:SubTotInfo";
                XmlNodeList SubTotInfo = xmlDoc.SelectNodes(PathSubtotales, ns);

                if (SubTotInfo != null)
                {
                    foreach (XmlNode childNode in SubTotInfo)
                    {
                        SubTotInfo objSubTotInfo = new SubTotInfo();
                        foreach (XmlNode child in childNode)
                        {
                            switch (child.Name)
                            {
                                case "NroSTI":
                                    objSubTotInfo.NroSTI = Int32.Parse(child.InnerText);
                                    break;
                                case "GlosaSTI":
                                    objSubTotInfo.GlosaSTI = child.InnerText;
                                    break;
                                case "OrdenSTI":
                                    objSubTotInfo.OrdenSTI = Int32.Parse(child.InnerText);
                                    break;
                                case "SubTotNetoSTI":
                                    objSubTotInfo.SubTotNetoSTI = Double.Parse(child.InnerText.Replace(".", ","));
                                    break;
                                case "SubTotIVASTI":
                                    objSubTotInfo.SubTotIVASTI = Double.Parse(child.InnerText.Replace(".", ","));
                                    break;
                                case "SubTotAdicSTI":
                                    objSubTotInfo.SubTotAdicSTI = Double.Parse(child.InnerText.Replace(".", ","));
                                    break;
                                case "SubTotExeSTI":
                                    objSubTotInfo.SubTotExeSTI = Double.Parse(child.InnerText.Replace(".", ","));
                                    break;
                                case "ValSubtotSTI":
                                    objSubTotInfo.ValSubtotSTI = Double.Parse(child.InnerText.Replace(".", ","));
                                    break;
                                case "LineasDeta":
                                    LineasDeta linDet = new LineasDeta();
                                    foreach (XmlNode child2 in child.ChildNodes)
                                    {
                                        if (child.Name.Equals("LineasDeta")) { linDet.iLineasDeta = Int32.Parse(child2.InnerText); }
                                    }
                                    objSubTotInfo.LineasDeta.Add(linDet);
                                    break;
                            }
                        }
                        objDTE.SubTotInfo.Add(objSubTotInfo);
                    }
                }

                #endregion

                // NODO DESCUENTOS Y/O RECARGOS
                #region DESCUENTOS y/o RECARGOS

                String PathDscRcgGlobal = "sii:DTE/sii:Documento/sii:DscRcgGlobal";
                XmlNodeList DscRcgGlobal = xmlDoc.SelectNodes(PathDscRcgGlobal, ns);

                if (DscRcgGlobal != null)
                {
                    foreach (XmlNode childNode in DscRcgGlobal)
                    {
                        DscRcgGlobal objDscRcgGlobal = new DscRcgGlobal();
                        foreach (XmlNode child in childNode)
                        {
                            switch (child.Name)
                            {
                                case "NroLinDR":
                                    objDscRcgGlobal.NroLinDR = Int32.Parse(child.InnerText);
                                    break;
                                case "TpoMov":
                                    objDscRcgGlobal.TpoMov = child.InnerText;
                                    break;
                                case "GlosaDR":
                                    objDscRcgGlobal.GlosaDR = child.InnerText;
                                    break;
                                case "TpoValor":
                                    objDscRcgGlobal.TpoValor = child.InnerText;
                                    break;
                                case "ValorDR":
                                    objDscRcgGlobal.ValorDR = Double.Parse(child.InnerText);
                                    break;
                                case "ValorDROtrMnda":
                                    objDscRcgGlobal.ValorDROtrMnda = Double.Parse(child.InnerText);
                                    break;
                                case "IndExeDR":
                                    objDscRcgGlobal.IndExeDR = Int32.Parse(child.InnerText);
                                    break;
                            }
                        }
                        objDTE.DscRcgGlobal.Add(objDscRcgGlobal);
                    }
                }

                #endregion

                // NODO REFERENCIAS
                #region REFERENCIAS

                String PathReferencia = "sii:DTE/sii:Documento/sii:Referencia";
                XmlNodeList Referencia = xmlDoc.SelectNodes(PathReferencia, ns);

                if (Referencia != null)
                {
                    foreach (XmlNode childNode in Referencia)
                    {
                        Referencia objReferencia = new Referencia();
                        foreach (XmlNode child in childNode)
                        {
                            switch (child.Name)
                            {
                                case "NroLinRef":
                                    objReferencia.NroLinRef = Int32.Parse(child.InnerText);
                                    break;
                                case "TpoDocRef":
                                    objReferencia.TpoDocRef = child.InnerText;
                                    break;
                                case "IndGlobal":
                                    objReferencia.IndGlobal = Int32.Parse(child.InnerText);
                                    break;
                                case "FolioRef":
                                    objReferencia.FolioRef = child.InnerText;
                                    break;
                                case "RUTOtr":
                                    objReferencia.RUTOtr = child.InnerText;
                                    break;
                                case "FchRef":
                                    objReferencia.FchRef = child.InnerText;
                                    break;
                                case "CodRef":
                                    objReferencia.CodRef = Int32.Parse(child.InnerText);
                                    break;
                                case "RazonRef":
                                    objReferencia.RazonRef = child.InnerText;
                                    break;
                            }
                        }
                        objDTE.Referencia.Add(objReferencia);
                    }
                }

                #endregion

                // NODO COMISIONES
                #region COMISIONES

                String PathComisiones = "sii:DTE/sii:Documento/sii:Comisiones";
                XmlNodeList Comisiones = xmlDoc.SelectNodes(PathComisiones, ns);

                if (Comisiones != null)
                {
                    foreach (XmlNode childNode in Comisiones)
                    {
                        Comisiones objComisiones = new Comisiones();
                        foreach (XmlNode child in childNode)
                        {
                            switch (child.Name)
                            {
                                case "NroLinCom":
                                    objComisiones.NroLinCom = Int32.Parse(child.InnerText);
                                    break;
                                case "TipoMovim":
                                    objComisiones.TipoMovim = child.InnerText;
                                    break;
                                case "Glosa":
                                    objComisiones.Glosa = child.InnerText;
                                    break;
                                case "TasaComision":
                                    objComisiones.TasaComision = Double.Parse(child.InnerText);
                                    break;
                                case "ValComNeto":
                                    objComisiones.ValComNeto = Int64.Parse(child.InnerText);
                                    break;
                                case "ValComExe":
                                    objComisiones.ValComExe = Int64.Parse(child.InnerText);
                                    break;
                                case "ValComIVA":
                                    objComisiones.ValComIVA = Int64.Parse(child.InnerText);
                                    break;
                            }
                        }
                        objDTE.Comisiones.Add(objComisiones);
                    }
                }

                #endregion

                rslt.Success = true;
                rslt.DTE = objDTE;
                return rslt;
            }
            catch(Exception ex)
            {
                rslt.Success = false;
                rslt.Mensaje = ex.Message;
                return rslt;
            }
        }

        /// <summary>
        /// Función que valida si un DTE se ha integrado antes o no.
        /// Retorna true si se puede integrar, false si ya existe y no se puede integrar
        /// </summary>
        public static ResultMessage ValidacionDTEIntegrado(String Rut, Int32 Tipo, string Folio)
        {
            ResultMessage result = new ResultMessage();
            SAPbobsCOM.Recordset oRecordset = null;
            String Query = String.Empty;
            string Query2 = null, Query3 = null, Query4 = null;


            try
            {
                // Consultar en las tablas involucradas si un documento ya existe por RUT - TIPO - FOLIO

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        //Query = "SELECT 1 FROM ODRF T0 WHERE T0.\"LicTradNum\" = '" + Rut + "' AND T0.\"FolioPref\" = '" + Tipo + "' AND T0.\"FolioNum\" = " + Folio + "";
                        //Query += " UNION";
                        //Query2 = " SELECT 1 FROM OPDN T0 WHERE T0.\"LicTradNum\" = '" + Rut + "' AND T0.\"FolioPref\" = '" + Tipo + "' AND T0.\"FolioNum\" = " + Folio + "";
                        //Query += " UNION";
                        Query3 = " SELECT 1 FROM OPCH T0 WHERE T0.\"LicTradNum\" = '" + Rut + "' AND T0.\"FolioPref\" = '" + Tipo + "' AND T0.\"FolioNum\" = " + Folio + "";
                        //Query += " UNION";
                        //Query4 = " SELECT 1 FROM ORPC T0 WHERE T0.\"LicTradNum\" = '" + Rut + "' AND T0.\"FolioPref\" = '" + Tipo + "' AND T0.\"FolioNum\" = " + Folio + "";
                        break;
                    default:
                        //Query = "SELECT 1 FROM ODRF T0 WHERE T0.LicTradNum = '" + Rut + "' AND T0.FolioPref = '" + Tipo + "' AND T0.FolioNum = " + Folio + "";
                        //Query += " UNION ";
                        //Query2 += " SELECT 1 FROM OPDN T0 WHERE T0.LicTradNum = '" + Rut + "' AND T0.FolioPref = '" + Tipo + "' AND T0.FolioNum = " + Folio + " ";
                        //Query += " UNION";
                        Query3 = " SELECT 1 FROM OPCH T0 WHERE T0.LicTradNum = '" + Rut + "' AND T0.FolioPref = '" + Tipo + "' AND T0.FolioNum = " + Folio + "";
                        //Query += " UNION";
                        //Query4 = " SELECT 1 FROM ORPC T0 WHERE T0.LicTradNum = '" + Rut + "' AND T0.FolioPref = '" + Tipo + "' AND T0.FolioNum = " + Folio + "";
                        break;
                }                

                oRecordset = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //oRecordset.DoQuery(Query);
                //// Si no hay datos
                //if (oRecordset.EoF)
                //{                   
                    oRecordset.DoQuery(Query3);
                    if (oRecordset.EoF)
                    {
                        //oRecordset.DoQuery(Query4);
                        //if (oRecordset.EoF)
                        //{
                            result.Success = true;
                        //}
                        //else
                        //{
                        //    result.Success = false;
                        //    result.Mensaje = String.Format("El DTE Nro {0} de Tipo {1} del proveedor {2} ya se encuentra integrado", Folio, Tipo, Rut);
                        //}
                    }
                    else
                    {
                        result.Success = false;
                        result.Mensaje = String.Format("El DTE Nro {0} de Tipo {1} del proveedor {2} ya se encuentra integrado", Folio, Tipo, Rut);
                    }                    
                //}
                //// si hay datos
                //else
                //{
                //    result.Success = false;
                //    result.Mensaje = String.Format("El DTE Nro {0} de Tipo {1} del proveedor {2} ya se encuentra integrado", Folio, Tipo, Rut);
                //}

                FuncionesComunes.LiberarObjetoGenerico(oRecordset);
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
        /// Crea Choose From list para el campo Socio de negocio
        /// </summary>
        public static void CrearCFL_SocioNegocio(SAPbouiCOM.Form oForm, SAPbouiCOM.EditText oEditText)
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            SAPbouiCOM.Conditions oCons = null;

            // Blindear edittext Socio de negocio
            oForm.DataSources.UserDataSources.Add("dsSN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
            oForm.Items.Item("etSN").Specific.Databind.SetBOund(true, "", "dsSN");

            oCFLs = oForm.ChooseFromLists;
            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
            oCFLCreationParams.MultiSelection = true;
            oCFLCreationParams.ObjectType = "2";
            oCFLCreationParams.UniqueID = "CflSN";
            
            oCFL = oCFLs.Add(oCFLCreationParams);

            oCons = new SAPbouiCOM.Conditions();
            //Dar condiciones al ChooseFromList Articulos Desde
            oCons = oCFL.GetConditions();

            SAPbouiCOM.Condition oCon = oCons.Add();
            oCon.Alias = "CardType";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "S";

            oCFL.SetConditions(oCons);


            //Asignamos el ChoosefromList al campo de texto
            oEditText.ChooseFromListUID = "CflSN";
            oEditText.ChooseFromListAlias = "LicTradNum";
        }

        public static void CrearCFL_Dimensiones(SAPbouiCOM.Form oForm, SAPbouiCOM.EditText oEditText, String Dim)
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            SAPbouiCOM.Conditions oCons = null;

            oCFLs = oForm.ChooseFromLists;
            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "62";
            oCFLCreationParams.UniqueID = "CflDim" + Dim;

            oCFL = oCFLs.Add(oCFLCreationParams);

            oCons = new SAPbouiCOM.Conditions();
            //Dar condiciones al ChooseFromList Articulos Desde
            oCons = oCFL.GetConditions();

            SAPbouiCOM.Condition oCon = oCons.Add();
            oCon.Alias = "DimCode";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = Dim;

            oCFL.SetConditions(oCons);


            //Asignamos el ChoosefromList al campo de texto
            oEditText.ChooseFromListUID = "CflDim" + Dim;
            oEditText.ChooseFromListAlias = "OcrCode";
        }


        public static void CrearCFL_CuentaContable(SAPbouiCOM.Form oForm, SAPbouiCOM.EditText oEditText)
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            SAPbouiCOM.Conditions oCons = null;

            oCFLs = oForm.ChooseFromLists;
            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "1";
            oCFLCreationParams.UniqueID = "CflCC";

            oCFL = oCFLs.Add(oCFLCreationParams);

            
            oCons = new SAPbouiCOM.Conditions();
            //Dar condiciones al ChooseFromList Articulos Desde
            oCons = oCFL.GetConditions();

            SAPbouiCOM.Condition oCon = oCons.Add();
            oCon.Alias = "Levels";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "5";

            oCFL.SetConditions(oCons);
            

            //Asignamos el ChoosefromList al campo de texto
            oEditText.ChooseFromListUID = "CflCC";
            oEditText.ChooseFromListAlias = "AcctCode";
        }

        /// <summary>
        /// Crea Choose From list para el campo Socio de negocio
        /// </summary>
        public static void CrearCFL_SocioNegocioNoRecibidos(SAPbouiCOM.Form oForm, SAPbouiCOM.EditText oEditText)
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            SAPbouiCOM.Conditions oCons = null;

            // Blindear edittext Socio de negocio
            oForm.DataSources.UserDataSources.Add("dsSN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
            oForm.Items.Item("etSN").Specific.Databind.SetBOund(true, "", "dsSN");

            oCFLs = oForm.ChooseFromLists;
            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "2";
            oCFLCreationParams.UniqueID = "CflSN";

            oCFL = oCFLs.Add(oCFLCreationParams);

            oCons = new SAPbouiCOM.Conditions();
            //Dar condiciones al ChooseFromList Articulos Desde
            oCons = oCFL.GetConditions();

            SAPbouiCOM.Condition oCon = oCons.Add();
            oCon.Alias = "CardType";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "S";

            oCFL.SetConditions(oCons);


            //Asignamos el ChoosefromList al campo de texto
            oEditText.ChooseFromListUID = "CflSN";
            oEditText.ChooseFromListAlias = "LicTradNum";
        }

        /// <summary>
        /// Función que valida la existencia del documento de referencia en SAP, devuelve true si existe
        /// </summary>
        public static Boolean ExisteReferencia(String Tipo, String Folio, String Cardcode, ref String TotalOC, String TipoOC)
        {
            Boolean existe = false;
            SAPbobsCOM.Recordset oRecordset = null;
            Int64 lFolio = 0;
            String Query = String.Empty;

            try
            {
                Int64.TryParse(Folio, out lFolio);
                if (Tipo.Equals("801"))
                {
                    switch (Conexion_SBO.m_oCompany.DbServerType)
                    {
                        case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                            //Query = "SELECT SUM(\"DocTotal\") FROM OPOR WHERE \"LicTradNum\" = '" + Cardcode + "' AND \"FolioNum\" IN (" + lFolio + ") ";
                            Query = "SELECT SUM(\"DocTotal\") FROM OPOR WHERE \"CardCode\" = '" + Cardcode + "' AND \"DocNum\" IN (" + lFolio + ") ";
                            break;
                        default:
                            //Query = "SELECT SUM(DocTotal) FROM OPOR WHERE LicTradNum = '" + Cardcode + "' AND FolioNum IN (" + lFolio + ") ";
                            Query = "SELECT SUM(DocTotal) FROM OPOR WHERE CardCode = '" + Cardcode + "' AND DocNum IN (" + lFolio + ") ";
                            break;
                    }

                }
                else
                {
                    switch (Conexion_SBO.m_oCompany.DbServerType)
                    {
                        case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                            Query = "SELECT SUM(\"DocTotal\") FROM OPCH WHERE \"CardCode\" = '" + Cardcode + "' AND \"FolioNum\" = '" + lFolio + "' AND \"FolioPref\" = '" + Tipo + "'";
                            break;
                        default:
                            Query = "SELECT SUM(DocTotal) FROM OPCH WHERE CardCode = '" + Cardcode + "' AND FolioNum = '" + lFolio + "' AND FolioPref = '" + Tipo + "'";
                            break;
                    }
                }
                //if (Tipo.Equals("801"))
                //{
                //    switch (Conexion_SBO.m_oCompany.DbServerType)
                //    {
                //        case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                //            Query = "SELECT SUM(\"DocTotal\") FROM OPOR WHERE \"CardCode\" = '" + Cardcode + "' AND \"DocNum\" IN (" + Folio + ") ";
                //            break;
                //        default:
                //            Query = "SELECT SUM(DocTotal) FROM OPOR WHERE CardCode = '" + Cardcode + "' AND DocNum IN (" + Folio + ") ";
                //            break;
                //    }

                //}
                //else
                //{
                //    switch (Conexion_SBO.m_oCompany.DbServerType)
                //    {
                //        case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                //            Query = "SELECT SUM(\"DocTotal\") FROM OPCH WHERE \"CardCode\" = '" + Cardcode + "' AND \"FolioNum\" = '" + lFolio + "' AND \"FolioPref\" = '" + Tipo + "'";
                //            break;
                //        default:
                //            Query = "SELECT SUM(DocTotal) FROM OPCH WHERE CardCode = '" + Cardcode + "' AND FolioNum = '" + lFolio + "' AND FolioPref = '" + Tipo + "'";
                //            break;
                //    }                    
                //}


                oRecordset = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(Query);

                TotalOC = oRecordset.Fields.Item(0).Value.ToString();

                // Si no hay datos
                if (oRecordset.EoF)
                {
                    existe = false;
                }
                // si hay datos
                else
                {
                    existe = true;
                }
                return existe;
            }
            catch (Exception)
            {

                return false;
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oRecordset);
            }
            
        }


        public static Boolean ExisteEntradaMercancia(String Tipo, String Folio, String CardCode, ref String TotalEM, ref String Folios)
        {
            Boolean existe = false;
            SAPbobsCOM.Recordset oRecordset = null;
            String Query = String.Empty;
            try
            {
                if (Tipo.Equals("801"))
                {
                    switch (Conexion_SBO.m_oCompany.DbServerType)
                    {
                        case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                            Query = " SELECT SUM(T3.\"DocTotal\"), STRING_AGG(T3.\"FolioNum\", ', ') ";
                            Query += " FROM OPOR T0 ";
                            Query += " INNER JOIN POR1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" ";
                            Query += " INNER JOIN PDN1 T2 ON T1.\"DocEntry\" = T2.\"BaseEntry\" AND T1.\"ObjType\" = T2.\"BaseType\" AND T1.\"LineNum\" = T2.\"BaseLine\" ";
                            Query += " INNER JOIN OPDN T3 ON T2.\"DocEntry\" = T3.\"DocEntry\" ";
                            Query += " INNER JOIN OCRD T4 ON T0.\"CardCode\" = T4.\"CardCode\" ";
                            Query += " WHERE T4.\"CardCode\" = '" + CardCode + "' AND T0.\"DocNum\" IN (" + Folio + ") AND T2.\"VisOrder\" = 0 ";
                            Query += " AND T3.\"DocStatus\" = 'O' ";// AND IFNULL(T3.\"FolioNum\", 0) <> 0";
                            break;
                        default:
                            Query = " SELECT SUM(T3.DocTotal), STRING_AGG(T3.FolioNum, ', ') ";
                            Query += " FROM OPOR T0 ";
                            Query += " INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry ";
                            Query += " INNER JOIN PDN1 T2 ON T1.DocEntry = T2.BaseEntry AND T1.ObjType = T2.BaseType AND T1.LineNum = T2.BaseLine ";
                            Query += " INNER JOIN OPDN T3 ON T2.DocEntry = T3.DocEntry ";
                            Query += " INNER JOIN OCRD T4 ON T0.CardCode = T4.CardCode ";
                            Query += " WHERE T4.CardCode = '" + CardCode + "' AND T0.DocNum IN (" + Folio + ") AND T2.VisOrder = 0 ";
                            Query += " AND T3.DocStatus = 'O' ";// AND IFNULL(T3.FolioNum, 0) <> 0";
                            break;
                    }




                }
                //else
                //{
                //    switch (Conexion_SBO.m_oCompany.DbServerType)
                //    {
                //        case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                //            //Query = "SELECT SUM(\"DocTotal\"), STRING_AGG(T3.\"FolioNum\", ', ') FROM OPDN WHERE \"CardCode\" = '" + CardCode + "' AND \"FolioNum\" IN (" + Folio + ") ";
                //            Query = "SELECT SUM(\"DocTotal\"), STRING_AGG(T3.\"FolioNum\", ', ') FROM OPDN WHERE \"CardCode\" = '" + CardCode + "' AND \"DocNum\" IN (" + Folio + ") ";
                //            break;
                //        default:
                //            //Query = "SELECT SUM(DocTotal), STRING_AGG(T3.FolioNum, ', ') FROM OPDN WHERE CardCode = '" + CardCode + "' AND FolioNum IN (" + Folio + ") ";
                //            Query = "SELECT SUM(DocTotal), STRING_AGG(T3.FolioNum, ', ') FROM OPDN WHERE CardCode = '" + CardCode + "' AND DocNum IN (" + Folio + ") ";
                //            break;
                //    }

                    
                //}

                oRecordset = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(Query);

                TotalEM = oRecordset.Fields.Item(0).Value.ToString();
                Folios = oRecordset.Fields.Item(1).Value.ToString();

                // Si no hay datos
                if (TotalEM.Equals("0"))
                {
                    existe = false;
                }
                // si hay datos
                else
                {
                    existe = true;
                }
                return existe;
            }
            catch (Exception)
            {

                return false;
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oRecordset);
            }
        }


        /// <summary>
        /// Función que valida la existencia del documento de referencia en SAP, devuelve true si existe
        /// </summary>
        public static Boolean ExisteReferenciaServicio(String Folio, String Cardcode, ref String DocEntryBase)
        {
            Boolean existe = false;
            SAPbobsCOM.Recordset oRecordset = null;
            String Query = String.Empty;

            switch (Conexion_SBO.m_oCompany.DbServerType)
            {
                case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                    Query = "SELECT \"DocEntry\" FROM OPOR WHERE \"CardCode\" = '" + Cardcode + "' AND \"NumAtCard\" = '" + Folio + "'";
                    break;
                default:
                    Query = "SELECT DocEntry FROM OPOR WHERE CardCode = '" + Cardcode + "' AND NumAtCard = '" + Folio + "'";
                    break;
            }            

            oRecordset = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery(Query);

            // Si no hay datos
            if (oRecordset.EoF)
            {
                existe = false;
            }
            // si hay datos
            else
            {
                existe = true;
                DocEntryBase = oRecordset.Fields.Item(0).Value.ToString();
            }

            FuncionesComunes.LiberarObjetoGenerico(oRecordset);
            return existe;
        }

        /// <summary>
        /// Función que valida la existencia del itemcode en el catalogo de SN en SAP, devuelve true si existe
        /// </summary>
        public static Boolean ExisteSubtituto(String CardCode, String Subtituto, ref String CodigoItem)
        {
            Boolean existe = false;
            SAPbobsCOM.Recordset oRecordset = null;
            String Query = String.Empty;

            switch (Conexion_SBO.m_oCompany.DbServerType)
            {
                case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                    Query = "SELECT \"ItemCode\" FROM OSCN WHERE \"CardCode\" = '" + CardCode + "' AND \"Substitute\" = '" + Subtituto + "'";
                    break;
                default:
                    Query = "SELECT ItemCode FROM OSCN WHERE CardCode = '" + CardCode + "' AND Substitute = '" + Subtituto + "'";
                    break;
            }            
            
            oRecordset = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery(Query);

            // Si no hay datos
            if (oRecordset.EoF)
            {
                existe = false;
            }
            // si hay datos
            else
            {
                CodigoItem = oRecordset.Fields.Item(0).Value;
                existe = true;
            }

            FuncionesComunes.LiberarObjetoGenerico(oRecordset);
            return existe;
        }

        /// <summary>
        /// Valida igualdad de lineas de detalle entre documentos
        /// </summary>
        public static Boolean ValidarCantidadDetalles(String Tipo, String Folio, String Cardcode, Int32 CantidadDetalle)
        {
            Boolean Iguales = false;
            Int64 lFolio = 0;
            Int64.TryParse(Folio, out lFolio);
            SAPbobsCOM.Recordset oRecordset = null;
            String Query = String.Empty;
            Int32 CantidadDetalleBase = 0;
            // Guia - OPDN
            if (Tipo.Equals("52"))
            {
                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        Query = "SELECT Count(T1.\"LineNum\") FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" WHERE T0.\"CardCode\" = '" + Cardcode + "' AND T0.\"FolioPref\" = '" + Tipo + "' AND T0.\"FolioNum\" = '" + lFolio + "'";
                        break;
                    default:
                        Query = "SELECT Count(T1.LineNum) FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry WHERE T0.CardCode = '" + Cardcode + "' AND T0.FolioPref = '" + Tipo + "' AND T0.FolioNum = '" + lFolio + "'";
                        break;
                }
                
            }
            // OC - OPOR
            else
            {
                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        Query = "SELECT Count(T1.\"LineNum\") FROM OPOR T0 INNER JOIN POR1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" WHERE T0.\"CardCode\" = '" + Cardcode + "' AND T0.\"FolioPref\" = '" + Tipo + "' AND T0.\"FolioNum\" = '" + lFolio + "'";
                        break;
                    default:
                        Query = "SELECT Count(T1.LineNum) FROM OPOR T0 INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry WHERE T0.CardCode = '" + Cardcode + "' AND T0.FolioPref = '" + Tipo + "' AND T0.FolioNum = '" + lFolio + "'";
                        break;
                }
                
            }

            oRecordset = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery(Query);

            // Si no hay datos
            if (oRecordset.EoF)
            {
                Iguales = false;
            }
            // si hay datos
            else
            {
                CantidadDetalleBase = Int32.Parse(oRecordset.Fields.Item(0).Value.ToString());
                if (CantidadDetalleBase.Equals(CantidadDetalle))
                {
                    Iguales = true;
                }
                else
                {
                    Iguales = false;
                }
            }

            return Iguales;
        }

        /// <summary>
        /// Recorre la matrix de documentos y retorna la cantidad de documentos con check en true
        /// </summary>
        public static Int32 CantidadDocumentosSeleccionados(SAPbouiCOM.Form oForm)
        {
            Int32 Contador = 0;
            SAPbouiCOM.Matrix oMatrix = null;
            Boolean check = false;

            try
            {
                oMatrix = oForm.Items.Item("oMtx").Specific;

                for (Int32 index = 1; index <= oMatrix.RowCount; index++)
                {
                    check = ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("co_Chk").Cells.Item(index).Specific).Checked;
                    if (check)
                    {
                        Contador++;
                    }
                }
            }
            catch
            {
                Contador = 0;
            }

            return Contador;
        }

        /// <summary>
        /// Valida las fechas ingresadas solo si ambas tienen datos
        /// </summary>
        public static void ValidarFechas(String DesdeFecha, String HastaFecha, SAPbouiCOM.Form oForm)
        {
            DateTime dtDesde = new DateTime(Int32.Parse(DesdeFecha.Substring(0, 4)), Int32.Parse(DesdeFecha.Substring(4, 2)), Int32.Parse(DesdeFecha.Substring(6, 2)));
            DateTime dtHasta = new DateTime(Int32.Parse(HastaFecha.Substring(0, 4)), Int32.Parse(HastaFecha.Substring(4, 2)), Int32.Parse(HastaFecha.Substring(6, 2)));

            // Validación HastaFecha > DesdeFecha
            if (dtHasta >= dtDesde)
            {
                // Hasta Fecha mayor, validación de intervalo no superior a 90 días
                TimeSpan ts = dtHasta - dtDesde;

                if (ts.Days > 90)
                {
                    oForm.Items.Item("etDesde").Specific.Value = String.Empty;
                    Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Intervalo de días no puede ser superior a 90 días", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }
            // Fechas erroneas
            else
            {
                oForm.Items.Item("etDesde").Specific.Value = String.Empty;
                Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Hasta Fecha no puede ser menor a Desde Fecha", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        /// <summary>
        /// Envia la respuesta comercial al proveedor
        /// </summary>
        public static ResultMessage EnviarRespuestaComercial(String FebosID, String TipoAccion, String Motivo, String Recinto)
        {
            ResultMessage result = new ResultMessage();
            result.Success = false;

            var client = new RestClient();
            var request = new RestRequest(String.Format(ConfigurationManager.AppSettings["Intercambio"], FebosID), Method.PUT);

            request.RequestFormat = DataFormat.Json;
            request.AddHeader("token", FuncionesComunes.ObtenerToken());
            request.AddHeader("empresa", FuncionesComunes.ObtenerRut());
            request.AddHeader("febosId", FebosID);
            request.AddHeader("motivo", Motivo);
            request.AddHeader("recinto", Recinto);
            request.AddHeader("tipoAccion", TipoAccion);
            request.AddHeader("vincularAlSii", "si");

            IRestResponse response = client.Execute(request);

            if (response.StatusDescription.Equals("OK"))
            {
                if (TipoAccion.Equals("ACD"))
                {
                    result = EnviarRespuestaComercial(FebosID, "ERM", String.Empty, ObtenerRecinto());
                }
                else
                {
                    result.Success = true;
                    result.Mensaje = "Proceso de intercambio ejecutado de forma correcta";
                }
            }
            else
            {
                result.Success = false;
                result.Mensaje = "Error al procesar documento";
            }
            result.Success = true;
            return result;
        }

        public static String ObtenerPago(String Pago)
        {
            if (Pago.Equals("-1") || Pago.Equals("0"))
            {
                Pago = "No Definido";
            }
            else if (Pago.Equals("1"))
            {
                Pago = "Contado";
            }
            else if (Pago.Equals("2"))
            {
                Pago = "Crédito";
            }
            else if (Pago.Equals("3"))
            {
                Pago = "Gratis";
            }
            else
            {
                Pago = "No Definido";
            }

            return Pago;
        }

        public static String ObtenerPlazo(String Plazo)
        {
            Int32 iPlazo = 0;
            Int32.TryParse(Plazo, out iPlazo);
            if (!iPlazo.Equals(0) && iPlazo > 0)
            {
                iPlazo = 8 - iPlazo;
                Plazo = iPlazo.ToString();
            }
            return Plazo;
        }

        public static String ObtenerTipoDocumento(String Tipo)
        {
            if (Tipo.Equals("33"))
            {
                Tipo = "Factura Electrónica";
            }
            else if (Tipo.Equals("34"))
            {
                Tipo = "Factura Exenta Electrónica";
            }
            else if (Tipo.Equals("52"))
            {
                Tipo = "Guía Despacho Electrónica";
            }
            else if (Tipo.Equals("56"))
            {
                Tipo = "Nota Débito Electrónica";
            }
            else if (Tipo.Equals("61"))
            {
                Tipo = "Nota Crédito Electrónica";
            }

            return Tipo;
        }

        public static String ObtenerTipoDocumentoNumero(String Tipo)
        {
            if (Tipo.Equals("Factura Electrónica"))
            {
                Tipo = "33";
            }
            else if (Tipo.Equals("Factura Exenta Electrónica"))
            {
                Tipo = "34";
            }
            else if (Tipo.Equals("Guía Despacho Electrónica"))
            {
                Tipo = "52";
            }
            else if (Tipo.Equals("Nota Débito Electrónica"))
            {
                Tipo = "56";
            }
            else if (Tipo.Equals("Nota Crédito Electrónica"))
            {
                Tipo = "61";
            }

            return Tipo;
        }

        public static String ObtenerCuentaDetalleOC(String DocEntry, Int32 LineNum)
        {
            SAPbobsCOM.Recordset oRecordset = null;
            String Query = String.Empty;
            String Cuenta = String.Empty;

            switch (Conexion_SBO.m_oCompany.DbServerType)
            {
                case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                    Query = "SELECT \"AcctCode\" FROM POR1 WHERE \"DocEntry\" = '" + DocEntry + "' AND \"LineNum\" = " + LineNum + "";
                    break;
                default:
                    Query = "SELECT AcctCode FROM POR1 WHERE DocEntry = '" + DocEntry + "' AND LineNum = " + LineNum + "";
                    break;
            }            

            oRecordset = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery(Query);

            if (!oRecordset.EoF)
            {
                Cuenta = oRecordset.Fields.Item(0).Value;
            }

            FuncionesComunes.LiberarObjetoGenerico(oRecordset);
            return Cuenta;
        }

        /// <summary>
        /// Cambia el campo EsContado en facturas que son emitidas como contado cuando no lo son, todo esto cuando llega la nota de credito de la factura
        /// </summary>
        public static void CambiarStatusFacturaEstadoDePagoErroneo(Referencia Ref, String CardCode)
        {
            SAPbobsCOM.Recordset oRec = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            String Query = null;
            switch (Conexion_SBO.m_oCompany.DbServerType)
            {
                case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                    Query = "SELECT \"U_SEI_CONTADO\" FROM ORPC WHERE \"FolioNum\" = " + Ref.FolioRef + " AND \"CardCode\" = " + CardCode + "";
                    break;
                default:
                    Query = "SELECT U_SEI_CONTADO FROM ORPC WHERE FolioNum = " + Ref.FolioRef + " AND CardCode = " + CardCode + "";
                    break;
            }

            String EsContado = String.Empty;

            oRec.DoQuery(Query);
            if (!oRec.EoF)
            {
                EsContado = oRec.Fields.Item(0).Value;                
            }
            FuncionesComunes.LiberarObjetoGenerico(oRec);

            if (EsContado.Equals("N"))
            {
                oRec = Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                switch (Conexion_SBO.m_oCompany.DbServerType)
                {
                    case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                        Query = "UPDATE ORPC SET \"U_SEI_CONTADO\" = 'S' WHERE \"FolioNum\" = " + Ref.FolioRef + " AND \"CardCode\" = " + CardCode + "";
                        break;
                    default:
                        Query = "UPDATE ORPC SET U_SEI_CONTADO = 'S' WHERE FolioNum = " + Ref.FolioRef + " AND CardCode = " + CardCode + "";
                        break;
                }
                
                oRec.DoQuery(Query);
            }
            
        }

        /// <summary>
        /// Obtiene solo los valores numericos del total de la linea y devuelve su conversion en double
        /// </summary>
        public static Double ObtenerTotalEnNumeros(String sTotal)
        {
            Double Total = 0;
            String result = Regex.Replace(sTotal, @"[^\d]", "");
            Total = Double.Parse(result);
            return Total;
        }
        #endregion
    }
}
