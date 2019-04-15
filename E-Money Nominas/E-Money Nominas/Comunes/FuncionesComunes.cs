using E_Money_Nominas.Conexiones;
using System;
namespace E_Money_Nominas.Comunes
{
    public class FuncionesComunes
    {
        /// <summary>
        /// Valida caracteres invalidos para nombres de archivos.
        /// </summary>
        /// <param name="oString"></param>
        public static string ValidarCaracteres(string oString)
        {
            oString = oString.Replace("~", string.Empty);
            oString = oString.Replace("#", string.Empty);
            oString = oString.Replace("%", string.Empty);
            oString = oString.Replace("&", string.Empty);
            oString = oString.Replace("*", string.Empty);
            oString = oString.Replace("{", string.Empty);
            oString = oString.Replace("}", string.Empty);
            oString = oString.Replace("\\", string.Empty);
            oString = oString.Replace(":", string.Empty);
            oString = oString.Replace("<", string.Empty);
            oString = oString.Replace(">", string.Empty);
            oString = oString.Replace("?", string.Empty);
            oString = oString.Replace("/", string.Empty);
            oString = oString.Replace("+", string.Empty);
            oString = oString.Replace("|", string.Empty);
            oString = oString.Replace("'", string.Empty);
            oString = oString.Replace(" ", string.Empty);
            oString = oString.Replace("-", string.Empty);
            oString = oString.Replace(".", string.Empty);

            return oString;
        }

        /// <summary>
        /// Limpia combobox
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="ItemUID"></param>
        public static void BorrarCombo(string FormUID, string ItemUID)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.ComboBox oCombobox = null;
            try
            {
                oForm = Conexion_SBO.m_SBO_Appl.Forms.Item(FormUID);
                oCombobox = (SAPbouiCOM.ComboBox)oForm.Items.Item(ItemUID).Specific;

                for (int i = oCombobox.ValidValues.Count - 1; i >= 0; i--)
                {
                    oCombobox.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                }

            }
            catch (Exception ex)
            {
                Comunes.Msj_Appl.Errores(14, ex.Message);
            }
        }

        /// <summary>
        /// Retorna un string con la consulta de pagos existentes.
        /// </summary>
        /// <returns></returns>
        public static string QueryPagosExistentes(string Fecha)
        {
            string query = string.Empty;

            query = @"Select WizardName 
                      From OPWZ
                      Where StatusDisc='Ejecutado'
                      and PmntDate = '" + Fecha + "'";

            return query;
        }

        /// <summary>
        /// Query string para obtener pagos masivos.
        /// </summary>
        /// <returns></returns>
        public static string QueryPagosMasivos(string Pago)
        {
            string query = string.Empty;

            query = @"SELECT Left(rtrim(ltrim(T3.dflAccount))+space(20),20) 'Cuenta Destino' 
                            ,T3.BankCode 'Código Banco'
                            ,Right('00000000'+substring(ltrim(rtrim(replace(T3.lictradnum,'.',''))),1,charindex('-',ltrim(rtrim(replace(T3.lictradnum,'.',''))))-1),8) + Right('0'+ltrim(rtrim(replace(T3.lictradnum,'.',''))),1) 'Rut Benificiario'
                            ,Left(rtrim(ltrim(T0.PayeeName))+space(45),45) 'Nombre Beneficiario'
                            ,T1.SumApplied Monto
                            ,case when T1.InvType = 19 then ' '  else T2.FolioNum end   Folio
                            ,T4.ValDateTo Fecha
                            ,T2.FolioPref 'Tipo Doc'
                            ,T2.DocDueDate
                            ,T0.PymBnkCode 'Banco Local'
                            ,T6.BankName
							,T2.DocCur
							,T0.PymBnkAcct
                            ,T3.E_Mail
                            FROM  [dbo].[OPEX] T0
                            INNER JOIN ODSC T6 ON T0.PymBnkCode = T6.BankCode
                            INNER JOIN OCRD T3 ON T3.CardCode = T0.VendorNum
                            INNER JOIN [dbo].[OPWZ] T4 on   T4.IdNumber = T0.PaymWizCod,   [dbo].[VPM2] T1 
                            LEFT JOIN [dbo].[OPCH] T2 ON T2.DocEntry = T1.DocEntry   
                            WHERE T0.PaymDocNum = T1.DocNum  
                            --AND  T1.DocEntry = T2.DocEntry 
                            and T4.WizardName  = '" + Pago + @"'
                            and  charindex('-',T3.lictradnum)>0
                            --ORDER BY T3.LicTradNum";

            return query;
        }

        /// <summary>
        /// Query string para obtener pagos masivos del banco santander.
        /// El comportamiento es distinto ya que se debe agrupar por cliente y sumar los montos
        /// </summary>
        /// <returns></returns>
        public static string QueryPagosMasivosSantander(string Pago)
        {
            string query = string.Empty;

            query = @"SELECT Left(rtrim(ltrim(T3.dflAccount))+space(20),20) 'Cuenta Destino' 
                    ,T3.BankCode 'Código Banco'
                    ,Right('00000000'+substring(ltrim(rtrim(replace(T3.lictradnum,'.',''))),1,charindex('-',ltrim(rtrim(replace(T3.lictradnum,'.',''))))-1),8) + Right('0'+ltrim(rtrim(replace(T3.lictradnum,'.',''))),1) 'Rut Benificiario'
                    ,Left(rtrim(ltrim(T0.PayeeName))+space(45),45) 'Nombre Beneficiario'
                    ,SUM(T1.SumApplied) Monto
                    ,T0.PymBnkCode 'Banco Local'
                    ,T6.BankName
		                ,T2.DocCur
		                ,T0.PymBnkAcct
                    ,T3.E_Mail
                    FROM  [dbo].[OPEX] T0
                    INNER JOIN ODSC T6 ON T0.PymBnkCode = T6.BankCode
                    INNER JOIN OCRD T3 ON T3.CardCode = T0.VendorNum
                    INNER JOIN [dbo].[OPWZ] T4 on   T4.IdNumber = T0.PaymWizCod,   [dbo].[VPM2] T1 
                    LEFT JOIN [dbo].[OPCH] T2 ON T2.DocEntry = T1.DocEntry   
                    WHERE T0.PaymDocNum = T1.DocNum  
                    --AND  T1.DocEntry = T2.DocEntry 
                    and T4.WizardName  = '" + Pago + @"'
                    and  charindex('-',T3.lictradnum)>0
	                GROUP BY T3.dflAccount,T3.BankCode,T3.LicTradNum,T0.PayeeName,T0.PymBnkCode,T6.BankName,T2.DocCur,T0.PymBnkAcct,T3.E_Mail
	                ORDER BY T0.PayeeName";


            return query;
        }

        /// <summary>
        ///  Query string para obtener el banco local al cual se enviará la nomina, con esto
        ///  se sabe que query de pago masivo realizar
        /// </summary>
        /// <returns></returns>
        public static string QueryObtenerBanco(string Pago)
        {
            string query = string.Empty;

            query = @"SELECT TOP (1) O.PymBnkCode
                    FROM OPEX O INNER JOIN OPWZ Z ON Z.IdNumber = O.PaymWizCod
                    WHERE Z.WizardName  = '" + Pago + "'";

            return query;
        }
    }
}
