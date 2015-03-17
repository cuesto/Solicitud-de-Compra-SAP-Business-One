using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Solicitudes_de_Compra.modelo.logica
{
    class Copiadora
    {
        private SAPbouiCOM.Application app;
        private SAPbobsCOM.Company oCompany;
        private SAPbouiCOM.Form forma;
        private bool copiarDe;

        public Copiadora()
        {
            app = null;
            oCompany = null;
            forma = null;
            copiarDe = false;
        }

        public void copyTo(string docNum)
        {
            SAPbouiCOM.EditText oEdit;
            string query;
            SAPbobsCOM.Recordset rs;
            SAPbouiCOM.Matrix oMatrix;

            try
            {
                oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("38").Specific;

                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                
                query = "SELECT * FROM [dbo].[@ZTV_LINES]  T0 , [dbo].[@ZTV]  T1 WHERE T0.[U_LnStatus] = 'O' AND T0.[DocEntry]  =  T1.[DocEntry] AND  T1.[DocNum]  = "
                    + docNum;

                rs.DoQuery(query);
                
                oEdit = (SAPbouiCOM.EditText)forma.Items.Item("4").Specific;

                if ("" == rs.Fields.Item("U_CardCode").Value.ToString())
                {
                    oEdit.Value = "GEN";
                }
                else
                {
                    oEdit.Value = rs.Fields.Item("U_CardCode").Value.ToString();
                }

                oEdit = (SAPbouiCOM.EditText)forma.Items.Item("24").Specific;
                oEdit.Value = rs.Fields.Item("U_DiscPrc").Value.ToString();

                rs.MoveFirst();
                int i = 1;
                while (!rs.EoF)
                {
                    if ("" != rs.Fields.Item("U_ItemCode").Value.ToString())
                    {
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value = rs.Fields.Item("U_ItemCode").Value.ToString();
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("3").Cells.Item(i).Specific).Value = rs.Fields.Item("U_Dscript").Value.ToString();
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(i).Specific).Value = rs.Fields.Item("U_OpenQty").Value.ToString();
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(i).Specific).Value = rs.Fields.Item("U_PrcBfDi").Value.ToString();
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(i).Specific).Value = rs.Fields.Item("U_DiscItem").Value.ToString();
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("160").Cells.Item(i).Specific).Value = rs.Fields.Item("U_TaxCode").Value.ToString();

                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_BsEntry").Cells.Item(i).Specific).Value = rs.Fields.Item("DocEntry").Value.ToString();
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_BsLine").Cells.Item(i).Specific).Value = rs.Fields.Item("LineId").Value.ToString();
                    }
                    rs.MoveNext();
                    i++;
                }
                oEdit = (SAPbouiCOM.EditText)forma.Items.Item("16").Specific;
                oEdit.Value = "Basado en Solicitud " + docNum;
                oEdit.Active = true;
            }
            catch (Exception e)
            {
                app.StatusBar.SetText("copiadora_method_copyTo -> " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        public void setButtonCpFrom()
        {
            try
            {
                SAPbouiCOM.Item oItem, temp;
                SAPbouiCOM.Button btnCpFrm;

                temp = forma.Items.Item("10000329");
                oItem = forma.Items.Add("btnCpFrom", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = temp.Left - 105;
                oItem.Width = 100;
                oItem.Top = temp.Top;
                oItem.Height = 19;
                oItem.Enabled = false;
                btnCpFrm = (SAPbouiCOM.Button)oItem.Specific;
                btnCpFrm.Caption = "Copiar de Solicitud";
            }
            catch (Exception e)
            {
                app.StatusBar.SetText("copiadora_method_setButtonCpFrom -> " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public void app_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            // se utiliza para capturar el evento insertar
            if (pVal.FormUID == forma.UniqueID && pVal.FormMode == 3 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "1" && pVal.BeforeAction == true)
            {
                SAPbouiCOM.Matrix oMatrix;
                SAPbobsCOM.Recordset rs,rsTemp;
                string query, queryTemp;
                double OpnQnty, temp;

                try
                {
                    oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("38").Specific;
                    rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rsTemp = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    OpnQnty = 0;
                    temp = 0;
                    int cnt;

                    for (int i = 1; i < oMatrix.RowCount; i++)
                    {
                        //esta consulta me trae la linea de la solicitud por el BaseEntry y LineId
                        query = "SELECT T0.[U_ItemCode], T0.[U_OpenQty] FROM [dbo].[@ZTV_LINES]  T0 WHERE T0.[DocEntry]  = "
                            + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_BsEntry").Cells.Item(i).Specific).String + " and  T0.[LineId] = "
                            + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_BsLine").Cells.Item(i).Specific).String + " and  T0.[U_ItemCode] = '"
                            + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).String + "'";

                        rs.DoQuery(query);

                        rs.MoveFirst();

                        while (!rs.EoF)
                        {              
                            OpnQnty = Convert.ToDouble(rs.Fields.Item("U_OpenQty").Value.ToString().Trim());
                            temp = Convert.ToDouble((((SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(i).Specific).String).ToString());
                            OpnQnty = OpnQnty - temp;

                            if (OpnQnty < 0)
                            {
                                OpnQnty = 0;
                            }

                            // Ya calculado el opnquantity entonces se actualiza el OpnQnty de la linea
                            queryTemp = "Update [dbo].[@ZTV_LINES] SET [U_OpenQty] = " + OpnQnty + " WHERE [LineId] = "
                                + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_BsLine").Cells.Item(i).Specific).String
                                + " AND [DocEntry] = " + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_BsEntry").Cells.Item(i).Specific).String + " ";

                            rsTemp.DoQuery(queryTemp);

                            if(OpnQnty == 0)
                            {
                                // se actualiza el status por linea, solo si el OpnQnty es igual a 0
                                queryTemp = "Update [dbo].[@ZTV_LINES] SET [U_LnStatus] = 'C' WHERE [LineId] = "
                                + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_BsLine").Cells.Item(i).Specific).String
                                + " AND [DocEntry] = " + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_BsEntry").Cells.Item(i).Specific).String + " ";

                                rsTemp.DoQuery(queryTemp);
                            }
                            rs.MoveNext();
                        }

                        queryTemp = "SELECT count(*) AS count FROM [dbo].[@ZTV_LINES]  T0 WHERE T0.[DocEntry]  = " 
                            + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_BsEntry").Cells.Item(i).Specific).String;
                        rs.DoQuery(queryTemp);
                        cnt = Convert.ToInt32(rs.Fields.Item("count").Value);

                        queryTemp = "SELECT count(*) AS count FROM [dbo].[@ZTV_LINES]  T0 WHERE T0.[DocEntry]  = "
                            + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_BsEntry").Cells.Item(i).Specific).String
                            + "AND T0.[U_LnStatus] = 'C'";
                        rs.DoQuery(queryTemp);
                        
                        if (cnt == Convert.ToInt32(rs.Fields.Item("count").Value))
                        {
                            queryTemp = "UPDATE [dbo].[@ZTV]  SET Status = 'C' WHERE DocEntry = "
                                + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_BsEntry").Cells.Item(i).Specific).String;
                            rs.DoQuery(queryTemp);
                        }   
                    }
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("copiadora_app_event_add " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

            // se utiliza para ajustar el rectangulo cuando la forma cambie de tamanio
            if (pVal.FormUID == forma.UniqueID && pVal.BeforeAction == false && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE)
            {
                SAPbouiCOM.Item oItem, temp;
                
                try
                {
                    temp = forma.Items.Item("10000329");
                    oItem = forma.Items.Item("btnCpFrom");
                    oItem.Left = temp.Left - 105;
                    oItem.Width = 100;
                    oItem.Top = temp.Top;
                    oItem.Height = 19;
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("copiadora_app_event_resize " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

            // se invoca una matrix con datos cabeceras de las solicitudes procesadas
            if ((pVal.FormUID == forma.UniqueID) && pVal.BeforeAction == false && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED 
                && pVal.ItemUID == "btnCpFrom")
            {
                ListaSolicitud lista;
                string cardCode;

                try
                {
                    if (!copiarDe)
                    {
                        lista = new ListaSolicitud();
                        lista.setApp(ref app);
                        lista.setCompany(ref oCompany);
                        lista.setForm(forma);
                        cardCode = ((SAPbouiCOM.EditText)forma.Items.Item("4").Specific).Value;
                        if (cardCode == "GEN")
                        {
                            cardCode = "";
                        }
                        lista.setCardCode(cardCode);
                        lista.crearForma();
                        copiarDe = true;
                    }
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("copiadora_app_event_cpfrom -> " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

            // aqui se verifica que la forma aunque este en modo de agregar
            // requiera que el campo proveedor no este vacio para activar el boton 'Copiar de'
            if ((pVal.FormUID == forma.UniqueID))
            {
                try
                {
                    if(((SAPbouiCOM.EditText)forma.Items.Item("4").Specific).Value != "" && pVal.FormMode == 3)
                    {
                        forma.Items.Item("btnCpFrom").Enabled = true;
                    }
                    if(pVal.FormMode == 0 || pVal.FormMode == 1 || pVal.FormMode == 2)
                    {
                        forma.Items.Item("btnCpFrom").Enabled = false;
                    }
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("copiadora_app_event_state_cpfrom -> " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }
        }

        public void setApp(ref SAPbouiCOM.Application app)
        {
            this.app = app;
        }

        public void setCompany(ref SAPbobsCOM.Company oCompany)
        {
            this.oCompany = oCompany;
        }

        public void setForma(ref SAPbouiCOM.Form forma)
        {
            this.forma = forma;
        }

        public void capturarEventos()
        {
            app.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(app_ItemEvent);
        }
    }
}
