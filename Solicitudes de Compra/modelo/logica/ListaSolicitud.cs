using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Solicitudes_de_Compra.modelo.logica
{
    class ListaSolicitud
    {
        private SAPbobsCOM.Company oCompany;
        private SAPbouiCOM.Application app;
        private SAPbouiCOM.Form forma, oForm;
        private string cardCode, colUID, colUID1, formUID;
        private List<int> filas;
        private bool smf;
        private int row;
        private int cont;

        public ListaSolicitud()
        {
            cont = 1;
            app = null;
            oCompany = null;
            forma = null;
            oForm = null;
            filas = new List<int>();
            colUID = "";
            colUID1 = "";
            row = 0;
            formUID = "";
            smf = false;
        }

        public void crearForma()
        {
            try
            {
                forma = app.Forms.Item("FLSC");
            }
            catch
            {
                // se carga el archivo desde el xml
                string xmlCargado = Config.getConfig().cargarDesdeXML("FLista de Solicitudes.srf");
                app.LoadBatchActions(ref xmlCargado);
                forma = app.Forms.Item("FLSC");

                forma.Visible = true;

                getDataFromDataSource();

                //se guarda la forma en xml
                Config.getConfig().guardarComoXML(forma.GetAsXML(), "FLista de Solicitudes.xml");
            }
            app.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(app_ItemEvent);
        }

        public void getDataFromDataSource()
        {
            SAPbouiCOM.Matrix oMatrix;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Condition oCon = null;

            try
            {
                oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix1").Specific;
                oCons = new SAPbouiCOM.Conditions();

                oCon = oCons.Add();
                oCon.BracketOpenNum = 2;
                oCon.Alias = "U_CardCode";
                if ("" == cardCode)
                {
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_IS_NULL;
                    oCon.CondVal = cardCode;
                }
                else
                {
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = cardCode;
                }
                oCon.BracketCloseNum = 1;
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCon = oCons.Add();
                oCon.BracketOpenNum = 1;
                oCon.Alias = "Status";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "O";
                oCon.BracketCloseNum = 2;
                
                oMatrix.Clear();
                forma.DataSources.DBDataSources.Item("@ZTV").Query(oCons);
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                oMatrix.SelectRow(1, true, true);
            }
            catch (Exception e)
            {
                app.StatusBar.SetText("ListaSolicitud_method_getDataFromDataSource -> " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
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

            // este es el evento del boton Atras del asistente
            if (pVal.FormUID == "FLSC" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "5" && pVal.BeforeAction == true)
            {
                try
                {
                    forma.Items.Item("4").Enabled = true;
                    forma.Items.Item("5").Enabled = false;
       
                    forma.Items.Item("matrix1").Visible = true;
                    forma.Items.Item("matrix2").Visible = false;
                    smf = false;
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("listaSolicitud_app_event_atras " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

            // este es el evento del boton Siguiente del asistente
            if (pVal.FormUID == "FLSC" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "4" && pVal.BeforeAction == false)
            {
                string query;
                SAPbouiCOM.Matrix oMatrix, oMtxTemp;
                SAPbobsCOM.Recordset rs;

                try
                {
                    oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix1").Specific;
                    oMtxTemp = (SAPbouiCOM.Matrix)forma.Items.Item("matrix2").Specific;

                    forma.DataSources.DBDataSources.Item("@ZTV").Clear();
                    forma.DataSources.DBDataSources.Item("@ZTV_LINES").Clear();
                    oMtxTemp.Clear();

                    oMtxTemp.FlushToDataSource();

                    for (int i = oMatrix.VisualRowCount; i > 0; i--)
                    {
                        if (oMatrix.IsRowSelected(i))
                        {
                            rs = null;
                            rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            query = "SELECT T0.[DocNum], T1.[U_ItemCode], T1.[U_Dscript], T1.[U_OpenQty], T1.[U_PrcBfDi], T1.[U_DiscItem], T1.[U_Total], T1.[LineId] FROM [dbo].[@ZTV]  T0 , [dbo].[@ZTV_LINES]  T1 WHERE T1.[U_LnStatus] = 'O' AND T1.[U_OpenQty] >=1 and  T0.[DocEntry] =  T1.[DocEntry] and  T0.[DocNum]  = "
                                + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("V_0").Cells.Item(i).Specific).Value + " ORDER BY T0.[DocNum]";

                            rs.DoQuery(query);
                            rs.MoveFirst();
                            int j = 0;
                            while (!rs.EoF)
                            {
                                forma.DataSources.DBDataSources.Item("@ZTV").InsertRecord(j);
                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").InsertRecord(j);
                                forma.DataSources.DBDataSources.Item("@ZTV").SetValue("DocNum", j, (rs.Fields.Item("DocNum").Value).ToString());
                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("LineId", j, (rs.Fields.Item("LineId").Value).ToString());
                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_ItemCode", j, (rs.Fields.Item("U_ItemCode").Value).ToString());
                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_Dscript", j, (rs.Fields.Item("U_Dscript").Value).ToString());
                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_OpenQty", j, (rs.Fields.Item("U_OpenQty").Value).ToString());
                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_PrcBfDi", j, (rs.Fields.Item("U_PrcBfDi").Value).ToString());
                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_DiscItem", j, (rs.Fields.Item("U_DiscItem").Value).ToString());
                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_Total", j, (rs.Fields.Item("U_Total").Value).ToString());
                                rs.MoveNext();
                                j++;
                            }
                        }
                    }

                    int temp = forma.DataSources.DBDataSources.Item("@ZTV").Size - 1;

                    forma.DataSources.DBDataSources.Item("@ZTV").RemoveRecord(temp);
                    forma.DataSources.DBDataSources.Item("@ZTV_LINES").RemoveRecord(temp);

                    oMtxTemp.LoadFromDataSource();

                    oMtxTemp.SelectRow(1, true, true);
                    forma.Items.Item("5").Enabled = true;
                    forma.Items.Item("4").Enabled = false;
                    forma.Items.Item("matrix1").Visible = false;
                    forma.Items.Item("matrix2").Visible = true;

                    smf = true;
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("listaSolicitud_app_event_siguiente " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

            // aqui se captura el evento del boton finalizar
            if (pVal.FormUID == "FLSC" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "3" && pVal.BeforeAction == true)
            {
                SAPbouiCOM.Matrix oMatrix;
                List<int> lineas;
                
                try
                {
                    lineas = new List<int>();

                    if (!smf)
                    {
                        oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix1").Specific;

                        for (int i = oMatrix.VisualRowCount; i > 0; i--)
                        {
                            if (oMatrix.IsRowSelected(i))
                            {
                                filas.Add(Convert.ToInt32(((SAPbouiCOM.EditText)oMatrix.Columns.Item("V_0").Cells.Item(i).Specific).Value));
                            }
                        }                       
                    }
                    else
                    {
                        oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix2").Specific;

                        for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                        {
                            if (oMatrix.IsRowSelected(i))
                            {
                                filas.Add(Convert.ToInt32(((SAPbouiCOM.EditText)oMatrix.Columns.Item("V_0").Cells.Item(i).Specific).Value));
                                lineas.Add(Convert.ToInt32(((SAPbouiCOM.EditText)oMatrix.Columns.Item("V_-1").Cells.Item(i).Specific).Value));
                            }
                        }
                    }

                    if (lineas.Count > 0)
                    {
                        for (int i = 0; i < lineas.Count; i++)
                        {
                            llenarMatrixDesdePersonalizada(filas[i], lineas[i]);
                        }
                        filas.Clear();
                        lineas.Clear();
                    }

                    if(filas.Count > 0)
                    {
                        for (int i = 0; i < filas.Count; i++)
                        {
                            llenarMatrix(filas[i]);
                        }
                        filas.Clear();
                    }
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("listaSolicitud_app_event_finalizar " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }
        }

        public void llenarMatrixDesdePersonalizada(int docNum, int lineId)
        {
            string query;
            SAPbobsCOM.Recordset rs;
            SAPbouiCOM.Matrix oMatrix;

            try
            {
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                query = "SELECT T1.[U_ItemCode], T1.[U_Dscript], T1.[U_OpenQty], T1.[U_PrcBfDi], T1.[U_DiscItem], T1.[U_TaxCode], T1.[DocEntry], T1.[LineId] FROM [dbo].[@ZTV]  T0 , [dbo].[@ZTV_LINES]  T1 WHERE T0.[DocEntry]  =  T1.[DocEntry] and  T0.[DocNum]  = "
                    + docNum + " and  T1.[LineId] = " + lineId;
                rs.DoQuery(query);

                int i = oMatrix.RowCount;

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
                forma.Close();
            }
            catch (Exception e)
            {
                app.StatusBar.SetText("listaSolicitud_method_llenarMatrixDesdePersonalizada -> " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        public void llenarMatrix(int docNum)
        {
            string query;
            SAPbobsCOM.Recordset rs;
            SAPbouiCOM.Matrix oMatrix;

            try
            {
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                query = "SELECT T0.[DocNum], T1.[U_ItemCode], T1.[U_Dscript], T1.[U_OpenQty], T1.[U_PrcBfDi], T1.[DocEntry], T1.[U_DiscItem], T1.[U_TaxCode], T1.[U_Total], T1.[LineId] FROM [dbo].[@ZTV]  T0 , [dbo].[@ZTV_LINES]  T1 WHERE T1.[U_LnStatus] = 'O' AND T1.[U_OpenQty] >=1 and  T0.[DocEntry] =  T1.[DocEntry] and  T0.[DocNum]  = "
                + docNum + " ORDER BY T0.[DocNum]";

                rs.DoQuery(query);
                rs.MoveFirst();
                int i = oMatrix.RowCount;

                while (!rs.EoF)
                {
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value = rs.Fields.Item("U_ItemCode").Value.ToString();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("3").Cells.Item(i).Specific).Value = rs.Fields.Item("U_Dscript").Value.ToString();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(i).Specific).Value = rs.Fields.Item("U_OpenQty").Value.ToString();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(i).Specific).Value = rs.Fields.Item("U_PrcBfDi").Value.ToString();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(i).Specific).Value = rs.Fields.Item("U_DiscItem").Value.ToString();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("160").Cells.Item(i).Specific).Value = rs.Fields.Item("U_TaxCode").Value.ToString();

                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_BsEntry").Cells.Item(i).Specific).Value = rs.Fields.Item("DocEntry").Value.ToString();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_BsLine").Cells.Item(i).Specific).Value = rs.Fields.Item("LineId").Value.ToString();
                    rs.MoveNext();
                    i++;
                }
                forma.Close();
            }
            catch (Exception e)
            {
                app.StatusBar.SetText("listaSolicitud_method_llenarMatrix -> " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        public void setCardCode(string cardCode)
        {
            this.cardCode = cardCode;
        }

        public void setForm(SAPbouiCOM.Form oForm)
        {
            this.oForm = oForm;
        }

        public void setFormUID(string formUID)
        {
            this.formUID = formUID;
        }

        public void setApp(ref SAPbouiCOM.Application app)
        {
            this.app = app;
        }

        public void setCompany(ref SAPbobsCOM.Company oCompany)
        {
            this.oCompany = oCompany;
        }
    }
}