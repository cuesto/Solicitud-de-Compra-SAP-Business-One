using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Solicitudes_de_Compra.modelo;

namespace Solicitudes_de_Compra.modelo.logica
{
    class SolicitudCompra
    {
        private SAPbouiCOM.Application app;
        private SAPbobsCOM.Company oCompany;
        private SAPbouiCOM.Form forma;
        private string colUID;
        private bool smf, copyTo;
        private int rowNum;
        private double rateITBIS;
        public static string docNum;
        public static bool copiar;
       
        public SolicitudCompra()
        {
            app = null;
            oCompany = null;
            forma = null;
            colUID = "";
            smf = false;
            copyTo = false;
            rowNum = 0;
            rateITBIS = 0;
            copiar = false;
        }

        public void crearForma()
        {
            //se comprueba que la forma no este creada para poder
            //crear una forma
            try
            {
                forma = app.Forms.Item("FSDC");
            }
            catch
            {
                // se carga el archivo desde el xml
                string xmlCargado = Config.getConfig().cargarDesdeXML("FSolicitud de compra.srf");
                app.LoadBatchActions(ref xmlCargado);
                forma = app.Forms.Item("FSDC");

                crearChooseFromList();
                llenarComboBox();
                setFecha();
                setMatrix();
                setImpuesto();

                // se posiciona el campo proveedor como 
                // el primer campo activo de la forma
                forma.ActiveItem = "txtCodProv";
                
                // el browseby activa las flechas de navegacion
                // para navegar entre documentos ya procesados
                forma.DataBrowser.BrowseBy = "txtDcEntry";

                // se asigna el indice a la columna
                setColCount();
                forma.Visible = true;
                app.SetStatusBarMessage("Se inicializado del Add-On", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                // se guarda el elemento xml
                Config.getConfig().guardarComoXML(forma.GetAsXML(), "FSolicitud de compra.xml");
            }
            app.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(app_ItemEvent);
        }

        // se obtiene la tasa de impuesto correspondiente al ITBIS
        public void setImpuesto()
        {
            SAPbobsCOM.Recordset rs = null;
            string query;

            try
            {
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                query = "SELECT Rate FROM OSTC T0 WHERE Code = 'IT'";

                rs.DoQuery(query);
                rs.MoveFirst();

                rateITBIS = Convert.ToDouble(rs.Fields.Item("Rate").Value);
            }
            catch(Exception e)
            {
                app.StatusBar.SetText("solicitud_method_setImpuesto " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        // este metodo asigna la fecha actual a los campos fechas
        // de la forma
        public void setFecha()
        {
            try
            {
                forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_DocDate", 0, DateTime.Today.ToString("yyyyMMdd"));
                forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_TaxDate", 0, DateTime.Today.ToString("yyyyMMdd"));
                forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_DcDDate", 0, DateTime.Today.ToString("yyyyMMdd"));
            }
            catch (Exception e)
            {
                app.StatusBar.SetText("solicitud_method_setFecha " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        public void setMatrix()
        {
            SAPbouiCOM.Matrix oMatrix;
            
            try
            {
                oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix1").Specific;
                oMatrix.Clear();
                oMatrix.AddRow(1, 0);
            }
            catch (Exception e)
            {
                app.StatusBar.SetText("solicitud_method_setMatrix " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        public void crearChooseFromList()
        {
            SAPbouiCOM.ChooseFromList oCfl = null;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Condition oCon = null;

            try
            {
                // se carga el objeto con la data del cfl que se utilizara
                // este se encuentra en la forma que se cargo  del xml
                // codigo del suplidor
                oCfl = (SAPbouiCOM.ChooseFromList)forma.ChooseFromLists.Item("cflCodProv");
                oCons = oCfl.GetConditions();
                oCon = oCons.Add();

                // se toma este alias como el campo que se utilizara para 
                // filtrar la info
                oCon.Alias = "CardType";

                // en este caso la info del campo cardtype debe ser igual para
                // que se muestre la data
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;

                // el valor de la condicion debe se S que significa del tipo suplidor
                oCon.CondVal = "S";

                // se asignan las condiciones al cfl
                oCfl.SetConditions(oCons);

                //nombre del suplidor
                oCfl = (SAPbouiCOM.ChooseFromList)forma.ChooseFromLists.Item("cflNmbProv");
                oCons = oCfl.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "CardType";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "S";
                oCfl.SetConditions(oCons);
            }
            catch (Exception e)
            {
                app.StatusBar.SetText("solicitud_method_crearChooseFromList " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        public void setColCount()
        {
            SAPbouiCOM.Matrix oMatrix;
            SAPbouiCOM.EditText oEdit;

            try
            {
                oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix1").Specific;

                for (int i = 1; i <= Convert.ToInt32(oMatrix.RowCount.ToString()); i++)
                {
                    oEdit = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("col0", i);
                    oEdit.Value = i.ToString().Trim();
                }
            }
            catch (Exception e)
            {
                app.StatusBar.SetText("solicitud_method_setColCount " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        public void llenarComboBox()
        {
            SAPbouiCOM.ComboBox oCombo;

            try
            {
                oCombo = ((SAPbouiCOM.ComboBox)forma.Items.Item("cmbSeries").Specific);
                oCombo.ValidValues.LoadSeries("OZTV", SAPbouiCOM.BoSeriesMode.sf_Add);
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                forma.DataSources.DBDataSources.Item("@ZTV").SetValue("DocNum", 0, forma.BusinessObject.GetNextSerialNumber(oCombo.Selected.Value, "OZTV").ToString());
            }
            catch (Exception e)
            {
                app.StatusBar.SetText("solicitud_method_llenarComboBox " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        /// 
        public void app_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            // se evalua que el boton presionado es el correspondiente al copy to
            // y se se asignan los valores a la clase copiadora la cual se encarga del transporte los itenes
            if(pVal.FormUID == "FSDC" && pVal.FormMode == 1 && pVal.BeforeAction == true && pVal.ItemUID == "btnCpTo" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            {
                try
                {
                    if (!copyTo)
                    {
                        docNum = forma.DataSources.DBDataSources.Item("@ZTV").GetValue("DocNum", 0);
                        copiar = true;
                        copyTo = true;
                        app.ActivateMenuItem("2305");
                    }
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("solicitud_app_event_copiar a -> " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

            // se captura el evento que se produce al clickear o presionar el campo
            // descuento del documento
            if ((pVal.FormUID == "FSDC" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.BeforeAction == false && (pVal.ItemUID == "txtDscPrc" || pVal.ItemUID == "txtDsc"))
                || (pVal.FormUID == "FSDC" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false && (pVal.ItemUID == "txtDscPrc" || pVal.ItemUID == "txtDsc")))
            {
                double discPcnt, discDoc, total;
                SAPbouiCOM.EditText oEdit;

                try
                {
                    switch (pVal.ItemUID)
                    { 
                        case "txtDscPrc":
                            oEdit = (SAPbouiCOM.EditText)forma.Items.Item(pVal.ItemUID).Specific;

                            discDoc = Convert.ToDouble(forma.DataSources.DBDataSources.Item("@ZTV").GetValue("U_TlBfDsc", 0)) * (Convert.ToDouble(oEdit.Value) / 100);

                            forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_DiscPrc", 0, oEdit.Value);
                            forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_Disc", 0, discDoc.ToString());

                            break;
                        case "txtDsc":
                            oEdit = (SAPbouiCOM.EditText)forma.Items.Item(pVal.ItemUID).Specific;

                            total = Convert.ToDouble(forma.DataSources.DBDataSources.Item("@ZTV").GetValue("U_TlBfDsc", 0).ToString());

                            discPcnt = ((Convert.ToDouble(oEdit.Value) / total)*100);

                            forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_Disc", 0, oEdit.Value);
                            forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_DiscPrc", 0, discPcnt.ToString());

                            break;
                        default:
                            break;
                    }
                    calcTaxDoc();
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("solicitud_app_event_descuento -> " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }
            
            // se utiliza esta comprobacion para activar un elemento de la 
            // matrix despues de utilizar el metodo loadfromdatasource()
            if(smf && pVal.FormUID == "FSDC")
            {
                SAPbouiCOM.EditText oEdit;
                SAPbouiCOM.Matrix oMatrix;

                try
                {
                    smf = false;
                    oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix1").Specific;
                    oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(colUID).Cells.Item(rowNum).Specific;
                    oEdit.Active = true;
                    calcTotalBfDsc();
                    calcTaxDoc();
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("solicitud_app_event_semaforo -> " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

            // se registraran los cambios en los campo de calculo como la cantidad y el descuento entre otros
            if(((pVal.FormUID == "FSDC" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.BeforeAction == false && pVal.ItemUID == "matrix1") 
                || (pVal.FormUID == "FSDC" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false && pVal.ItemUID == "matrix1")))
            {
                SAPbouiCOM.Matrix oMatrix;
                SAPbouiCOM.EditText oEdit;
                double qnty;
                double price;

                try
                {    
                    oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix1").Specific;
                    oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colItemCod").Cells.Item(pVal.Row).Specific;

                    // se verifica que el campo codigo cliente no este vacio para proceder a 
                    // realizar los calculos
                    if(oEdit.Value.Trim() != "")
                    {
                        switch (pVal.ColUID)
                        {
                            case "colItemCod":

                                if(pVal.CharPressed == 9)
                                {
                                    colUID = "colDscrip";
                                    smf = true;
                                }
                                break;

                            case "colCant":

                                oMatrix.FlushToDataSource();

                                oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific;

                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_Quantity", pVal.Row - 1, oEdit.Value);

                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_OpenQty", pVal.Row - 1, oEdit.Value);

                                qnty = Convert.ToDouble(forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_Quantity", pVal.Row - 1));
                                price = Convert.ToDouble(forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_PrcBfDi", pVal.Row - 1));

                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_Total", pVal.Row - 1, (qnty * price).ToString());

                                if ("0.0" != forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_DiscItem", pVal.Row - 1))
                                {
                                    goto Descuento;
                                }

                                if (pVal.CharPressed == 9)
                                {
                                    colUID = "colPrecioU";
                                    smf = true;
                                    oMatrix.LoadFromDataSource();
                                }
                                break;

                            case "colPrecioU":

                                oMatrix.FlushToDataSource();

                                oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific;

                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_PrcBfDi", pVal.Row - 1, oEdit.Value);

                                price = Convert.ToDouble(forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_PrcBfDi", pVal.Row - 1));
                                qnty = Convert.ToDouble(forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_Quantity", pVal.Row - 1));
                                
                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_Total", pVal.Row - 1, (qnty * price).ToString());

                                if ("0.0" != forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_DiscItem", pVal.Row - 1))
                                {
                                    goto Descuento;
                                }

                                if (pVal.CharPressed == 9)
                                {
                                    colUID = "colDiscPrc";
                                    smf = true;
                                    oMatrix.LoadFromDataSource();
                                }
                                break;

                            case "colDiscPrc":
                            
                                double temp;
                                double discLine;

                                oMatrix.FlushToDataSource();

                            Descuento:
                                
                                oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colDiscPrc").Cells.Item(pVal.Row).Specific;
                                discLine = Convert.ToDouble(oEdit.Value);

                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_DiscItem", pVal.Row - 1, oEdit.Value);

                                price = Convert.ToDouble(forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_PrcBfDi", pVal.Row - 1));
                                qnty = Convert.ToDouble(forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_Quantity", pVal.Row - 1));
                                temp = (price * qnty);
                                discLine = temp * (discLine / 100);

                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_Total", pVal.Row - 1, (temp - discLine).ToString());

                                if (pVal.CharPressed == 9)
                                {
                                    if (colUID == "colCant")
                                    {
                                        colUID = "colPrecioU";
                                    }
                                    else
                                    {
                                        colUID = "colTaxCod";
                                    }
                                    smf = true;
                                    oMatrix.LoadFromDataSource();
                                }
                                break;
                            default:
                                break;
                        }
                    }
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("solicitud_app_event_calc_prc_qnt_dsc " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

            // se utiliza para manejar todos los eventos que procuren actualizar la forma
            if (pVal.FormUID == "FSDC" && pVal.FormMode == 1 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            {
                try
                {
                    if ((pVal.ItemUID == "txtCodProv" || pVal.ItemUID == "matrix1" || pVal.ItemUID == "txtNmbProv" || pVal.ItemUID == "btnCntPr"))
                    {
                        forma.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                    }
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("solicitud_app_event_mode_edit " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

            // se utiliza esta comprobacion por si dos o mas personas estan utilizando el modulo y deciden procesar una
            // solicitud con el mismo numero de documento 
            if (pVal.FormUID == "FSDC" && pVal.FormMode == 3 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "1" && pVal.BeforeAction == true)
            {
                SAPbouiCOM.ComboBox oCombo;
                SAPbouiCOM.Matrix oMatrix;
                
                try
                {
                    oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix1").Specific;

                    if (oMatrix.RowCount <= 1 && (forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_ItemCode", oMatrix.RowCount).ToString().Trim()) == "")
                    {
                        app.StatusBar.SetText("No se puede procesar la orden - Debe tener al menos un articulo", SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                    }
                    else
                    {
                        if ("" == forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_ItemCode", oMatrix.RowCount - 1).ToString().Trim())
                        {
                            oMatrix.FlushToDataSource();
                            forma.DataSources.DBDataSources.Item("@ZTV_LINES").RemoveRecord(forma.DataSources.DBDataSources.Item("@ZTV_LINES").Size - 1);
                            oMatrix.LoadFromDataSource();
                        }
                    }
                    oCombo = ((SAPbouiCOM.ComboBox)forma.Items.Item("cmbSeries").Specific);
                    oMatrix.FlushToDataSource();
                    forma.DataSources.DBDataSources.Item("@ZTV").SetValue("DocNum", 0, forma.BusinessObject.GetNextSerialNumber(oCombo.Selected.Value, "OZTV").ToString());
                    oMatrix.LoadFromDataSource();
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("solicitud_app_event_mode_add " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

            // luego de que se procesa una orden se agrega una linea nueva en la matrix para se puedan seguir aniadiendo registros
            if (pVal.FormUID == "FSDC" && pVal.FormMode == 3 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "1" && pVal.BeforeAction == false)
            {
                SAPbouiCOM.Matrix oMatrix;
                SAPbouiCOM.Item oButton;

                try
                {
                    oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix1").Specific;
                    if(oMatrix.RowCount < 1)
                    {
                        oMatrix.AddRow(1, 0);
                    }
                    oButton = forma.Items.Item("btnCpTo");
                    llenarComboBox();
                    setFecha();
                    setColCount();
                    forma.ActiveItem = "txtCodProv";
                    oButton.Enabled = false;
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("solicitud_app_event_post_add " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

            // si la solicitud esta en modo busqueda entonces se habilita el campo que contiene el DocNum
            // y el campo series
            if (pVal.FormUID == "FSDC" && pVal.FormMode == 0)
            {
                SAPbouiCOM.Item oButton;

                try
                {
                    oButton = forma.Items.Item("btnCpTo");
                    oButton.Enabled = false;

                    if (!forma.Items.Item("txtSeries").Enabled && !forma.Items.Item("txtStatus").Enabled)
                    {
                        forma.Items.Item("txtSeries").Enabled = true;
                        forma.Items.Item("txtStatus").Enabled = true;
                    }
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("solicitud_app_event_mode_find " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

            // si la solicitud esta en modo de ok que no puedan alterar el tipo de serie con el
            // cual se proceso el documento ni tampoco su numeracion
            if ((pVal.FormUID == "FSDC" && pVal.FormMode == 1) && pVal.BeforeAction == false && pVal.ActionSuccess == true)
            {
                SAPbouiCOM.Item oButton;
                SAPbouiCOM.Matrix oMatrix;

                try
                {
                    oButton = forma.Items.Item("btnCpTo");
                    oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix1").Specific;

                    if (forma.Items.Item("cmbSeries").Enabled || forma.Items.Item("cmbDcType").Enabled)
                    {
                        forma.Items.Item("cmbSeries").Enabled = false;
                        forma.Items.Item("cmbDcType").Enabled = false;
                    }
                    if (pVal.FormMode == 1)
                    {
                        if("O" == forma.DataSources.DBDataSources.Item("@ZTV").GetValue("Status", 0).ToString())
                        {
                            oButton.Enabled = true;
                        }
                        else
                        {
                            oButton.Enabled = false;
                        }
                        if ("" != forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_ItemCode", oMatrix.RowCount - 1).ToString().Trim())
                        {
                            oMatrix.FlushToDataSource();
                            forma.DataSources.DBDataSources.Item("@ZTV_LINES").InsertRecord(forma.DataSources.DBDataSources.Item("@ZTV_LINES").Size);
                            oMatrix.LoadFromDataSource();
                        }
                    }
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("solicitud_app_event_mode_ok " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

            // si la solicitud esta en modo de UPDATE que no puedan alterar el tipo de serie con el
            // cual se proceso el documento ni tampoco su numeracion
            if ((pVal.FormUID == "FSDC" && pVal.FormMode == 2) && pVal.BeforeAction == true)
            {
                SAPbouiCOM.Item oButton;
                SAPbouiCOM.Matrix oMatrix;

                try
                {
                    oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix1").Specific;
                    oButton = forma.Items.Item("btnCpTo");

                    oButton.Enabled = false;
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "1")
                    {
                        if ("" == forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_ItemCode", oMatrix.RowCount - 1).ToString().Trim())
                        {
                            oMatrix.FlushToDataSource();
                            forma.DataSources.DBDataSources.Item("@ZTV_LINES").RemoveRecord(forma.DataSources.DBDataSources.Item("@ZTV_LINES").Size - 1);
                            oMatrix.LoadFromDataSource();
                        }
                    }
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("solicitud_app_event_mode_update " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }

            // se utiliza para ajustar el rectangulo cuando la forma cambie de tamanio
            if (pVal.FormUID == "FSDC" && pVal.BeforeAction == false && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && pVal.ActionSuccess == true)
            {
                SAPbouiCOM.Item oItem;
                SAPbouiCOM.Item rectangle;
                SAPbouiCOM.Matrix oMatrix;

                try
                {
                    oItem = forma.Items.Item("matrix1");
                    oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;
                    rectangle = forma.Items.Item("rtgl");

                    rectangle.Width = (oItem.Width + 19);
                    rectangle.Height = (oItem.Height + 52);
                    oMatrix.AutoResizeColumns();
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("solicitud_app_event_resize " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }
            
            // se utiliza para capturar dos tipos de eventos:
            // el choose from list y el combo select
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:

                    SAPbouiCOM.IChooseFromListEvent oCflEvento = null;
                    oCflEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                    string sCflId = null;
                    sCflId = oCflEvento.ChooseFromListUID;
                    if (oCflEvento.Before_Action == false)
                    {
                        SAPbouiCOM.DataTable oDataTable = null;
                        oDataTable = oCflEvento.SelectedObjects;
                        if (pVal.ItemUID == "txtCodProv" || pVal.ItemUID == "txtNmbProv" && pVal.BeforeAction == false)
                        {
                            try
                            {
                                forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_CardCode", 0, System.Convert.ToString(oDataTable.GetValue("CardCode", 0)));
                                forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_CardName", 0, System.Convert.ToString(oDataTable.GetValue("CardName", 0)));
                                forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_CntCod", 0, System.Convert.ToString(oDataTable.GetValue("CntctPrsn", 0)));
                            }
                            catch (Exception e)
                            {
                                app.StatusBar.SetText("solicitud_app_event_cfl_BPCode " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            }
                        }

                        if (pVal.ItemUID == "btnCntPr" && pVal.Action_Success)
                        {
                            try
                            {
                                forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_CntCod", 0, Convert.ToString(oDataTable.GetValue(2, 0)));
                            }
                            catch (Exception e)
                            {
                                app.StatusBar.SetText("solicitud_app_event_BPCntName " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            }
                        }

                        if (sCflId == "CFL_7" || sCflId.ToString() == "CFL_5")
                        {
                            SAPbouiCOM.Matrix oMatrix;

                            try
                            {
                                int i = pVal.Row;
                                oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix1").Specific;

                                // la manera en como se agregara los itenes a la matrix siempre sera
                                // en modo edicion de un registro ya que se clickea una columna existente
                                // en este caso una columna sin ningun valor
                                forma.DataSources.DBDataSources.Item("@ZTV_LINES").Clear();
                                oMatrix.FlushToDataSource();
                                
                                if(sCflId == "CFL_5")
                                {
                                    forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_ItemCode", (i-1), oDataTable.GetValue("ItemCode", 0).ToString());
                                    forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_Dscript", (i-1), oDataTable.GetValue("ItemName", 0).ToString());
                                    forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_Quantity", (i-1), "1");
                                    forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_OpenQty", (i-1), forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_Quantity", (i-1)).ToString());
                                    forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_PrcBfDi", (i-1), oDataTable.GetValue("LastPurPrc", 0).ToString());
                                    forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_UnitMsr", (i-1), oDataTable.GetValue("BuyUnitMsr", 0).ToString());
                                    forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_Total", (i - 1), oDataTable.GetValue("LastPurPrc", 0).ToString());

                                    if (oDataTable.GetValue("VATLiable", 0).ToString() == "N")
                                    {
                                        forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_TaxCode", (i - 1), "EX");
                                    }
                                    else
                                    {
                                        forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_TaxCode", (i - 1), "IT");
                                    }
                                }

                                if(sCflId == "CFL_7")
                                {
                                    if ((oDataTable.GetValue("Code", 0).ToString() == "IT") || (oDataTable.GetValue("Code", 0).ToString() == "EX"))
                                    {
                                        forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_TaxCode", (i - 1), oDataTable.GetValue("Code", 0).ToString());
                                    }
                                    else
                                    {
                                        app.StatusBar.SetText("No se puede utilizar este tipo de impuesto.", SAPbouiCOM.BoMessageTime.bmt_Short,
                                        SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    }
                                }

                                if ((i == oMatrix.RowCount) && (sCflId == "CFL_5"))
                                {
                                    forma.DataSources.DBDataSources.Item("@ZTV_LINES").InsertRecord(i);
                                    forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_ItemCode", i, "");
                                    forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_Dscript", i, "");
                                    forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_Quantity", i, "1");
                                    forma.DataSources.DBDataSources.Item("@ZTV_LINES").SetValue("U_OpenQty", i, forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_Quantity", i).ToString());
                                }
                                oMatrix.LoadFromDataSource();
                                smf = true;
                                rowNum = pVal.Row;
                                colUID = pVal.ColUID;
                                
                                calcTotalBfDsc();
                                calcTaxDoc();
                            }
                            catch (Exception e)
                            {
                                app.StatusBar.SetText("solicitud_app_event_matrix1 " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            }  
                        }
                    }
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:

                    SAPbouiCOM.ComboBox oCombo;
                    int lNum = 0;

                    // con este procedimiento se llena el combo de la series
                    if (pVal.ItemUID == "cmbSeries" && pVal.Before_Action == false
                        && pVal.ItemChanged == true)
                    {
                        oCombo = ((SAPbouiCOM.ComboBox)forma.Items.Item("cmbSeries").Specific);
                        try
                        {
                            SAPbouiCOM.EditText oEdit;

                            lNum = forma.BusinessObject.GetNextSerialNumber(oCombo.Selected.Value, "OZTV");
                            oEdit = (SAPbouiCOM.EditText)forma.Items.Item("txtSeries").Specific;
                            oEdit.String = lNum.ToString();
                        }
                        catch (Exception e)
                        {
                            app.StatusBar.SetText("solicitud_app_event_comboSeries " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                    }
                    break;
            }
        }

        public void calcTotalBfDsc()
        {
            SAPbouiCOM.Matrix oMatrix;
            int rowCont;
            double totalBf = 0;

            try
            {
                oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix1").Specific;
                rowCont = oMatrix.RowCount - 1;

                for (int i = 0; i < rowCont; i++)
                {
                    totalBf = totalBf + Convert.ToDouble(forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_Total", (i)));
                }
                    forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_TlBfDsc", 0, totalBf.ToString());
            }
            catch (Exception e)
            {
                app.StatusBar.SetText("solicitud_method_calcTotalBfDsc " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        public void calcTaxDoc()
        {
            SAPbouiCOM.Matrix oMatrix;
            int rowCont;
            double taxDoc, taxLine, totalLn, disc, discDoc;
            string taxCode;
            
            try
            {
                taxLine = 0;
                taxDoc = 0;
                totalLn = 0;
                disc = 0;
                discDoc = 0;
                oMatrix = (SAPbouiCOM.Matrix)forma.Items.Item("matrix1").Specific;
                rowCont = oMatrix.RowCount - 1;
                taxCode = "";

                for (int i = 0; i < rowCont; i++)
                {
                    totalLn = Convert.ToDouble(forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_Total", i));
                    taxCode = forma.DataSources.DBDataSources.Item("@ZTV_LINES").GetValue("U_TaxCode", i).ToString();

                    if ("IT".Trim().Equals(taxCode.Trim()))
                    {
                        taxLine = totalLn * 0.16;
                        taxDoc = taxDoc + taxLine;
                    }
                    
                }

                disc = Convert.ToDouble(forma.DataSources.DBDataSources.Item("@ZTV").GetValue("U_DiscPrc", 0));

                if(disc > 0)
                {
                    taxDoc = taxDoc - (taxDoc * (disc / 100));
                    discDoc = Convert.ToDouble(forma.DataSources.DBDataSources.Item("@ZTV").GetValue("U_TlBfDsc", 0)) * (disc / 100);
                    forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_Disc", 0, discDoc.ToString());
                }

                forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_TaxDoc", 0, taxDoc.ToString());
                double totalDoc = (Convert.ToDouble(forma.DataSources.DBDataSources.Item("@ZTV").GetValue("U_TlBfDsc", 0))
                    - (Convert.ToDouble(forma.DataSources.DBDataSources.Item("@ZTV").GetValue("U_Disc", 0)) - taxDoc));

                forma.DataSources.DBDataSources.Item("@ZTV").SetValue("U_TotalDue",0,totalDoc.ToString());
            }
            catch (Exception e)
            {
                app.StatusBar.SetText("solicitud_method_calcTaxDoc " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
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
    }
}