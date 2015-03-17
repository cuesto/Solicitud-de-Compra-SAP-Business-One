using System;
using System.Collections.Generic;
using System.Text;
using Solicitudes_de_Compra.modelo.logica;



namespace Solicitudes_de_Compra.Vista.Menues
{
    // como es un componente que no contiene logica todo se especifica en
    // el construtor
   class Menu
    {
       private SAPbouiCOM.Application app = null;
       private SAPbobsCOM.Company oCompany = null;
       private SAPbouiCOM.Menus menues = null;
       private SAPbouiCOM.MenuItem menuItem = null;
       private SAPbouiCOM.MenuCreationParams paquete = null;
       private SolicitudCompra sltdCompra = null;
       
        public Menu()
        {
            crearMenu();
            // se instancia un objeto que contiene las formas y logica de negocios
            sltdCompra = new SolicitudCompra();
            app.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(app_MenuEvento);
        }

        public void crearMenu()
        {
            try
            {
                Config.getConfig().conectarGuiApi(ref app);
                Config.getConfig().setCompany(ref oCompany, ref app);
                menues = app.Menus;
                paquete = (SAPbouiCOM.MenuCreationParams)app.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

                // el menu item con el Id 2304 pertenece al menu de las compras
                menuItem = app.Menus.Item("2304");
                menues = menuItem.SubMenus;

                paquete.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                paquete.UniqueID = "mnuSolicitudCmp";
                paquete.String = "Solicitud de Compra";
                paquete.Enabled = true;
                paquete.Position = 0;

                menues.AddEx(paquete);
                Config.getConfig().guardarComoXML(app.Menus.GetAsXML(), "menu.xml");
            }
            catch
            {
            }      
        }

       /// <summary>
       /// 
       /// </summary>
       /// <param name="pVal"></param>
       /// <param name="BubbleEvent"></param>
        public void app_MenuEvento(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        { 
            BubbleEvent = true;

            if ((pVal.MenuUID == "mnuSolicitudCmp") & (pVal.BeforeAction == false))
            {
                sltdCompra = null;
                try
                {
                    // se instancia un objeto que contiene las formas y logica de negocios
                    sltdCompra = new SolicitudCompra();
                    sltdCompra.setApp(ref app);
                    sltdCompra.setCompany(ref oCompany);
                    // este elemento crea una forma nueva que se vizualizara en la pantalla
                    // cada vez que se presione el menu
                    sltdCompra.crearForma();
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("menu_app_MenuEvent_solicitud " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }
            if ((pVal.MenuUID == "2305") & (pVal.BeforeAction == false))
            {
                Copiadora copiadora = null;
                SAPbouiCOM.Form forma;
                    
                try
                {
                    copiadora = new Copiadora();
                    copiadora.setApp(ref app);
                    copiadora.setCompany(ref oCompany);
                    forma = app.Forms.ActiveForm;

                    copiadora.setForma(ref forma);
                    copiadora.setButtonCpFrom();
                    if (SolicitudCompra.copiar)
                    {
                        copiadora.copyTo(SolicitudCompra.docNum);
                        SolicitudCompra.copiar = false;
                        forma.Items.Item("btnCpFrom").Enabled = true;
                    }

                    // se guarda el elemento xml
                    Config.getConfig().guardarComoXML(forma.GetAsXML(), "Orden de Compra.xml");
                    copiadora.capturarEventos();
                }
                catch (Exception e)
                {
                    app.StatusBar.SetText("menu_app_MenuEvent_ordenCompra " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }
        }
    }
}
