using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace Solicitudes_de_Compra
{
    // esta es la clase que se utilizara para cada conectar el add-on en el sbo
    public class Config
    {
        private static Config config = null;

        // el modificador de acceso del constructor de la clase
        // es privado para aplicar el patron de disenio singleton
        private Config()
        { }

        public static Config getConfig()
        {
            if (config == null)
            {
                config = new Config();
            }
            return config;
        }

        public void conectarGuiApi(ref SAPbouiCOM.Application app)
        {
            try
            {
                SAPbouiCOM.SboGuiApi conn = new SAPbouiCOM.SboGuiApi();
                //string cadena = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
                string cadena = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
                //string cadena = "";
                
                conn.Connect(cadena);
                conn.AddonIdentifier = "5645523035496D706C656D656E746174696F6E3A5331343030393633343339891936B4434B1E01D7DD48DAB7C6A050FD7D3F35";
                app = conn.GetApplication(-1);
            }
            catch (Exception e)
            {
                app.StatusBar.SetText("config_method_conectarGuiApi " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        // este metodo se usa para inicializar un objeto tipo company
        public void setCompany(ref SAPbobsCOM.Company oCompany, ref SAPbouiCOM.Application app)
        {
            try
            {
                oCompany = (SAPbobsCOM.Company)app.Company.GetDICompany();
            }
            catch(Exception e)
            {
                app.StatusBar.SetText("config_method_setCompany " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }            
        }

        public void guardarComoXML(string cadenaXML, string nombreArchivo)
        {
            System.Xml.XmlDocument xmlDoc = null;
            string rutaArchivo;

            xmlDoc = new System.Xml.XmlDocument();

            //se toma la forma como una cadena XML
            //se omitira ya que no se pueden pasar algunos elementos
            //como parametros por referencia

            //se carga la cadena xml al objeto de documento XML
            xmlDoc.LoadXml(cadenaXML);

            //se obtiene la rutaArchivo de la applicacion
            rutaArchivo = System.IO.Directory.GetParent(Application.StartupPath).ToString();
            
            //se guarda el documento
            xmlDoc.Save((rutaArchivo + @"\" + nombreArchivo));
        }

        public string cargarDesdeXML(string nombreArchivo)
        {
            string cadenaXML = null;
            System.Xml.XmlDocument xmlDoc = null;
            xmlDoc = new System.Xml.XmlDocument();
            string rutaArchivo = null;

            //se obtiene la ruta del documento
            rutaArchivo = System.IO.Directory.GetParent(Application.StartupPath).ToString();
            rutaArchivo = System.IO.Directory.GetParent(rutaArchivo).ToString();

            //se carga al objeto xml
            xmlDoc.Load(rutaArchivo + "\\" + nombreArchivo);
            cadenaXML = xmlDoc.InnerXml.ToString();

            return cadenaXML;
        }
    }
}
