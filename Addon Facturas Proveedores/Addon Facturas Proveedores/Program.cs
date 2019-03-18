using System;
using System.Windows.Forms;
using Addon_Facturas_Proveedores.Comunes;
using Addon_Facturas_Proveedores.Conexiones;

namespace Addon_Facturas_Proveedores
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Conexion_SBO oConexion = null;
            oConexion = new Conexion_SBO();
            Eventos_SBO oEvent = null;
            oEvent = new Eventos_SBO();
            Application.Run();
        }
    }
}
