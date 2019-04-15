using System;
using System.Windows.Forms;
using E_Money_Nominas.Conexiones;
using E_Money_Nominas.Comunes;

namespace E_Money_Nominas
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
