using System;
using System.Threading;
using System.Windows.Forms;

namespace Facturas
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            bool created;
            using (var mutex = new Mutex(true, "Facturas_SingleInstance_Mutex", out created))
            {
                if (!created)
                {
                    MessageBox.Show("Facturas ya está abierto.", "Facturas",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                Application.ThreadException += (s, e) =>
                    MessageBox.Show("Ocurrió un error inesperado:\n\n" + e.Exception.Message +
                        "\n\nTus facturas ya guardadas no se perdieron.", "Facturas",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);

                Application.Run(new MainForm());
            }
        }
    }
}
