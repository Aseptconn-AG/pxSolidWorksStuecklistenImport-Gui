using Microsoft.VisualBasic.Logging;

namespace pxSolidWorksStuecklistenImport
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static async Task Main(string[] args)
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.

            //foreach (var arg in args)
            //{
            //    MessageBox.Show("Übergebenes Argument: " + arg);
            //    //Log.WriteLog("Übergebenes Argument: " + arg);
            //}

            // Sicherstellen, dass eine Artikelnummer als Parameter übergeben wurde
            if (args.Length == 0)
            {
                MessageBox.Show("Keine Artikelnummer als Parameter angegeben.");
                //Log.WriteLog("Keine Artikelnummer als Parameter angegeben.");
                return;
            }

            // Extrahiere die Artikelnummer aus den übergebenen Argumenten
            string artikelNr = null;
            //foreach (var arg in args)
            //{
            //    if (arg.StartsWith("dfsArtikelNrLAG"))
            //    {
            //        artikelNr = arg.Split('=')[1].Trim();
            //        break;
            //    }
            //}
            //
            artikelNr = args[0];
            //MessageBox.Show(artikelNr);
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1(artikelNr));
        }
    }
}