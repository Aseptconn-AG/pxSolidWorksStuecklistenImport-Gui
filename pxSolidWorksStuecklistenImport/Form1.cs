using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.VisualBasic;
using OkoAdvanced;
using System.Data;
using System.Globalization;
using Microsoft.UI.Xaml;
using System;
using Windows.Storage.Pickers;
using WinRT.Interop;
using CsvHelper.Configuration.Attributes;

//using Excel = Microsoft.Office.Interop.Excel;
namespace pxSolidWorksStuecklistenImport
{
    public partial class Form1 : Form
    {
        public static OkoConfig Ini;
        public static OkoDB Db;
        public static OkoLog Log;
        public static string _StlKopf;
        private static string dateiname;
        //private static Excel.Application xlsApplication;
        //private static Excel.Workbook xlsWorkbooks;
        //private static Excel.Worksheet xlsWorksheet;

        //private static int indexArtikelSpalte = 5;//"E"
        //private static int indexMengenSpalte = 2;//"B";
        //private static int indexPositionSpalte = 1;//"C";
        //private static int indexBemerkungsSpalte = 6;//= "F";

        private static int indexArtikelSpalte = 4;//"E"
        private static int indexMengenSpalte = 1;//"B";
        private static int indexPositionSpalte = 0;//"C";
        private static int indexBemerkungsSpalte = 5;//= "F";
        private static string configpath = "\\\\172.16.11.11\\erp$\\PROFFIX\\Prog\\StlImport\\";



        public Form1(string stlKopf = "")
        {
            InitializeComponent();
            _StlKopf = stlKopf;
        }

        private  void Form1_Load(object sender, EventArgs e)
        {
            LoadConfig();
           // await Filesuchen();
        }

        private void LoadConfig()
        {

            Log = new OkoLog();
            string configFilePath = Path.Combine(configpath, "config", "pxSolidWorksStuecklistenImport.ini");
            if (!File.Exists(configFilePath))
            {
                Console.WriteLine("Konfigurationsdatei nicht gefunden: " + configFilePath);
                Log.WriteLog("Konfigurationsdatei nicht gefunden: " + configFilePath);
                return;
            }

            Ini = new OkoConfig(configFilePath, "150757bd-7d48-4fa3-95cd-98454c720e43");
            Ini.Environment = "Prod";
            Db = new OkoDB(Ini, "Prod_Proffix_DB");
            TestDatabaseConnection();
        }
        static void TestDatabaseConnection()
        {
            try
            {

                string sqlQuery = "SELECT 1";
                var dt = Db.ExecuteDataTable(sqlQuery);

                if (dt != null && dt.Rows.Count > 0)
                {
                    Log.WriteLog("Datenbankverbindung erfolgreich.");
                    Console.WriteLine("Datenbankverbindung erfolgreich.");
                }
                else
                {
                    Log.WriteLog("Datenbankverbindung fehlgeschlagen: Keine Zeilen zurückgegeben.");
                    Console.WriteLine("Datenbankverbindung fehlgeschlagen: Keine Zeilen zurückgegeben.");
                }

            }
            catch (Exception ex)
            {
                Log.WriteLog("Datenbankverbindung fehlgeschlagen: " + ex.Message);
                Console.WriteLine("Datenbankverbindung fehlgeschlagen: " + ex.Message);
            }
        }

        static void Importstueckliste()
        {

            string message = "Soll die Stückliste beim Artikel " + _StlKopf + " zuerst geleert werden?";
            string title = "löschen";

            var result1 = MessageBox.Show(message, title, MessageBoxButtons.YesNo);

            if (result1 == DialogResult.Yes)
            {

                string Sql = "DELETE FROM LAG_StuecklistenPos where StlNrLAG='" + _StlKopf + "'";

                Db.ExecuteNonQuery(Sql);
            }
            try
            {
                string stlnrlag;
                string stlnrlagOld = "";
                int position = 0;
                bool lastwithpoint = false;
                //int usedRangeRowsCount = xlsWorksheet.UsedRange.Rows.Count;
                //Excel.Range usedRange = xlsWorksheet.UsedRange;
                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    HasHeaderRecord = true,
                    Delimiter = "\t",
                };

                using (var reader = new StreamReader(dateiname))
                using (var csv = new CsvReader(reader, config))
                {

                    csv.Read();
                    csv.ReadHeader();
                    while (csv.Read())
                    {

                        string cellArtikelValue = csv.GetField(indexArtikelSpalte);

                        if (ArtikelExists(cellArtikelValue))
                        {
                            string cellPositionValue = csv.GetField(indexPositionSpalte);

                            string cellMengenValue = csv.GetField(indexMengenSpalte);

                            string cellBemerkungsValue = csv.GetField(indexBemerkungsSpalte);


                            if (cellPositionValue.Contains("."))
                            {

                                if (!lastwithpoint) 
                                {
                                    message = "Soll die Stückliste beim Artikel " + stlnrlagOld + " zuerst geleert werden?";
                                    result1 = MessageBox.Show(message, title, MessageBoxButtons.YesNo);

                                    if (result1 == DialogResult.Yes)
                                    {

                                        string Sql = "DELETE FROM LAG_StuecklistenPos where StlNrLAG='" + stlnrlagOld + "'";

                                        Db.ExecuteNonQuery(Sql);
                                    }
                                }

                                stlnrlag = stlnrlagOld;
                                lastwithpoint = true;
                                string result = GetStringAfterCharacter(cellPositionValue, '.');
                                position = Convert.ToInt32(result) * 10;
                            }
                            else
                            {
                                stlnrlag = _StlKopf;
                                position = Convert.ToInt32(cellPositionValue) * 10;
                                lastwithpoint = false;
                            }


                            Log.WriteLog(cellMengenValue + @"\"  + cellArtikelValue + @"\" + cellBemerkungsValue + @"\" + position + @"\" + stlnrlag);

                            InsertStücklistepos(cellMengenValue, cellArtikelValue, cellBemerkungsValue, position, stlnrlag);


                            if (!cellPositionValue.Contains("."))
                            {
                                stlnrlagOld = cellArtikelValue;
                            }


                        }
                        else
                        {
                            MessageBox.Show($@" Artikel: {cellArtikelValue} existiert nicht in PROFFIX dieser muss zuert Erstellt werden und dann  entweder händisch in die Stücklisten hinzufügen oder Stücklisten import nochmals durchführen nach dem erstellen.");
                        }

                    }
                }
                MessageBox.Show("Stückliste wurde eingelesen");
           
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //xlsApplication.Quit();
                //Marshal.ReleaseComObject(xlsApplication);
            }

        }

        public static string GetStringAfterCharacter(string input, char character)
        {
            int index = input.IndexOf(character);
            if (index != -1 && index + 1 < input.Length)
            {
                return input.Substring(index + 1);
            }
            return string.Empty; // Return empty string if the character is not found or is the last character
        }
     
        static async Task Filesuchen()
        {
            Log.WriteLog("start Filesuchen" );
         

            var picker = new Windows.Storage.Pickers.FileOpenPicker();
            picker.ViewMode = Windows.Storage.Pickers.PickerViewMode.Thumbnail;
            picker.SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.PicturesLibrary;
            picker.FileTypeFilter.Add(".txt");
            picker.FileTypeFilter.Add(".csv");

            // Initialize with the window handle
            IntPtr hwnd = WindowHandleUtility.GetActiveWindow();
            InitializeWithWindow.Initialize(picker, hwnd);

            Windows.Storage.StorageFile file =  await picker.PickSingleFileAsync();
            dateiname = file.Path;
            if(!string.IsNullOrEmpty(dateiname))
                Importstueckliste();
        }
        public static class WindowHandleUtility
        {
            [System.Runtime.InteropServices.DllImport("user32.dll", ExactSpelling = true, CharSet = System.Runtime.InteropServices.CharSet.Auto)]
            public static extern IntPtr GetActiveWindow();
        }
        static int LaufNrHollenUndErhöhen( string Tabelle)
        {
            try
            {

                int laufnr = 0;

                string Sql = "SELECT * FROM LaufNummern WHERE Tabelle = '" + Tabelle + "'";

                System.Data.DataTable dt = Db.ExecuteDataTable(Sql);

                foreach (DataRow rs in dt.Rows)
                {
                    laufnr = Convert.ToInt32(rs["LaufNR"].ToString()) + 1;

                    Sql = "UPDATE LaufNummern SET LaufNR = " + laufnr.ToString() + " WHERE Tabelle = '" + Tabelle + "'";
                    Db.ExecuteNonQuery(Sql);
                }



                return laufnr;
            }
            catch (Exception ex)
            {
                Log.WriteLog("LaufNrHollenUndErhöhen" + ex.Message);
                return -1; // Return a default value indicating an error occurred
            }
        }

        static Boolean ArtikelExists(string artikel)
        {
            try
            {
                string Sql = "SELECT COUNT(*) as Anz FROM LAG_Artikel where ArtikelNrLAG ='" + artikel + "'";

                System.Data.DataTable dt = Db.ExecuteDataTable(Sql);

                foreach (DataRow rs in dt.Rows)
                {

                    if (Convert.ToInt32(rs["Anz"].ToString()) == 1)
                    {
                        return true;
                    }

                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("LaufNrHollenUndErhöhen" + ex.Message);
                return false; // Return a default value indicating an error occurred
            }
            return false;
        }

        static void InsertStücklistepos(string anzahl, string artikelnr, string bemerkung, int positionnr, string stnrlag)
        {
            //string delSql = "DELETE FROM LAG_StuecklistenPos where StlNrLAG='" + stnrlag + "' and ArtikelnrLAG ='"+artikelnr +"'";

            //Db.ExecuteNonQuery(delSql);

            string Sql = "INSERT INTO dbo.LAG_StuecklistenPos(Anzahl, ArtikelNrLAG, Bemerkung, PositionNr, Preis, StlNrLAG, LaufNr, ImportNr," +
                " ErstelltAm, ErstelltVon, GeaendertAm, GeaendertVon, Geaendert, Exportiert, BemerkungRTF, SprachePRO, PreisSW, Lagerabtrag, NichtAnzeigen, NichtBestellen)" +
                " VALUES" +
                " (" + anzahl + "," +
                " '" + artikelnr + "'," +
                " '" + bemerkung + "'," +
                " " + (positionnr ) + "," +
                " 0," +
                " '" + stnrlag + "'," +
                " " + LaufNrHollenUndErhöhen("LAG_StuecklistenPOS") + "," +
                " 0," +
                " GETDATE()," +
                " 'StuecklistenImport'," +
                " GETDATE()," +
                " 'StuecklistenImport'," +
                " 0," +
                " 0," +
                " NULL," +
                " NULL," +
                " NULL," +
                " 1," +
                " 0," +
                " 0)";
            Log.WriteLog(Sql);
            Db.ExecuteNonQuery(Sql);
        }

        private async void button1_Click(object sender, EventArgs e)
        {
           await Filesuchen();
        }
    }
}
