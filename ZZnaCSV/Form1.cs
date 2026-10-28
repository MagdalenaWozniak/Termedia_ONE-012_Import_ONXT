using cdn_api;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ImportONXT
{
    public partial class Form1 : Form
    {

        string connection_string = "";
        int sesja = 0;
        private int XLWersjaAPI = 20202;
        private int XLTrybWsadowy = 1;
        private int XLWinieta = -1;
        private int XLUtworzWlasnaSesje = 1;
        private string XLSQLServer = "192.168.4.24";
        private string XLSQLDatabase = "ERPXL_Termedia_Test";//"ERPXL_PoloMarket"
        private string XLUser = "ADMIN"; //
        private string XLPass = "NE!w0rk12";//"Dupa2003!";
        public Form1(string connectionString, string xlSqlDatabase)
        {
            InitializeComponent();
            connection_string = connectionString;
            XLSQLDatabase = xlSqlDatabase;
            string[] stringMiesiace = new[] { "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12" };
            Miesiace.Items.AddRange(stringMiesiace);
            string[] stringLata = new[] { "2020", "2021", "2022", "2023", "2024", "2025", "2026", "2027", "2028", "2029", "2030", "2031", "2032", "2033", "2034", "2035", "2036", "2037", "2038", "2039", "2040" };
            Lata.Items.AddRange(stringLata);


        }

        private void button_Click(object sender, EventArgs e)
        {
            if (Convert.ToString(Lata.SelectedItem) == "" || Convert.ToString(Miesiace.SelectedItem) == "")
            {
                MessageBox.Show("Proszę wybrać rok i miesiąc ", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            else
            {
                try
                {
                    //coś  jest nie tak
                    //string date3 = data2.Value.ToString("yyyy-MM-dd"); 
                    //MessageBox.Show("Data3:" + date3, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    //int dataFS = getDate(date3);

                    //MessageBox.Show("DataFS:" + dataFS, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);




                    string FSyQuery = "select ISNULL([AkronimIDKontrahentaBazaONXT],0)  , ISNULL([Nr dokumentu],''), ISNULL([Dokument pierwotny], ''), ISNULL([Korekta], ''), [status], [typ], [czy_skorygowany], [data_wystawienia], " +
                                        " [data_sprzedazy], [Pacjent/kontrahent], ISNULL([nip], ''), ISNULL([case], ''), ISNULL([miejscowosc], ''), ISNULL([kod_pocztowy],''), ISNULL([Forma Płatności],''), ISNULL([Termin płatności], ''), " +
                                        " ISNULL([kwota_netto],0), ISNULL([kwota_podatku],0), ISNULL([Wartość do zapłaty],0), ISNULL([Netto ZW],0), ISNULL([Podatek ZW],0), ISNULL([Brutto ZW],0), ISNULL([Netto 23%],0), ISNULL([Podatek 23%],0), " +
                                        " ISNULL([Brutto 23%] ,0) from openquery(ONXT, 'select * from onxtprod.public.v_eksport_dok_fin') " +
                                        " WHERE MONTH(data_wystawienia)= " + Miesiace.SelectedItem.ToString() + " and YEAR(data_wystawienia)= " + Lata.SelectedItem.ToString() + " and (typ = 'Faktura' ) " +
                                        " and status NOT IN ('usunięty','anulowany', '') and [Wartość do zapłaty]!= 0 and [Forma płatności] is not null" +
                                        " ORDER BY [data_wystawienia], CAST (SUBSTRING([Nr dokumentu], CHARINDEX('/',[Nr dokumentu])+1,CHARINDEX('/',[Nr dokumentu], CHARINDEX('/',[Nr dokumentu])+1)-CHARINDEX('/',[Nr dokumentu])-1) as int)";



                    #region logowanie
                    XLLoginInfo_20202 nowyUzytkownik = new XLLoginInfo_20202();
                    nowyUzytkownik.Wersja = XLWersjaAPI;
                    nowyUzytkownik.Winieta = XLWinieta;
                    nowyUzytkownik.TrybWsadowy = XLTrybWsadowy;
                    nowyUzytkownik.UtworzWlasnaSesje = XLUtworzWlasnaSesje;
                    nowyUzytkownik.SerwerKlucza = XLSQLServer;
                    nowyUzytkownik.Baza = XLSQLDatabase;
                    nowyUzytkownik.OpeIdent = XLUser;
                    nowyUzytkownik.OpeHaslo = XLPass;
                    nowyUzytkownik.PortHTTPSerweraKlucza = 1;
                    nowyUzytkownik.RodzajSerweraKlucza = 0;
                    nowyUzytkownik.SerwerKlucza = "localhost::5000185870";
                    nowyUzytkownik.ProgramID = "ONE-012 Import ONXT";

                    int blad = cdn_api.cdn_api.XLLogin(nowyUzytkownik, ref sesja);

                    if (blad != 0)
                    {
                        MessageBox.Show("Błąd logowania: " + blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    #endregion

                    //MessageBox.Show("FSQuery:" + FSyQuery, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    List<FS> FSy = SelectListRowColumn(FSyQuery, "openquery(ONXT, 'select * from onxtprod.public.v_eksport_dok_fin') ");

                    try
                    {
                        FSyQuery = "select ISNULL([AkronimIDKontrahentaBazaONXT],0)  , ISNULL([Nr dokumentu],''), ISNULL([Dokument pierwotny], ''), ISNULL([Korekta], ''), [status], [typ], [czy_skorygowany], [data_wystawienia], " +
                                            " [data_sprzedazy], [Pacjent/kontrahent], ISNULL([nip], ''), ISNULL([case], ''), ISNULL([miejscowosc], ''), ISNULL([kod_pocztowy],''), ISNULL([Forma Płatności],''), ISNULL([Termin płatności], ''), " +
                                            " ISNULL([kwota_netto],0), ISNULL([kwota_podatku],0), ISNULL([Wartość do zapłaty],0), ISNULL([Netto ZW],0), ISNULL([Podatek ZW],0), ISNULL([Brutto ZW],0), ISNULL([Netto 23%],0), ISNULL([Podatek 23%],0), " +
                                            " ISNULL([Brutto 23%] ,0) from openquery(ONXT, 'select * from onxtprod.public.v_eksport_dok_fin') " +
                                            " WHERE MONTH(data_wystawienia)= " + Miesiace.SelectedItem.ToString() + " and YEAR(data_wystawienia)= " + Lata.SelectedItem.ToString() + " and (typ ='Dokument za świadczenia udzielone uprawnionym (zbiorczy)' ) " +
                                            " and status NOT IN ('usunięty','anulowany', '') and [Wartość do zapłaty]!= 0 and [Forma płatności] is not null" +
                                            " ORDER BY [data_wystawienia], [Nr dokumentu]";


                        List<FS> FSy2 = SelectListRowColumn(FSyQuery, "openquery(ONXT, 'select * from onxtprod.public.v_eksport_dok_fin') ");

                        foreach (FS dok in FSy2)
                        {
                            FSy.Add(dok);
                        }
                    }
                    catch(Exception en) {
                        MessageBox.Show("Błąd dodawania dokumentów NFZ: " + en.ToString() , "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }





                    Pasek pp = new Pasek();
                    pp.Show();
                    pp.pBar1.Visible = true;
                    // Set Minimum to 1 to represent the first file being copied.
                    pp.pBar1.Minimum = 1;
                    // Set Maximum to the total number of files to copy.
                    pp.pBar1.Maximum = FSy.Count() + 1;
                    // Set the initial value of the ProgressBar.
                    pp.pBar1.Value = 1;
                    // Set the Step property to a value of 1 to represent each file being copied.
                    pp.pBar1.Step = 1;

                    
                    foreach (FS faktura in FSy)
                    {
                        pp.pBar1.PerformStep();
                        
                        //MessageBox.Show("Numer :" + faktura.nrDokumentu);
                        string akronim = "";
                        if (faktura.nip != "")
                            akronim = ConnectToSQL("SELECT knt_Akronim FROM [CDN].[KntKarty] WHERE Knt_nip= '" + faktura.nip + "'");
                        else
                        
                            akronim = ConnectToSQL("SELECT knt_Akronim FROM [CDN].[KntKarty] WHERE Knt_Nazwa1 = '" + faktura.knt + "' and Knt_KodP='" + faktura.kod_pocztowy + "'");
                        

                        string czyIstnieje = ConnectToSQL(" if exists(SELECT TOP 1 1 FROM [CDN].[TraNag] WHERE TrN_DokumentObcy='" + faktura.nrDokumentu + "' and (TrN_TrNSeria='FNFZ' or TrN_TrNSeria='FSNZ')) SELECT 1 else SELECT 0");

                        if (akronim == "" && czyIstnieje == "0")
                        {
                           // pp.label2.Text = faktura.nrDokumentu;
                           akronim= dodawanieKnt(faktura);

                        }

                        if (akronim != "" && czyIstnieje == "0")
                        {
                            pp.label2.Text = faktura.nrDokumentu;
                            dodawanieDokumentu(faktura, akronim);

                        }
                        //else
                        //{
                            //if (akronim == "" && czyIstnieje=="0")
                               // MessageBox.Show("Kontrahent " + faktura.knt + " nie istnieje w bazie", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        //}


                    }

                    pp.Dispose();
                    pp.Close();

                    cdn_api.cdn_api.XLLogout(sesja);



                }
                catch (Exception en)
                {
                    MessageBox.Show("Błąd: " + en, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                
                }

                MessageBox.Show("Zakończono generowanie FS", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);


                Application.Exit();
            }
        }

        public void dodawanieDokumentu(FS faktura, string akronim)
        {

            int dokumentID = 0;
            int i_Typ = 0;
            int i_Numer = 0;
            try
            {
              

                int forma = 0;
                switch (faktura.platnosc)
                {
                    case "Karta":
                        {
                            forma = 50;
                            break;
                        }
                    case "Gotówka":
                        {
                            forma = 10;
                            break;
                        }
                    case "Przelew":
                        {
                            forma = 20;
                            break;
                        }
                    default:
                        {
                            forma = 60;
                            break;
                        }
                }

                #region dodawanie dokumentu
                XLDokumentNagInfo_20202 nowyFS = new XLDokumentNagInfo_20202();
                nowyFS.GIDFirma = 0;
                nowyFS.GIDLp = 0;
                nowyFS.GIDNumer = 0;
                nowyFS.GIDTyp = 0;
                nowyFS.Typ = 2033;
                nowyFS.Tryb = 2;
                nowyFS.Wersja = XLWersjaAPI;
                nowyFS.Seria = (faktura.typ == "Faktura" ? "FSNZ" : "<PUS>");
                nowyFS.Akronim = (faktura.typ == "Faktura" ? akronim : "NFZ W-WA");
                nowyFS.Data = Convert.ToInt32(ConnectToSQL("SELECT datediff(day, '18001228', '" + faktura.data_wystawienia.ToString("yyyyMMdd") + "')"));
                nowyFS.DataSpr = Convert.ToInt32(ConnectToSQL("SELECT datediff(day, '18001228', '" + faktura.data_sprzedazy.ToString("yyyyMMdd") + "')"));
                nowyFS.DokumentObcy = faktura.nrDokumentu;
                nowyFS.Forma = forma;
                nowyFS.Termin = Convert.ToInt32(ConnectToSQL("SELECT datediff(day, '18001228', '" + faktura.terminPlat.ToString("yyyyMMdd") + "')"));
                nowyFS.DeklaracjaVAT7 = (faktura.dokumentPierw != "" ? 1 : 2);
                nowyFS.Avista = 1;
                nowyFS.Rejestr= (faktura.typ == "Faktura" ? "FSNZ" : "FNFZ");

               

                int bladDok = cdn_api.cdn_api.XLNowyDokument(sesja, ref dokumentID, nowyFS);
                if (bladDok != 0)
                {
                        MessageBox.Show("Błąd podczas dodawania dokumentu API: " + bladDok+" dla dokumentu: "+faktura.nrDokumentu, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    
                }

                #endregion

                //przypisanie do update
                i_Numer = nowyFS.GIDNumer;
                i_Typ = nowyFS.GIDTyp;

                int return_value = 0;
                XLVatInfo_20202 xLVatInfo_20202 = new XLVatInfo_20202
                {


                    Wersja = XLWersjaAPI,
                    FlagaVat = (faktura.brutto23!=0 ? 1 : 0),
                    StawkaPod = (faktura.brutto23 != 0 ? 2300 : 0),
                    Wartosc = Math.Round(faktura.zaplata, 2, MidpointRounding.AwayFromZero).ToString(),
                    WartoscVat = Math.Round(faktura.podatek, 2, MidpointRounding.AwayFromZero).ToString(),
                    DeklRok = nowyFS.Rok, 
                    DeklMiesiac = nowyFS.Miesiac, 
                };
                try
                {
                    return_value = cdn_api.cdn_api.XLDodajVAT(dokumentID, xLVatInfo_20202);

                    if(return_value!=0)
                     
                        MessageBox.Show("Błąd podczas dodawania VATu API: " + return_value+ "dla dokumentu: "+faktura.nrDokumentu, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch (Exception e)
                {
                    throw new System.Exception("Wystapil wyjatek przy dodawaniu VAT faktury. Kod API: " + return_value + ", Wyjątek: " + e);
                }

                
                if(faktura.dokumentPierw != "")
                {
                    

                    string updateJPK = "UPDATE cdn.TraNag set TrN_DokTypJpk='FP' WHERE TrN_GIDNumer=" + nowyFS.GIDNumer + " and TrN_GIDTyp = " + nowyFS.GIDTyp;
                    Update(updateJPK);


                }



                #region zamykanie dokumentu
                XLZamkniecieDokumentuInfo_20202 closeDokument = new XLZamkniecieDokumentuInfo_20202();
                closeDokument.Wersja = XLWersjaAPI;
                closeDokument.Tryb = 0;
                int blad = cdn_api.cdn_api.XLZamknijDokument(dokumentID, closeDokument);
                //if (blad != 0)

                  //  MessageBox.Show("Błąd podczas zamykania dokumentu API: " + blad + "dla dokumentu: " + faktura.nrDokumentu, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //traNag.Clear();

                #endregion

                string updateStan= "UPDATE cdn.TraNag set TrN_Stan=3, TrN_Aktywny = 0 WHERE TrN_GIDNumer=" + nowyFS.GIDNumer + " and TrN_GIDTyp = " + nowyFS.GIDTyp;
                Update(updateStan);

                string numerVat = "SELECT TrN_VatNumer FRom cdn.TraNag WHERE TrN_GIDNumer=" + nowyFS.GIDNumer + "  and TrN_GIDTyp=" + nowyFS.GIDTyp;
                int nrVat = Convert.ToInt32(ConnectToSQL(numerVat));

                if(nrVat==0)
                {
                    numerVat = "SELECT MAX(Trn_VatNumer) From cdn.TraNag WHERE  TrN_VatTyp =2 AND TrN_VatRejestr ='" + nowyFS.Rejestr + "' AND TrN_VatRok =" + Lata.SelectedItem.ToString() + " AND TrN_VatMiesiac =" + Miesiace.SelectedItem.ToString();
                    string vat = ConnectToSQL(numerVat);
                    //MessageBox.Show("Numer Vat: " + vat+ " dla dokumentu: " + faktura.nrDokumentu, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (vat == "")
                        nrVat = 1;
                    else
                        nrVat = Convert.ToInt32(vat) + 1;


                    string updateNr = "update cdn.TraNag set TrN_VatNumer="+nrVat+" WHERE TrN_GIDNumer = " + nowyFS.GIDNumer + " and TrN_GIDTyp =" + nowyFS.GIDTyp;
                    Update(updateNr);
                }



            }
            catch (Exception e)
            {
                MessageBox.Show("Błąd tworzenia faktury: " + e+" dla dokumentu: "+faktura.nrDokumentu, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                string updateStan = "UPDATE cdn.TraNag set TrN_Stan=3, TrN_Aktywny=0 WHERE TrN_GIDNumer=" + i_Numer + " and TrN_GIDTyp = " + i_Typ;
                Update(updateStan);

            }
            finally
            {
                XLZamkniecieDokumentuInfo_20202 closeDokument = new XLZamkniecieDokumentuInfo_20202();
                closeDokument.Wersja = XLWersjaAPI;
                closeDokument.Tryb = 0;
                cdn_api.cdn_api.XLZamknijDokument(dokumentID, closeDokument);
            }

        }

        public string dodawanieKnt(FS faktura)
        {

            if(faktura.miejscowosc=="")
            {
                return "BŁĘDNE DANE";
            }
            else
            {
                string akronim = "";
                //if(faktura.nip=="")
                //{
                    akronim=(faktura.knt.Length > 11 ?  faktura.knt.Substring(0, 11) :  faktura.knt);
                    akronim += " #" + faktura.idKntONXT;

                //}
                //else
               // {
                //    akronim = "Termedica #" + faktura.idKntONXT;
                //}
                int kntID = 0;
                XLKontrahentInfo_20202 knt = new XLKontrahentInfo_20202();
                knt.Akronim = akronim;
                knt.KodP = faktura.kod_pocztowy;
                knt.Miasto = faktura.miejscowosc;
                knt.Nazwa1 = faktura.knt;
                knt.Ulica = faktura.ulica;
                knt.Wersja = XLWersjaAPI;
                knt.NipE = faktura.nip;
                knt.Status = (faktura.nip != "" ? 1 : 2);
                knt.PlatnikVat = ((faktura.nip != "" && faktura.netto23 != 0) ? 1 : 0);
                knt.LimitTerminowy = 0;
                knt.MaxDniPoTerminie = 999;

                int blad = cdn_api.cdn_api.XLNowyKontrahent(sesja, ref kntID, knt);
                if(blad!=0)
                {
                    MessageBox.Show("Błąd dodawania kontrahenta: "+faktura.knt+"\nKod API: "+blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return "BŁĘDNE DANE";
                }

                return knt.Akronim;
            }
            

            
            return "";
        }

        public int getDate(string date3)
        {
            string date1 = "1800-12-28";
            //string date3 = DateTime.Now.ToString("yyyy-MM-dd");
            string Datequery = "SELECT DATEDIFF(day, '" + date1 + "','" + date3 + "'); ";
            return Convert.ToInt32(ConnectToSQL(Datequery));
        }

        public List<string> SelectListColumn(string query, string srcTableName)
        {
            List<string> str = new List<string>();
            string sttr = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, srcTableName);
                    foreach (DataColumn column in (InternalDataCollectionBase)dataSet.Tables[0].Columns)
                        foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                            str.Add(row[column].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(sttr = sttr + "Bład SQL:" + ex.ToString() + "\n");
            }
            return str;
        }

        public string ConnectToSQL(string query)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    if (dataSet.Tables.Count == 0)
                    {
                        return "";
                    }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        str = row[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                str = str + "Bład SQL:" + ex.ToString() + "\n";
            }
            return str;
        }


        public List<FS> SelectListRowColumn(string query, string srcTableName)
        {
            List<string> str = new List<string>();
            List<FS> myList = new List<FS>();
            string sttr = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, srcTableName);
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        foreach (DataColumn column in (InternalDataCollectionBase)dataSet.Tables[0].Columns)
                        {
                            str.Add(row[column].ToString());

                            //MessageBox.Show(column + ": " + row[column].ToString());
                        }
                        myList.Add(new FS(str[0], str[1], str[2], str[3], str[4], str[5], str[6], str[7], str[8], str[9], str[10], str[11], str[12], str[13], str[14], str[15], str[16], str[17], str[18], str[19],
                                            str[20], str[21], str[22], str[23], str[24]));
                        str.Clear();

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(sttr = sttr + "Bład SQL:" + ex.ToString() + "\n");
            }
            return myList;
        }


        public List<int> SelectListColumnZdr(string query, string srcTableName)
        {
            List<int> str = new List<int>();
            string sttr = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, srcTableName);
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        foreach (DataColumn column in (InternalDataCollectionBase)dataSet.Tables[0].Columns)
                        {
                            str.Add(Convert.ToInt32(row[column].ToString()));

                            //MessageBox.Show(column + ": " + row[column].ToString());
                        }
                       
                    }
                }

                
            }
            catch (Exception ex)
            {
                MessageBox.Show(sttr = sttr + "Bład SQL:" + ex.ToString() + "\n");
            }
            return str;
        }

        public string Update(string sql)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    connection.Open();
                    adapter.UpdateCommand = new SqlCommand(sql, connection);
                    adapter.UpdateCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row updated !! ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return sql;
        }


    }


}
