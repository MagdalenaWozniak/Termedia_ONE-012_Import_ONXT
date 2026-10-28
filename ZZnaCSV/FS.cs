using cdn_api;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ImportONXT
{
    public class FS
    {
        
        public string idKntONXT { get; set; }
        public string nrDokumentu { get; set; }
        public string dokumentPierw { get; set; }
        public string korekta { get; set; }
        public string status { get; set; } 
        public string typ { get; set; }
        public short czy_skorygowany { get; set; }
        public DateTime data_wystawienia { get; set; }
        public DateTime data_sprzedazy { get; set; }
        public string knt { get; set; }
        public string nip { get; set; }
        public string ulica { get; set; }
        public string miejscowosc { get; set; }
        public string kod_pocztowy { get; set; }
        public string platnosc { get; set; }
        public DateTime terminPlat {get; set;}
        public double netto { get; set; }
        public double podatek {get; set;}
        public double zaplata {get; set;}
        public double nettoZW {get; set;}
        public double podatekZW {get; set;}
        public double bruttoZW {get; set;}
        public double netto23 {get; set;}
        public double podatek23 {get; set;}	
        public double brutto23 {get; set;}

        public FS( string idKntONXT, string nrDokumentu, string dokumentPierw, string korekta, string status, string typ, string czy_skorygowany, string data_wystawienia, string data_sprzedazy, string knt, string nip, string ulica,
                string miejscowosc, string kod_pocztowy, string platnosc, string terminPlat, string netto, string podatek, string zaplata, string nettoZW, string podatekZW, string bruttoZW, string netto23, string podatek23,
                string brutto23 )
        {
            try
            {
                this.idKntONXT = idKntONXT;
                this.nrDokumentu = nrDokumentu;
                this.dokumentPierw = dokumentPierw;
                this.korekta = korekta;
                this.status = status;
                this.typ = typ;
                this.czy_skorygowany = Convert.ToSByte(czy_skorygowany);
                this.data_wystawienia = Convert.ToDateTime(data_wystawienia);
                this.data_sprzedazy = Convert.ToDateTime(data_sprzedazy);
                this.knt = knt;
                this.nip = nip;
                this.ulica = ulica;
                this.miejscowosc = miejscowosc;
                this.kod_pocztowy = kod_pocztowy;
                this.platnosc = platnosc;
                this.terminPlat = Convert.ToDateTime(terminPlat);
                this.netto = Convert.ToDouble(netto);
                this.podatek = Convert.ToDouble(podatek);
                this.zaplata = Convert.ToDouble(zaplata);
                this.nettoZW = Convert.ToDouble(nettoZW);
                this.podatekZW = Convert.ToDouble(podatekZW);
                this.bruttoZW = Convert.ToDouble(bruttoZW);
                this.netto23 = Convert.ToDouble(netto23);
                this.podatek23 = Convert.ToDouble(podatek23);
                this.brutto23 = Convert.ToDouble(brutto23);
            }
            catch(Exception e)
            {
                MessageBox.Show("Błąd podczas dodawania obiektu FS: " + e.Message, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
    }





    }
}
