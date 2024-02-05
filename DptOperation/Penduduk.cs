using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DptOperation
{
    public class Penduduk
    {
        #region Properties  
        
        public int id_penduduk { get; set; }

        public string nkk { get; set; } = string.Empty;

        public string nik { get; set; } = string.Empty;

        public string nama { get; set; } = string.Empty;

        public string lahir_tempat { get; set; } = string.Empty;

        public DateTime? lahir_tanggal { get; set; }
       
        public string jenis_kelamin { get; set; } = string.Empty;

        public string alamat { get; set; } = string.Empty;

        public string rt { get; set; } = string.Empty;

        public string rw { get; set; } = string.Empty;

        public string kelurahan { get; set; } = string.Empty;

        public string kecamatan { get; set; } = string.Empty;

        public string kota { get; set; } = string.Empty;

        public string provinsi { get; set; } = string.Empty;

        public string kode_kelurahan { get; set; } = string.Empty;

        #endregion


        #region Function
        public static string InsertString(string tblname)
        {
            string kode = $"INSERT INTO {tblname} (" +
            "nkk," +
            "nik," +
            "nama," +
            "lahir_tempat," +
            "lahir_tanggal," +
            "jenis_kelamin," +
            "alamat," +
            "rt," +
            "rw," +
            "kelurahan," +
            "kecamatan," +
            "kota," +
            "provinsi," +
            "kode_kelurahan)" +
            "VALUES" +
            "(" +
            "@nkk," +
            "@nik," +
            "@nama," +
            "@lahir_tempat," +
            "@lahir_tanggal," +
            "@jenis_kelamin," +
            "@alamat," +
            "@rt," +
            "@rw," +
            "@kelurahan," +
            "@kecamatan," +
            "@kota," +
            "@provinsi," +
            "@kode_kelurahan);";
            return kode;
        }

        #endregion
    }


}
