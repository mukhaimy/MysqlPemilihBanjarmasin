using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DptOperation.Part2
{
    [Table("pemilih_bjm")]
    public class PemilihBjm
    {
        public int Id { get; set; }

        public int no_urut { get; set; } = 0;

        public string nama { get; set; } = string.Empty;

        public string jenis_kelamin { get; set; } = string.Empty;

        public int usia { get; set; }

        public string rt { get; set; } = string.Empty;

        public string rw { get; set; } = string.Empty;

        public string kelurahan { get; set; } = string.Empty;

        public string kecamatan { get; set; } = string.Empty;

        public string tps { get; set; } = string.Empty;

        public bool is_dukung { get; set; } = false;

        public string nik { get; set; } = string.Empty;



        public static string InsertString(string namaTabel)
        {
            string head = $"INSERT INTO {namaTabel} ";
            var sb = new System.Text.StringBuilder(700);
            sb.AppendLine(@"(");
            sb.AppendLine(@"`no_urut`,");
            sb.AppendLine(@"`nama`,");
            sb.AppendLine(@"`jenis_kelamin`,");
            sb.AppendLine(@"`usia`,");
            sb.AppendLine(@"`rt`,");
            sb.AppendLine(@"`rw`,");
            sb.AppendLine(@"`kelurahan`,");
            sb.AppendLine(@"`kecamatan`,");
            sb.AppendLine(@"`tps`,");
            sb.AppendLine(@"`is_dukung`,");
            sb.AppendLine(@"`nik`) ");
            sb.AppendLine(@"VALUES");
            sb.AppendLine(@"(");
            sb.AppendLine(@"@no_urut,");
            sb.AppendLine(@"@nama,");
            sb.AppendLine(@"@jenis_kelamin,");
            sb.AppendLine(@"@usia,");
            sb.AppendLine(@"@rt,");
            sb.AppendLine(@"@rw,");
            sb.AppendLine(@"@kelurahan,");
            sb.AppendLine(@"@kecamatan,");
            sb.AppendLine(@"@tps,");
            sb.AppendLine(@"@is_dukung,");
            sb.AppendLine(@"@nik);");

            return head + sb.ToString();
        }

        public static string IsExistStr(string namaTabel)
        {
            var sb = new System.Text.StringBuilder(400);
            sb.AppendLine($"SELECT count(id) FROM {namaTabel} ");
            sb.AppendLine(@"where (kecamatan='@kecamatan' and kelurahan = '@kelurahan' and tps='@tps' and no_urut = '@no_urut');");

            return sb.ToString();
        }


    }
}
