using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace DptOperation
{
    public class PemilihTps
    {
        public string nama { get; set; } = string.Empty;

        public string jenis_kelamin { get; set; } = string.Empty;

        public int usia { get; set; }

        public string rt { get; set; } = string.Empty;

        public string rw { get; set; } = string.Empty;

        public string kelurahan { get; set; } = string.Empty;

        public string kecamatan { get; set; } = string.Empty;

        public string kota { get; set; } = string.Empty;

        public string nik { get; set; } = string.Empty;

        public string nik2 { get; set; } = string.Empty;

        public string nik3 { get; set; } = string.Empty;

        public string tps { get; set; } = string.Empty;


        public static string InsertString(string namaTabel)
        {
            string head = $"INSERT INTO {namaTabel} ";
            var sb = new System.Text.StringBuilder(300);            
            sb.AppendLine(@"(");
            sb.AppendLine(@"`nama`,");
            sb.AppendLine(@"`jenis_kelamin`,");
            sb.AppendLine(@"`usia`,");
            sb.AppendLine(@"`rt`,");
            sb.AppendLine(@"`rw`,");
            sb.AppendLine(@"`kelurahan`,");
            sb.AppendLine(@"`kecamatan`,");
            sb.AppendLine(@"`kota`,");
            sb.AppendLine(@"`tps`,");
            sb.AppendLine(@"`nik`,");
            sb.AppendLine(@"`nik2`,");
            sb.AppendLine(@"`nik3`)");
            sb.AppendLine(@"VALUES");
            sb.AppendLine(@"(");
            sb.AppendLine(@"@nama,");
            sb.AppendLine(@"@jenis_kelamin,");
            sb.AppendLine(@"@usia,");
            sb.AppendLine(@"@rt,");
            sb.AppendLine(@"@rw,");
            sb.AppendLine(@"@kelurahan,");
            sb.AppendLine(@"@kecamatan,");
            sb.AppendLine(@"@kota,");
            sb.AppendLine(@"@tps,");
            sb.AppendLine(@"@nik,");
            sb.AppendLine(@"@nik2,");
            sb.AppendLine(@"@nik3);");

            return head + sb.ToString();
        }
    }
}
