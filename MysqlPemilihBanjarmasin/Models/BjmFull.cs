using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MysqlPemilihBanjarmasin.Models
{
    [Table("bjm_full")]
    public class BjmFull
    {
        public int Id { get; set; }

        public int nomor_urut { get; set; }

        public string kecamatan { get; set; } = string.Empty;

        public string kelurahan { get; set; } = string.Empty;
              
        public string nkk { get; set; } = string.Empty;

        public string nik { get; set; } = string.Empty;
        
        public string nama { get; set; } = string.Empty;

        public string tempat_lahir { get; set; } = string.Empty;

        public DateTime? tanggal_lahir { get; set; }

        public string sts_kawin { get; set; } = string.Empty;

        public string jenis_kelamin { get; set; } = string.Empty;

        public string alamat { get; set; } = string.Empty;

        public string rt { get; set; } = string.Empty;

        public string rw { get; set; } = string.Empty;
        
        public string tps { get; set; } = string.Empty;

        public int usia_2024 { get; set; }

        public bool is_dukung { get; set; } = false;


        public void SetUsia2024()
        {
            if(tanggal_lahir != null)
            {
                usia_2024 = CalculateAge(new DateTime(2024, 2, 1), tanggal_lahir.Value);
            }
        }

        public static int CalculateAge(DateTime baseDate, DateTime birthDate)
        {
            var age = baseDate.Year - birthDate.Year;

            if (birthDate.Date > baseDate.AddYears(-age)) age--;

            return age;
        }
    }


}
