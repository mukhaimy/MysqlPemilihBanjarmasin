using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DptOperation.Part2
{
    [Table("luar")]
    public class Luar
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

        //  XX  `namafolder` varchar(50) DEFAULT NULL,
        //   `namafile` varchar(100) DEFAULT NULL,
        //   `nama_sheet` varchar(100) DEFAULT NULL,
        //   `nomor_sheet` int,
        //   `is_checked` tinyint(1) NOT NULL DEFAULT '0',
        //   `is_kembar_ada` tinyint(1) NOT NULL DEFAULT '0',
        
        public string namafolder { get; set; } = string.Empty;

        public string namafile { get; set; } = string.Empty;

        public string nama_sheet { get; set; } = string.Empty;

        public int  nomor_sheet { get; set; } = 0;

        public bool is_checked { get; set; } = false;

        public bool is_kembar_ada { get; set; } = false;


    }
}
