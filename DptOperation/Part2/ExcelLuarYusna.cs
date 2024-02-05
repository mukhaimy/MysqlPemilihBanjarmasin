using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DptOperation.Part2
{
    public class ExcelLuarYusna
    {
        
        public int RunYusnaAlalak(string filepath, string namafolder, string namafile)
        {
            int total = 0;
            try
            {
                MainContext context = new MainContext();
                FileInfo existingFile = new FileInfo(filepath);
                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    int totalSheet = package.Workbook.Worksheets.Count;
                    // baca per sheet --> dapat list
                    for (int i = 0; i < totalSheet; i++)
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[i];
                        int sheetId = i;
                        List<Luar> lstLuar =
                            ReadDptYusnaAlalak(worksheet, namafolder, namafile, sheetId, worksheet.Name);
                        context.LuarSet.AddRange(lstLuar);
                        context.SaveChanges();
                        Console.WriteLine($"ReadDpt on: {namafile} -- {sheetId}.{worksheet.Name}");
                        total += lstLuar.Count;
                    }
                }
                Console.WriteLine();
                Console.WriteLine($"------------------ TOTAL on read: {total}");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return total;
        }
        private List<Luar> ReadDptYusnaAlalak(ExcelWorksheet worksheet,
            string namafolder, string namafile,
            int sheetId, string sheetName)
        {
            DateTime time0 = DateTime.Now;


            int colCount = worksheet.Dimension.End.Column;  //get Column Count
            int rowCount = worksheet.Dimension.End.Row;     //get row count

            // string namaKecamatan = worksheet.Cells[7, 8].Text.Replace(":", "").Trim();
            // string namaKelurahan = worksheet.Cells[8, 8].Text.Replace(":", "").Trim();
            // string kodeTps = worksheet.Cells[9, 8].Text.Replace(":", "").Trim();

            List<Luar> lstPemilihTps = new List<Luar>();

            string kKelurahan = worksheet.Cells[1, 1].Text.Trim().ToUpper();
            string kTps = worksheet.Cells[2, 1].Text.Replace("TPS", "").Trim().ToUpper();
            if(kTps.Length == 1)
            {
                kTps = "00" + kTps;
            }
            else if(kTps.Length == 2)
            {
                kTps = "0" + kTps;
            }
            for (int row = 5; row <= rowCount; row++)
            // for (int row = 12; row < 16; row++)
            {

                // uji apakah ada data berdasarkan nomor urut
                string cell1 = worksheet.Cells[row, 1].Text;
                int icell1;
                if (!int.TryParse(cell1, out icell1))
                {
                    continue;
                }

                if (row % 60 == 0)
                {
                    Console.WriteLine($"[{sheetId}] - row: {row}");
                }

                Luar p1 = new Luar();
                int xusia = 10;
                int xno_urut = 0;
                p1.no_urut = 0;
                if (int.TryParse(worksheet.Cells[row, 1].Text, out xno_urut))
                {
                    p1.no_urut = xno_urut;
                }
                p1.tps = kTps;
                p1.nama = worksheet.Cells[row, 2].Text.Trim().ToUpper();
                p1.jenis_kelamin = "";
                p1.kecamatan = "";
                p1.kelurahan = kKelurahan;
                p1.rt = worksheet.Cells[row, 3].Text.Trim();
                p1.rw = "";
                p1.nik = worksheet.Cells[row, 4].Text.Trim().ToUpper().Replace("'", "").Replace("\"", "");
                p1.is_dukung = true;
                p1.namafolder = namafolder;
                p1.namafile = namafile;
                p1.nama_sheet = sheetName;
                p1.nomor_sheet = sheetId;

                lstPemilihTps.Add(p1);
            }


            var elTime = DateTime.Now - time0;
            Console.WriteLine($"[{sheetId}] !{elTime.TotalMinutes} menit :: {time0}");
            return lstPemilihTps;
        }


    } // EOF: public class ExcelLuar
}
