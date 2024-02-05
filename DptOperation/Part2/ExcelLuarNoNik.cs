using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DptOperation.Part2
{
    public class ExcelLuarNoNik
    {

        #region HBudi1
        public int RunNoNIk(string filepath, string namafolder, string namafile)
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
                            ReadDptNoNik(worksheet, namafolder, namafile, sheetId, worksheet.Name);
                        context.LuarSet.AddRange(lstLuar);
                        context.SaveChanges();
                        Console.WriteLine($"ReadDptNoNik on: {sheetId}.{worksheet.Name}");
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

        private List<Luar> ReadDptNoNik(ExcelWorksheet worksheet,
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
            string kTps = string.Empty;
            string kJenisKelamin = string.Empty;

            for (int row = 2; row <= rowCount; row++)
            // for (int row = 12; row < 16; row++)
            {
                kTps = string.Empty;
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
                p1.nama = worksheet.Cells[row, 2].Text.Trim().ToUpper();
                p1.kelurahan = worksheet.Cells[row, 5].Text.Trim().ToUpper();
                p1.kecamatan = worksheet.Cells[row, 6].Text.Trim().ToUpper();
                // p1.kota = worksheet.Cells[row, 7].Text.Trim().ToUpper();
                kTps = worksheet.Cells[row, 8].Text.Trim();
                if (kTps.Length == 1)
                {
                    kTps = "00" + kTps;
                }
                else if (kTps.Length == 2)
                {
                    kTps = "0" + kTps;
                }
                p1.tps = kTps;

                
                p1.jenis_kelamin = "";
                p1.rt = "";
                p1.rw = "";
                p1.nik = "";
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

        #endregion



    } // EOF: public class ExcelLuar
}
