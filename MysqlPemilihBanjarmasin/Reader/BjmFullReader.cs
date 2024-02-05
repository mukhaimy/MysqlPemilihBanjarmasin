using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MysqlPemilihBanjarmasin.Reader
{
    public class BjmFullReader
    {
        public int Run(string filepath)
        {
            int total = 0;
            try
            {
                Data.MainContext context = new Data.MainContext();
                FileInfo existingFile = new FileInfo(filepath);
                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    int totalSheet = package.Workbook.Worksheets.Count;
                    // baca per sheet --> dapat list
                    for (int i = 0; i < totalSheet; i++)
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[i];
                        int sheetId = i;
                        List<Models.BjmFull> lst1 = ReadExcel(worksheet, sheetId, worksheet.Name);
                        context.BjmFullSet.AddRange(lst1);
                        context.SaveChanges();
                        Console.WriteLine($"READ on: {sheetId}.{worksheet.Name} -- total: {lst1.Count}");
                        total += lst1.Count;
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

        private List<Models.BjmFull> ReadExcel(ExcelWorksheet worksheet,
            int sheetId, string sheetName)
        {
            DateTime time0 = DateTime.Now;


            int colCount = worksheet.Dimension.End.Column;  //get Column Count
            int rowCount = worksheet.Dimension.End.Row;     //get row count

            List<Models.BjmFull> lst1 = new List<Models.BjmFull>();

            string strTanggalLahir = string.Empty;
            string kTps = string.Empty;

            int nSelKosong = 0;
            for (int row = 2; row <= rowCount; row++)
            // for (int row = 12; row < 16; row++)
            {
                if(nSelKosong > 3)
                {
                    break;
                }
                // uji apakah ada data berdasarkan nomor urut
                string cell1 = worksheet.Cells[row, 1].Text;
                int icell1;
                if (!int.TryParse(cell1, out icell1))
                {
                    ++nSelKosong;
                    continue;
                }

                if (row % 400 == 0)
                {
                    Console.WriteLine($"[{sheetId}] - row: {row}");
                }

                Models.BjmFull p1 = new Models.BjmFull();

                int xno_urut = 0;
                p1.nomor_urut = 0;
                if (int.TryParse(worksheet.Cells[row, 1].Text, out xno_urut))
                {
                    p1.nomor_urut = xno_urut;
                }
                p1.kecamatan = worksheet.Cells[row, 2].Text.Trim().ToUpper();
                p1.kelurahan = worksheet.Cells[row, 3].Text.Trim().ToUpper();
                p1.nkk = worksheet.Cells[row, 4].Text.Trim().ToUpper();
                p1.nik = worksheet.Cells[row, 5].Text.Trim().ToUpper();
                p1.nama = worksheet.Cells[row, 6].Text.Trim().ToUpper();
                p1.tempat_lahir = worksheet.Cells[row, 7].Text.Trim().ToUpper();
                try
                {
                    strTanggalLahir = worksheet.Cells[row, 8].Text.Trim();
                    p1.tanggal_lahir = DateTime.ParseExact(strTanggalLahir, "dd|MM|yyyy", System.Globalization.CultureInfo.InvariantCulture);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ERR: Tanggal Lahir :: " + ex.Message);
                }
                p1.sts_kawin = worksheet.Cells[row, 9].Text.Trim().ToUpper();
                p1.jenis_kelamin = worksheet.Cells[row, 10].Text.Trim().ToUpper();

                if (p1.jenis_kelamin.Length > 2)
                {
                    Console.WriteLine($"JENIS KELAMIN err:{row}");
                }
                p1.alamat = worksheet.Cells[row, 11].Text.Trim().ToUpper();
                p1.rt = worksheet.Cells[row, 12].Text.Trim().ToUpper();
                p1.rw = worksheet.Cells[row, 13].Text.Trim().ToUpper();
                
                kTps = worksheet.Cells[row, 14].Text.Trim().ToUpper();
                if (kTps.Length == 1)
                {
                    kTps = "00" + kTps;
                }
                else if (kTps.Length == 2)
                {
                    kTps = "0" + kTps;
                }
                p1.tps = kTps;
                
                p1.is_dukung = false;
                p1.SetUsia2024();

                lst1.Add(p1);
            }


            var elTime = DateTime.Now - time0;
            Console.WriteLine($"[{sheetId}] !{elTime.TotalMinutes} menit :: {time0}");
            return lst1;
        }

    }
}

