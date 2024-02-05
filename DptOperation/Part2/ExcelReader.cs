using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DptOperation.Part2
{
    public class ExcelReader
    {

        //public void Run(string filename, int sheetId)
        //{
        //    List<PemilihBjm> lstPemilihBjm;

        //    FileInfo existingFile = new FileInfo(filename);
        //    using (ExcelPackage package = new ExcelPackage(existingFile))
        //    {
        //        Console.WriteLine("~~~~~ SHEET: " + sheetId.ToString());
        //        var worksheetCurr = package.Workbook.Worksheets[sheetId];
        //        lstPemilihBjm = ReadDptWorksheet(worksheetCurr, sheetId);
        //    }
        //    Console.WriteLine();

        //    // AddPemilihTpsToTabel(lstPemilihTps);
        //}

        public List<PemilihBjm>? Run1(string filename, int sheetId)
        {
            List<PemilihBjm> lstPemilihBjm;

            FileInfo existingFile = new FileInfo(filename);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                if(sheetId >= package.Workbook.Worksheets.Count)
                {
                    return null;
                }
                Console.WriteLine("~~~~~ SHEET: " + sheetId.ToString());
                var worksheetCurr = package.Workbook.Worksheets[sheetId];
                lstPemilihBjm = ReadDptWorksheet(worksheetCurr, sheetId);
            }
            Console.WriteLine();

            return lstPemilihBjm;

        }

        public int GetTotalSheet(string filename)
        {
            FileInfo existingFile = new FileInfo(filename);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                return package.Workbook.Worksheets.Count;
                
            }
        }

        private List<PemilihBjm> ReadDptWorksheet(ExcelWorksheet worksheet, int sheetId)
        {
            DateTime time0 = DateTime.Now;


            int colCount = worksheet.Dimension.End.Column;  //get Column Count
            int rowCount = worksheet.Dimension.End.Row;     //get row count

            string namaKecamatan = worksheet.Cells[7, 8].Text.Replace(":", "").Trim();
            string namaKelurahan = worksheet.Cells[8, 8].Text.Replace(":", "").Trim();
            string kodeTps = worksheet.Cells[9, 8].Text.Replace(":", "").Trim();

            List<PemilihBjm> lstPemilihTps = new List<PemilihBjm>();

            for (int row = 12; row < rowCount; row++)
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

                PemilihBjm p1 = new PemilihBjm();
                int xusia = 0;
                int xno_urut = 0;
                p1.no_urut = 0;
                if (int.TryParse(worksheet.Cells[row, 1].Text, out xno_urut))
                {
                    p1.no_urut = xno_urut;
                }
                p1.nama = worksheet.Cells[row, 2].Text;
                p1.jenis_kelamin = worksheet.Cells[row, 3].Text;
                if (int.TryParse(worksheet.Cells[row, 4].Text, out xusia))
                {
                    p1.usia = xusia;
                }
                p1.kelurahan = worksheet.Cells[row, 5].Text;
                p1.rt = worksheet.Cells[row, 6].Text;
                p1.rw = worksheet.Cells[row, 7].Text;
                p1.kecamatan = namaKecamatan;
                p1.tps = kodeTps;

                p1.is_dukung = false;
                string strIsDukung = worksheet.Cells[row, 8].Text;
                if (!string.IsNullOrEmpty(strIsDukung))
                {
                    p1.is_dukung = true;
                }
                p1.nik = worksheet.Cells[row, 9].Text;
                // --- terakhir

                if (p1.is_dukung)
                {
                    lstPemilihTps.Add(p1);
                }
            }


            var elTime = DateTime.Now - time0;
            Console.WriteLine($"[{sheetId}] !{elTime.TotalMinutes} menit :: {time0}");
            return lstPemilihTps;
        }

    }
}
