using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System.Data;



namespace DptOperation
{
    public class DptTpsOperation
    {
        string strConn = "server=localhost;database=dpt_caleg_ea;uid=root;pwd=P@ssw0rd;port=3306";

        public void TR2(string namaKota)
        {

            string[] folderList = new string[] { "BAJUIN", "bati-bati", "batu ampar",
                "bumi makmur", "jorong", "kintap", "kurau", "PANYIPATAN",
                "PELAIHARI", "takisung", "tambang ulang" };

            string baseFolderPath = @"D:\CalegEA\DPT EXCEL\";

            foreach (var folderItem in folderList)
            // string folderItem = folderList[8];
            {
                DirectoryInfo d = new DirectoryInfo(baseFolderPath + folderItem);
                if (!d.Exists)
                {
                    continue;
                }

                FileInfo[] Files = d.GetFiles("*.xlsx"); //Getting Text files

                Console.WriteLine(d.FullName);
                Console.WriteLine("  ---------------");

                foreach (FileInfo file1 in Files)
                // FileInfo file1 = Files[0];
                {
                    Console.WriteLine(file1.FullName);
                    T1(namaKota, file1.FullName);
                }

                Console.WriteLine();
            }

        }

        // ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        public void T1(string namaKota, string filename)
        {

            FileInfo existingFile = new FileInfo(filename);
            int nWorksheet;
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                nWorksheet = package.Workbook.Worksheets.Count;
            }
            DateTime t0 = DateTime.Now;

            Thread[] threadArray = new Thread[nWorksheet];
            for (int i = 0; i < nWorksheet; ++i)
            // for (int i = 0; i < 2; i++)
            {
                // Console.WriteLine($"i = {i} / {nWorksheet}");
                int sheetId = i;
                threadArray[sheetId] = new Thread(() => RunThread(filename, sheetId, namaKota));
                threadArray[sheetId].Start();
            }

            for (int i = 0; i < nWorksheet; ++i)
            {
                int sheetId = i;
                threadArray[sheetId].Join();                                
            }

            var totalAllThreadTime = DateTime.Now - t0;
            Console.WriteLine($"All threads complete :: {totalAllThreadTime.TotalMinutes}");
        }

        private void RunThread(string filename, int sheetId, string namaKota)
        {
            List<PemilihTps> lstPemilihTps;

            FileInfo existingFile = new FileInfo(filename);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                Console.WriteLine("~~~~~ SHEET: " + sheetId.ToString());
                var worksheetCurr = package.Workbook.Worksheets[sheetId];
                lstPemilihTps = ReadDptWorksheet(worksheetCurr, namaKota, sheetId);
            }
            Console.WriteLine();

            AddPemilihTpsToTabel(lstPemilihTps);
        }

        //private List<PemilihTps> ReadDptWorksheet(IWorksheet worksheet, string namaKota)
        //{
        //    DateTime time0 = DateTime.Now;


        //    MySqlConnection conn = new MySqlConnection(strConn);


        //    // int colCount = worksheet.Columns.Count;  //get Column Count
        //    int rowCount = worksheet.Rows.Count();     //get row count

        //    string namaKecamatan = worksheet[7, 8].Text.Replace(":", "").Trim();
        //    string namaKelurahan = worksheet[8, 8].Text.Replace(":", "").Trim();
        //    string kodeTps = worksheet[9, 8].Text.Replace(":", "").Trim();

        //    List<PemilihTps> lstPemilihTps = new List<PemilihTps>();

        //    // for (int row = 12; row < rowCount; row++)
        //    for (int row = 18; row < 24; row++)
        //    {

        //        // uji apakah ada data berdasarkan nomor urut
        //        //string cell1 = worksheet[row, 1].Text;
        //        //int icell1;
        //        //if (!int.TryParse(cell1, out icell1))
        //        //{
        //        //    continue;
        //        //}

        //        if (row % 60 == 0)
        //        {
        //            Console.WriteLine($"- row: {row}");
        //        }

        //        PemilihTps p1 = new PemilihTps();
        //        int xusia = 0;

        //        p1.nama = worksheet[row, 2].Text;
        //        p1.jenis_kelamin = worksheet[row, 3].Text;
        //        if (int.TryParse(worksheet[row, 4].Text, out xusia))
        //        {
        //            p1.usia = xusia;
        //        }
        //        p1.rt = worksheet[row, 6].Text;
        //        p1.rw = worksheet[row, 7].Text;
        //        p1.kelurahan = worksheet[row, 5].Text;
        //        p1.kecamatan = namaKecamatan;
        //        p1.kota = namaKota;

        //        List<string> niks = GetNikList(conn, p1.nama, p1.kelurahan, p1.kecamatan, p1.usia, p1.rt.TrimStart('0'));
        //        if (niks.Count >= 1)
        //        {
        //            p1.nik = niks[0];
        //            if (niks.Count >= 2)
        //            {
        //                p1.nik2 = niks[1];
        //                if (niks.Count >= 3)
        //                {
        //                    p1.nik3 = niks[2];
        //                }
        //            }
        //        }

        //        p1.tps = kodeTps;

        //        // --- terakhir
        //        lstPemilihTps.Add(p1);
        //    }


        //    var elTime = DateTime.Now - time0;
        //    Console.WriteLine($"!{elTime.TotalMinutes} menit :: {time0}");
        //    return lstPemilihTps;
        //}


        private List<PemilihTps> ReadDptWorksheet(ExcelWorksheet worksheet, string namaKota, int sheetId)
        {
            DateTime time0 = DateTime.Now;


            int colCount = worksheet.Dimension.End.Column;  //get Column Count
            int rowCount = worksheet.Dimension.End.Row;     //get row count

            string namaKecamatan = worksheet.Cells[7, 8].Text.Replace(":", "").Trim();
            string namaKelurahan = worksheet.Cells[8, 8].Text.Replace(":", "").Trim();
            string kodeTps = worksheet.Cells[9, 8].Text.Replace(":", "").Trim();

            List<PemilihTps> lstPemilihTps = new List<PemilihTps>();

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

                PemilihTps p1 = new PemilihTps();
                int xusia = 0;

                p1.nama = worksheet.Cells[row, 2].Text;
                p1.jenis_kelamin = worksheet.Cells[row, 3].Text;
                if (int.TryParse(worksheet.Cells[row, 4].Text, out xusia))
                {
                    p1.usia = xusia;
                }
                p1.rt = worksheet.Cells[row, 6].Text;
                p1.rw = worksheet.Cells[row, 7].Text;
                p1.kelurahan = worksheet.Cells[row, 5].Text;
                p1.kecamatan = namaKecamatan;
                p1.kota = namaKota;

                List<string> niks = GetNikList(p1.nama, p1.kelurahan, p1.kecamatan, p1.usia, p1.rt.TrimStart('0'));
                if (niks.Count >= 1)
                {
                    p1.nik = niks[0];
                    if (niks.Count >= 2)
                    {
                        p1.nik2 = niks[1];
                        if (niks.Count >= 3)
                        {
                            p1.nik3 = niks[2];
                        }
                    }
                }

                p1.tps = kodeTps;

                // --- terakhir
                lstPemilihTps.Add(p1);
            }


            var elTime = DateTime.Now - time0;
            Console.WriteLine($"[{sheetId}] !{elTime.TotalMinutes} menit :: {time0}");
            return lstPemilihTps;
        }

        public List<string> GetNikList(string pnama, string pkelurahan, string pkecamatan, int pumur, string prt)
        {
            if (pkecamatan.Length > 44 || pkelurahan.Length > 44)
            {
                return new List<string>();
            }
            List<string> niks = new List<string>();

            using (MySqlConnection conn = new MySqlConnection(strConn))
            {
                conn.Open();
                using (MySqlCommand cmd = new MySqlCommand("sp_get_nik", conn))
                {                    
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@pnama", pnama);
                    cmd.Parameters.AddWithValue("@pkelurahan", pkelurahan);
                    cmd.Parameters.AddWithValue("@pkecamatan", pkecamatan);
                    cmd.Parameters.AddWithValue("@pumur", pumur);
                    // cmd.Parameters.AddWithValue("@prt", prt);

                    var reader = cmd.ExecuteReader();                    
                    while (reader.Read())
                    {
                        niks.Add(reader[2].ToString());
                    }
                }
            }
                
            if (niks.Count < 1)
            {
                Console.WriteLine($"  ## NIK TIDAK ADA: {pnama} / {pkelurahan} / {pkecamatan} / {pumur}");
            }

            return niks;
        }

        private void AddPemilihTpsToTabel(List<PemilihTps> lstPemilihTps)
        {
            using (MySqlConnection conn = new MySqlConnection(strConn)) 
            {
                conn.Open();
                string insertQuery = PemilihTps.InsertString("dpt_pemilih_tps_2024");
                int trow = 1;
                string tnama = "";
                foreach (var item in lstPemilihTps)
                {
                    try
                    {
                        tnama = item.nama;
                        using (MySqlCommand cmd = new MySqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@nama", item.nama);
                            cmd.Parameters.AddWithValue("@jenis_kelamin", item.jenis_kelamin);
                            cmd.Parameters.AddWithValue("@usia", item.usia);
                            cmd.Parameters.AddWithValue("@rt", item.rt);
                            cmd.Parameters.AddWithValue("@rw", item.rw);
                            cmd.Parameters.AddWithValue("@kelurahan", item.kelurahan);
                            cmd.Parameters.AddWithValue("@kecamatan", item.kecamatan);
                            cmd.Parameters.AddWithValue("@kota", item.kota);
                            cmd.Parameters.AddWithValue("@tps", item.tps);

                            cmd.Parameters.AddWithValue("@nik", item.nik);
                            cmd.Parameters.AddWithValue("@nik2", item.nik2);
                            cmd.Parameters.AddWithValue("@nik3", item.nik3);

                            cmd.ExecuteNonQuery();

                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"{trow}/{tnama} \t {ex.Message}");
                    }
                    finally
                    {
                        ++trow;
                    }
                }
            }
            
            
        }

    }
}

