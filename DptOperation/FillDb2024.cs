using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System.Data;
using System.Globalization;

namespace DptOperation
{
    public class FillDb2024
    {
        string strConn = "server=localhost;database=caleg_nu2;uid=root;pwd=P@ssw0rd;port=3306";
        string _kecamatan = string.Empty;

        public FillDb2024(string kecamatan)
        {
            _kecamatan = kecamatan;

        }

        public string BacaFileExcel(string filename)
        {
            
            FileInfo existingFile = new FileInfo(filename);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count

                // for (int row = 2; row <= 4; row++)
                for (int row = 2; row <= rowCount; row++)
                {
                    if (row % 2000 == 0)
                    {
                        Console.Write(row);
                        Console.Write("  ");
                    }
                    try
                    {
                        bool isDatang = worksheet.Cells[row, 2].Text == "1";
                        string strKelurahan = worksheet.Cells[row, 1].Text;
                        string strNik = worksheet.Cells[row, 3].Text;
                        string strNama = worksheet.Cells[row, 4].Text;

                        if (isDatang)
                        {
                            int rId = GetIdOfPenduduk2024(strKelurahan, strNik, strNama);
                            Console.WriteLine($"{strNama} --> {rId}");
                            UpdatePenduduk2024(rId);
                        }
                        
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine();
                        Console.WriteLine($"ERR ({row}): {ex.Message}");
                        return ($"ERR ({row}): {ex.Message}");
                    }

                }
            }
            return "Selesai";
        }

        public int GetIdOfPenduduk2024(string pkelurahan, string pnik, string pnama)
        {

            string selectCmdStr = "SELECT id FROM dpt_penduduk_2024 where (kelurahan = @kelurahan and nik = @nik and nama = @nama)";
            //string selectCmdStr = "SELECT * FROM dpt_penduduk_2024 limit 3";

            int sId = 0;
            using (MySqlConnection conn = new MySqlConnection(strConn))
            {
                conn.Open();
                using (MySqlCommand cmd = new MySqlCommand(selectCmdStr, conn))
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.AddWithValue("@kelurahan", pkelurahan);
                    cmd.Parameters.AddWithValue("@nik", pnik);
                    cmd.Parameters.AddWithValue("@nama", pnama);

                    using MySqlDataReader rdr = cmd.ExecuteReader();

                    while (rdr.Read())
                    {
                        sId = rdr.GetInt32(0);
                        // Console.WriteLine("{0}", rdr.GetInt32(0));
                    }

                }
            }

            if (sId == 0)
            {
                Console.WriteLine($"  ## TIDAK ADA: {pnik} / {pnama} / {pkelurahan}");
            }

            return sId;
        }

        public void UpdatePenduduk2024(int pid)
        {

            string updCmdStr = "UPDATE `dpt_penduduk_2024` SET `kunjungan1_tgl` = @ptgl WHERE `id` = @pid;";
            //string selectCmdStr = "SELECT * FROM dpt_penduduk_2024 limit 3";

            int sId = 0;
            using (MySqlConnection conn = new MySqlConnection(strConn))
            {
                conn.Open();
                using (MySqlCommand cmd = new MySqlCommand(updCmdStr, conn))
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.AddWithValue("@ptgl", DateTime.Now);
                    cmd.Parameters.AddWithValue("@pid", pid);                    

                    cmd.ExecuteNonQuery();
                }
            }

        }

        // UPDATE `caleg_nu2`.`dpt_penduduk_2024` SET `kunjungan1_tgl` = @ptgl WHERE `id` = @pid;
    }
}
