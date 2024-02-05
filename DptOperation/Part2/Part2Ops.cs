using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DptOperation.Part2
{
    public class Part2Ops
    {
        // string strConn = "server=localhost;database=dpt_caleg_ea;uid=root;pwd=P@ssw0rd;port=3306";
        string _filename = string.Empty;
        int _totalSheet = 0;
        ExcelReader _reader = new ExcelReader();
        string _strConn = string.Empty;

        public Part2Ops(string filename, string strConn)
        {
            _filename = filename;
            _totalSheet = _reader.GetTotalSheet(filename);
            _strConn = strConn;
        }

        public void InsertFromFile()
        {
            for (int i = 0; i < _totalSheet; i++)
            {
                List<PemilihBjm>? pemilihBjms = _reader.Run1(_filename, i);
                if(pemilihBjms != null)
                {
                    Console.WriteLine($" ++++++++ write to db. SHEET# {i}");
                    InsertFromList(pemilihBjms);
                }
            }
        }

        private void InsertFromList(List<PemilihBjm> pemilihBjms)
        {
            using (MySqlConnection conn = new MySqlConnection(_strConn))
            {
                conn.Open();
                string insertQuery = PemilihBjm.InsertString("pemilih_bjm");
                int trow = 1;
                string tnama = "";
                foreach (var item in pemilihBjms)
                {
                    try
                    {
                        tnama = item.nama;
                        using (MySqlCommand cmd = new MySqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@no_urut", item.no_urut);
                            cmd.Parameters.AddWithValue("@nama", item.nama);
                            cmd.Parameters.AddWithValue("@jenis_kelamin", item.jenis_kelamin);
                            cmd.Parameters.AddWithValue("@usia", item.usia);
                            cmd.Parameters.AddWithValue("@rt", item.rt);
                            cmd.Parameters.AddWithValue("@rw", item.rw);
                            cmd.Parameters.AddWithValue("@kelurahan", item.kelurahan);
                            cmd.Parameters.AddWithValue("@kecamatan", item.kecamatan);
                            cmd.Parameters.AddWithValue("@tps", item.tps);
                            cmd.Parameters.AddWithValue("@is_dukung", item.is_dukung);
                            cmd.Parameters.AddWithValue("@nik", item.nik);

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
            
            } //using (MySqlConnection conn = new MySqlConnection(strConn))

        }

    }
    
}
