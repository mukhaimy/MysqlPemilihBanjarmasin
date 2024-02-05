using DptOperation.Part2;
using OfficeOpenXml;

namespace DptOperation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            T2();
            
        }

        //static void T8()
        //{
        //    ExcelLuarNoNik excelLuar = new ExcelLuarNoNik();

        //    string filepath = @"D:\Development2023\15\endangAgustina\NoNik\DATA TPS KOTA BANJARMASIN.xlsx";
        //    excelLuar.RunNoNIk(filepath, "NoNik", "DATA TPS KOTA BANJARMASIN.xlsx");

        //}

        //static void T7()
        //{
        //    ExcelLuarYusna2 excelLuar = new ExcelLuarYusna2();
        //    string filepathHead = @"D:\Development2023\15\endangAgustina\YUSNA - LINA R\";
        //    string[] fileSet = new string[]
        //    {
        //        // "_ALALAK SELATAN TPS 1 - 37 (222 orang).xlsx",
        //        // "_ALALAK TENGAH TPS 1 - 27 (162 orang).xlsx",
        //        // "_KELURAHAN ANTASAN KECIL TIMUR (174 orang).xlsx",
        //        // "_kelurahan pangeran TPS 1 - 27 (162 orang).xlsx",
        //        // "_KUIN UTARA TPS 1 - 33 (198 orang).xlsx",
        //        // "_SUNGAI ANDAI 1-82 (492 orang).xlsx",
        //        // "_SUNGAI MIAI TPS 1 - 41 (246 orang).xlsx",
        //        // "_SURGI MUFTI TPS 1 - 45 (270 orang).xlsx",
        //        // "_TPS 1 - 35(1).xlsx"
        //    };
        //    foreach (string file1 in fileSet)
        //    {
        //        string filepath = filepathHead + file1;
        //        Console.WriteLine($"*** **** {filepath}");
        //        excelLuar.RunYusna(filepath, "YUSNA - LINA R", file1);

        //    }
        //}

        //static void T6()
        //{
        //    ExcelLuarYusna excelLuar = new ExcelLuarYusna();
        //    string filepathHead = @"D:\Development2023\15\endangAgustina\YUSNA - LINA R\";
        //    string[] fileSet = new string[] 
        //    { 
        //        "ALALAK UTARA TPS 39 (246 orang).xlsx",
        //        "ALALAK UTARA TPS 40(100 orang).xlsx",
        //        "ALALAK UTARA TPS 41 (123 orang).xlsx",
        //        "ALALAK UTARA TPS 42(240 orang).xlsx",
        //        "ALALAK UTARA TPS 45 (337 orang).xlsx",
        //        "FIX SUNGAI ANDAI TPS 81 _ 82 (194 orang).xlsx",
        //        "PANGERAN TPS 1 - 12 (1090 orang).xlsx"
        //    };
        //    foreach (string file1 in fileSet)
        //    {
        //        string filepath = filepathHead + file1;
        //        Console.WriteLine($"*** **** {filepath}");
        //        excelLuar.RunYusnaAlalak(filepath, "YUSNA - LINA R", file1);

        //    }
        //}

        //static void T5()
        //{
        //    ExcelLuar excelLuar = new ExcelLuar();
        //    //string filepath = @"D:\Development2023\15\endangAgustina\H. BUDI\basirih selatan.xlsx";
        //    //excelLuar.RunHBudi(filepath, "H. BUDI", "basirih selatan.xlsx");

        //    string filepath = @"D:\Development2023\15\endangAgustina\H. BUDI\haji_budi_file.xlsx";
        //    excelLuar.RunHBudi(filepath, "H. BUDI", "haji_budi_file.xlsx");
        //}

        //static void T4()
        //{
        //    ExcelLuar excelLuar = new ExcelLuar();
        //    string filepath = @"D:\Development2023\15\endangAgustina\SYAHREZA\Daftar Nama.xlsx";
        //    // excelLuar.RunSyahreza(filepath, "SYAHREZA", "Daftar Nama.xlsx");
        //}

        //static void T3()
        //{
        //    Part2.MainContext context = new Part2.MainContext();
        //    context.TheDemoSet.Add(new TheDemo() { Name = "Tom" });
        //    context.TheDemoSet.Add(new TheDemo() { Name = "Hank" });
        //    context.SaveChanges();
        //}


        static void T2()
        {
            string strConn = "server=localhost;database=caleg_ea_bjm02;uid=root;pwd=P@ssw0rd;port=3306";

            string[] filenameSet =  System.IO.Directory.GetFiles(@"E:\Development2024\04\Data Caleg\drive-download-20240205T033112Z-001");
            foreach (string file1 in filenameSet)
            {
                Console.WriteLine(file1);
                Console.WriteLine("#################################");
                Part2Ops ops =new Part2Ops(file1, strConn);
                ops.InsertFromFile();
            }
            //string filename = @"D:\Development2023\15\demo.xlsx";
            //Part2Ops ops =new Part2Ops(filename, strConn);
            //ops.InsertFromFile();
            //Part2.ExcelReader reader = new Part2.ExcelReader();
            //string filename = @"D:\Development2023\15\demo.xlsx";
            //reader.Run1(filename);
        }

        static void T1()
        {
            // F1();
            //DptTpsOperation op = new DptTpsOperation();            
            //op.TR2("KABUPATEN TANAH LAUT");
            //// op.TestGetNik();
            FillDb2024 f = new FillDb2024("BAJUIN");
            string path1 = @"E:\Development2023\17\new-data\s1\";
            for (int i = 3; i <= 11; i++)
            {
                Console.Write(i);
                Console.WriteLine("  --------------------------");
                f.BacaFileExcel(path1 + i.ToString() + ".xlsx");
            }


        }

    }
}