using OfficeOpenXml;

namespace MysqlPemilihBanjarmasin
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            T3();
        
        }

        static void T3()
        {
            Reader.BjmFullReader reader1 = new();
            
            string filepath = @"E:\Development2024\04\MysqlPemilihBanjarmasin\DPT2024 Kecamatan.xlsx";
            reader1.Run(filepath);
        }

        

        static void T1()
        {
            string strDate = "10|12|1965";
            DateTime myDate = DateTime.ParseExact(strDate, "dd|MM|yyyy", System.Globalization.CultureInfo.InvariantCulture);
            Console.WriteLine(myDate.ToLongDateString());
            Console.WriteLine("==========================");

            Data.MainContext context = new Data.MainContext();
            List<Models.TheDemo> lst1 = context.TheDemoSet.ToList();
            foreach (var item in lst1)
            {
                Console.WriteLine($"{item.Id}. {item.Name}" );
            }
        }

        
    }
}
