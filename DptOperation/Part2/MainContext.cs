using Microsoft.EntityFrameworkCore;

namespace DptOperation.Part2
{
    public class MainContext : DbContext
    {
        public MainContext()
        {
        }

        public MainContext(DbContextOptions<MainContext> options) : base(options)
        {
        }

        
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            // base.OnConfiguring(optionsBuilder);

            string mainConnstr = "server=localhost;database=banjar_ea;uid=root;pwd=P@ssw0rd;port=3306";
            optionsBuilder.UseMySQL(mainConnstr);
        }


        #region DbSet
        public DbSet<TheDemo> TheDemoSet { get; set; }
        
        public DbSet<PemilihBjm> PemilihBjmSet { get; set; }
        public DbSet<Luar> LuarSet { get; set; }
        #endregion

    }
}
