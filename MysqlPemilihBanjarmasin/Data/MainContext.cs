using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MysqlPemilihBanjarmasin.Models;

namespace MysqlPemilihBanjarmasin.Data
{
    public class MainContext : DbContext
    {
        public MainContext()
        {
        }

        public MainContext(DbContextOptions<MainContext> options) : base(options)
        {
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<TheDemo>().HasData(
                new TheDemo { Id = 1, Name = "Aubrey Hamilton " },
                new TheDemo { Id = 2, Name = "Dustin Quinn" }    
            );
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            // base.OnConfiguring(optionsBuilder);


            string connString = System.IO.File.ReadAllText("MysqlConn.dat");

            optionsBuilder.UseMySQL(connString);
        }


        #region DbSet
        public DbSet<TheDemo> TheDemoSet { get; set; }
        public DbSet<BjmFull> BjmFullSet { get; set; }


        #endregion
    }
}
