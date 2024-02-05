using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DptOperation.Part2
{
    [Table(name: "thedemo")]
    public class TheDemo
    {        
        public int Id { get; set; }

        [Required]
        [StringLength(100)]
        public string Name { get; set; } = string.Empty;

        public static List<TheDemo> Contoh ()
        {
            List<TheDemo> hsl = new ();
            hsl.Add(new TheDemo { Id = 1, Name = "Xavier Garner" });
            hsl.Add(new TheDemo { Id = 2, Name = "Donn Mcknight" });
            hsl.Add(new TheDemo { Id = 3, Name = "Patricia Blackburn" });
            hsl.Add(new TheDemo { Id = 4, Name = "Louella Adkins" });
            return hsl;
        }
    }
}
