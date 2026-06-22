using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;

namespace Impulse.Data
{
    public class ApplicationDbContext : IdentityDbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
            var connectionString = Database.GetConnectionString(); 
            //Console.WriteLine($"Using connection string: {connectionString}");
        }
        public DbSet<User> Users { get; set; }
    }
    public class User 
    {         
        public int UserID { get; set; }
        public string UserName { get; set; }
        //public string EmpID { get; set; }
        public string Password { get; set; } 
    }
}
