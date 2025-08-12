using Microsoft.EntityFrameworkCore;

namespace DefectManagement.Models
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
                : base(options)
        {
        }
        public DbSet<SVN_quality_reason> SVN_quality_reason { get; set; }

        public DbSet<SVN_Defect_Record_History> SVN_Defect_Record_History { get; set; }

        public DbSet<SVN_Defect_Record_Copy> sVN_Defect_Record_Copy { get; set; }


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<SVN_Defect_Record_Copy>().HasNoKey();
            base.OnModelCreating(modelBuilder);
        }
    }
}
