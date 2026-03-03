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

        public DbSet<SVN_Defect_Record_History_test> SVN_Defect_Record_History_test { get; set; }

        public DbSet<SVN_Defect_Cookie> SVN_Defect_Cookie { get; set; }


    }
}
