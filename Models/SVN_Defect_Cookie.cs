using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DefectManagement.Models
{
    public class SVN_Defect_Cookie
    {
        [Key]
        [Column("id")]
        public int id { get; set; }

        [Column("cookie")]
        [StringLength(300)]
        public string cookie { get; set; }
    }
}