using System.ComponentModel.DataAnnotations;

namespace WindowsFormsApplication4
{
    public class Contact
    {
        [Key]
        public int contactID { get; set; }

        [Required]
        [StringLength(50)]
        public string fname { get; set; }

        [Required]
        [StringLength(50)]
        public string lname { get; set; }

        //[Required]
        [StringLength(50)]
        public string email { get; set; }

        //[Required]
        [StringLength(15)]
        public string mobilephone { get; set; }

        //[Required]
        [StringLength(10)]
        public string birthdate { get; set; }

        //[Required]
        [StringLength(100)]
        public string address { get; set; }

        //[Required]
        [StringLength(100)]
        public string description { get; set; }

    }
}
