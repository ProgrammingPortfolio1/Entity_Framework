﻿using System.Data.Entity;

namespace WindowsFormsApplication4
{
    public class ContactContext : DbContext
    {
        public ContactContext() : base("name=ContactContext")
        {
        }

        public DbSet<Contact> Contacts { get; set; }
    }
}
