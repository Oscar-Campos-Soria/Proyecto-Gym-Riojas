using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp6
{
    class Cliente : Persona
    {
        public string Membresia { get; set; }

        // Constructor
        public Cliente(string nombres, string apellidos, DateTime fechaNcimiento, string membresia)
            : base(nombres, apellidos, fechaNcimiento)
        {
            Membresia = membresia;
        }


        public override string ToString()
        {
            return base.ToString() + $" - Membresía: {Membresia}";
        }
    }
}