using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp6
{
    class Empleado : Persona
    {

        public string Puesto { get; set; }

        public  Empleado(string nombres, string apellidos, DateTime fechaNcimiento, string puesto)
            : base(nombres, apellidos, fechaNcimiento)
        {
            Puesto = puesto;
        }


        public override string ToString()
        {
            return base.ToString() + $" - Puesto: {Puesto}";
        }
    }
}