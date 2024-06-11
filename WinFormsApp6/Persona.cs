using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp6
{
    class Persona :IPersona
    {
        // Implementación de la interfaz
        public string ObtenerNombreCompleto()
        {
            return $"{Nombres} {Apellidos}";
        }
        

        private static int contador = 1;
        public int Id { get; set; }
        public string Nombres { get; set; }
        public string Apellidos { get; set; }
        public DateTime FechaNcimiento { get; set; }
        public DateTime FechaRegistro { get; set; }
        

        public Persona(string nombres, string apellidos, DateTime fechaNcimiento)
        {
            Id = contador++;
            Nombres = nombres;
            Apellidos = apellidos;
            FechaNcimiento = fechaNcimiento;
            FechaRegistro = DateTime.Today;
        }

        public Persona()
        {
            
        }

        public override string ToString()
        {
            return $"{Id}: {Nombres} {Apellidos}";
        }
    }
}