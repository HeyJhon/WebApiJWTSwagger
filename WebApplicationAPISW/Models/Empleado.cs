using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WebApplicationAPISW.Models
{
    public class Empleado
    {
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        [Display(Name = "Edad")]
        public string Age { get; set; }
        [Display(Name = "Salario")]
        public decimal Salary { get; set; }
        public Empleado()
        {

        }

        public List<Empleado> GetDataDummy()
        {
            List<Empleado> lista = new List<Empleado>();
            for (int i = 1; i < 11; i++)
            {
                lista.Add(new Empleado()
                {
                    Id = i,
                    FirstName = $"Nombre {i}",
                    LastName = $"LastName{i}",
                    Age = $"Age {i}",
                    Salary = 1000 * i

                });
            }

            return lista;
        }
    }
}