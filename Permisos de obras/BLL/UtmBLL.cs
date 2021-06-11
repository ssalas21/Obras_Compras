using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Permisos_de_obras.BLL
{
    class UtmBLL
    {
        BaseEntities context;

        public List<Utm> ObtenerTodos() {
            context = new BaseEntities();
            return (from l in context.Utm select l).ToList();
        }

        public Utm ObtenerMesAnno(int anno, int mes) {
            context = new BaseEntities();
            return (from l in context.Utm where l.Anno == anno && l.Mes == mes select l).FirstOrDefault();
        }

    }
}
