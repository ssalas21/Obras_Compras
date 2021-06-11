using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Permisos_de_obras.Entity
{
    class Listado
    {
        public int AnnoD { get; set; }

        public int AnnoV { get; set; }

        public Listado(int annoD, int annoV) {
            AnnoD = annoD;
            AnnoV = annoV;
        }

    }
}
