using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SacEtRit
{
    interface ISacReport
    {
        void Create(double heures, string developpeur, string client, out int totalJours, int? mois = null, string outPath = null);
    }
}
