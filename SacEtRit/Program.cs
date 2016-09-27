using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SacEtRit
{
    class Program
    {
        static void Main(string[] args)
        {
            SacReport sr = new SacReport();

            int T;
            //sr.Create(38.5, "Yovanis SANTIESTEBAN ALGANZA", "Dekra", out T);//, 9, "D:\\Yovanis\\Emplois\\ITS GROUP\\Suivi missions\\Suivi_Activite_Mensuel-Auto.xlsx");

            int? m = null;
            if (args.Length > 3)
                m = Convert.ToInt32(args[3]);

            string path = null;
            if (args.Length > 4)
                path = args[4];

            sr.Create(Convert.ToDouble(args[0]), args[1], args[2], out T, m, path);

            Console.Write($"Jours travaillés: {T}");

            Console.ReadKey();
        }
    }
}
