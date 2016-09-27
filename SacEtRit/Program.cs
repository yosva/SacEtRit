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
            try
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
            }
            catch(Exception e)
            {
                Console.WriteLine(e.ToString());

                Console.WriteLine(
                    @" Créer le rapport de Suivi d'Activité en format XLSX pour ITS Group

                    usage: SacEtRit <heures> <developpeur> <client> [mois] [ouotPath]
 
                    <heures> Nombre d'heures hebdomadaires du contrar de travail (35, 37, 38.5 etc...)
                    <developpeur> Nom et prénom du collaborateur
                    <client> Le client de la mission
                    [mois] Le mois pour lequel on genère le rapport, si non renseigné on prend le mois en cours
                    [outPath] La chemin dans lequel on place le rapport .xlsx, si non renseigné on prend le path du .exe

                    Ex.1 SacEtRit 38,5 ""Yovanis SANTIESTEBAN ALGANZA"" Dekra
                    Ex.2 SacEtRit 38,5 ""Yovanis SANTIESTEBAN ALGANZA"" Dekra 10 D:  "
                );
            }

            Console.ReadKey();
        }
    }
}
