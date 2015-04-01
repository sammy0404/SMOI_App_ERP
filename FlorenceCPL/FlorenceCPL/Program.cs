using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FlorenceCPL
{
    class Program
    {
        public static List<RawMaterials> rawMaterials;
        public static List<BeerType> beertype;
        public static List<ProductieClusters> productieCluster;
        public static IT it;
        public static InvoerParser ip = new InvoerParser();
        public static Marketing marketing;
        static void Main()
        {
            Gegevensgeneratie();
            IsHetVerstandigTeKopen();
        }

        static void Gegevensgeneratie()
        {
            rawMaterials = ip.GenerateRawMaterials();
            beertype = ip.GenerateBeerType();
            productieCluster = ip.GenerateProductionCluster();
            it = ip.FixIT();
            marketing = ip.FixMarketing();
        
        }
        static void IsHetVerstandigTeKopen()
        {
            ip.HowMuchToBuy(beertype);
            UitvoerCalculator uvc = new UitvoerCalculator(rawMaterials, beertype, productieCluster, it, marketing);
        }
        
    }

}
