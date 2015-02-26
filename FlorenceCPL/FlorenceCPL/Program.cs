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

        public static InvoerParser ip = new InvoerParser();
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
        }
        static void IsHetVerstandigTeKopen()
        {
            ip.HowMuchToBuy(beertype);
            UitvoerCalculator uvc = new UitvoerCalculator(rawMaterials, beertype, productieCluster);
        }
    }

}
