using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;



namespace FlorenceCPL
{

    class InvoerParser
    {
        public List<RawMaterials> rawMaterials = new List<RawMaterials>();
        public List<BeerType> beerType = new List<BeerType>();
        public List<ProductieClusters> PCL = new List<ProductieClusters>();
        public IT it;
        public void InvoerParsers()
        {


        }
        public List<RawMaterials> GenerateRawMaterials()
        {
            RawMaterials nieuw;
            StreamReader sr = new StreamReader("rawMaterials.csv");
            string invoer;
            string[] readedLine;
            while ((invoer = sr.ReadLine()) != null)
            {
                readedLine = invoer.Replace('.', ',').Split(';');
                nieuw = new RawMaterials(readedLine);
                rawMaterials.Add(nieuw);
            }
            return rawMaterials;
        }
        public IT FixIT()
        {
            StreamReader sr = new StreamReader("it.csv");
            string invoer;
            invoer = sr.ReadLine();
            invoer = sr.ReadLine();
            string[] readedLine;
            readedLine = invoer.Replace('.', ',').Split(';');
            it = new IT(readedLine);
            return it;
        }
        public List<BeerType> GenerateBeerType()
        {
            BeerType nieuw;
            StreamReader sr = new StreamReader("beerType.csv");
            string invoer;
            string[] readedLine;

            while ((invoer = sr.ReadLine()) != null)
            {

                readedLine = invoer.Replace('.', ',').Split(';');
                nieuw = new BeerType(readedLine);
                beerType.Add(nieuw);
            }
            return beerType;
        }
        public List<BeerType> HowMuchToBuy(List<BeerType> bt)
        {
            beerType = bt;
            int teller = 0;
            BeerType biert;
            StreamReader sr = new StreamReader("whatToBuy.csv");
            string invoer;
            string[] readedLine;
            while ((invoer = sr.ReadLine()) != null)
            {
                biert = beerType[teller];
                readedLine = invoer.Replace('.', ',').Split(';');
                biert.howMuchToBuyNL = int.Parse(readedLine[1]);
                biert.howMuchToBuyBE = int.Parse(readedLine[2]);
                biert.howMuchToBuySW = int.Parse(readedLine[3]);
                biert.prducl = readedLine[4];
                biert.VerkoopPercentage = double.Parse(readedLine[5]);
                teller++;
            }

            return beerType;
        }
        public List<ProductieClusters> GenerateProductionCluster()
        {
            ProductieClusters nieuw;
            StreamReader sr = new StreamReader("productionCL.csv");
            string invoer;
            string test = "";
            string[] readedLine;
            while ((invoer = sr.ReadLine()) != null)
            {
                readedLine = invoer.Replace('.', ',').Split(';');
                if (readedLine[0] != test)
                {
                    test = readedLine[0];
                    nieuw = helperGenerateProductionCluster1(readedLine);
                    PCL.Add(nieuw);
                }
                else
                {
                    nieuw = helperGenerateProductionCluster2(readedLine);
                }

            }


            return PCL;
        }
        public ProductieClusters helperGenerateProductionCluster1(string[] readedLine)
        {
            ProductieClusters nieuw = new ProductieClusters(readedLine);
            BeerType bt = beerType.Find(x => x.beerName == readedLine[1]);
            nieuw.addTypeCapacity(bt, readedLine[2]);
            return nieuw;
        }
        public ProductieClusters helperGenerateProductionCluster2(string[] readedLine)
        {
            ProductieClusters nieuw = PCL.Find(x => x.productieclusternaam == readedLine[0]);
            BeerType bt = beerType.Find(x => x.beerName == readedLine[1]);
            nieuw.addTypeCapacity(bt, readedLine[2]);
            return nieuw;
        }






















    }


}
