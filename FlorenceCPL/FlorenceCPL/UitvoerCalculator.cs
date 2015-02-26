using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace FlorenceCPL
{
    class UitvoerCalculator
    {
        public static List<RawMaterials> rawMaterials;
        public static List<BeerType> beertype;
        public static List<ProductieClusters> PCL;
        public static InvoerParser ip = new InvoerParser();
        public int capital = 1000000;
        public double totalcostQ0 = 0;
        public UitvoerCalculator(List<RawMaterials> rw, List<BeerType> bt, List<ProductieClusters> pcl)
        {
            rawMaterials = rw;
            beertype = bt;
            PCL = pcl;
            DoWeHaveToProduce();
        }
        public void DoWeHaveToProduce()
        {
            foreach (BeerType bt in beertype)
            {
                if ((bt.howMuchToBuyNL + bt.howMuchToBuyBE + bt.howMuchToBuySW) > bt.voorraad)
                {
                    Produce(bt);
                }
                else
                {
                    ProduceBeer(bt);
                }
            }
            PrettyPrintingResults();

        }
        public void Produce(BeerType bt)
        {
            if (bt.malt != 0) { DoWeHaverawMaterials(rawMaterials[0], bt.malt, bt); }
            if (bt.naturalWater != 0) { DoWeHaverawMaterials(rawMaterials[1], bt.naturalWater, bt); }
            if (bt.hops != 0) { DoWeHaverawMaterials(rawMaterials[2], bt.hops, bt); }
            if (bt.munichMalt != 0) { DoWeHaverawMaterials(rawMaterials[3], bt.munichMalt, bt); }
            if (bt.caramelMalt != 0) { DoWeHaverawMaterials(rawMaterials[4], bt.caramelMalt, bt); }
            if (bt.wheat != 0) { DoWeHaverawMaterials(rawMaterials[5], bt.wheat, bt); }
            if (bt.rye != 0) { DoWeHaverawMaterials(rawMaterials[6], bt.rye, bt); }
            if (bt.maize != 0) { DoWeHaverawMaterials(rawMaterials[7], bt.maize, bt); }
            if (bt.caneSugar != 0) { DoWeHaverawMaterials(rawMaterials[8], bt.caneSugar, bt); }
            if (bt.coriander != 0) { DoWeHaverawMaterials(rawMaterials[9], bt.coriander, bt); }
            if (bt.orangePeel != 0) { DoWeHaverawMaterials(rawMaterials[10], bt.orangePeel, bt); }

            ProduceBeer(bt);

        

        }

        public void ProduceBeer(BeerType bt)
        {
            if (bt.howMuchToBuyNL > bt.voorraad)
            {
                if (bt.malt != 0) { rawMaterials[0].voorraad -= (bt.howMuchToBuyNL * bt.malt); }
                if (bt.naturalWater != 0) { rawMaterials[1].voorraad -= (bt.howMuchToBuyNL * bt.naturalWater); }
                if (bt.hops != 0) { rawMaterials[2].voorraad -= (bt.howMuchToBuyNL * bt.hops); }
                if (bt.munichMalt != 0) { rawMaterials[3].voorraad -= (bt.howMuchToBuyNL * bt.munichMalt); }
                if (bt.caramelMalt != 0) { rawMaterials[4].voorraad -= (bt.howMuchToBuyNL * bt.caramelMalt); }
                if (bt.wheat != 0) { rawMaterials[5].voorraad -= (bt.howMuchToBuyNL * bt.wheat); }
                if (bt.rye != 0) { rawMaterials[6].voorraad -= (bt.howMuchToBuyNL * bt.rye); }
                if (bt.maize != 0) { rawMaterials[7].voorraad -= (bt.howMuchToBuyNL * bt.maize); }
                if (bt.caneSugar != 0) { rawMaterials[8].voorraad -= (bt.howMuchToBuyNL * bt.caneSugar); }
                if (bt.coriander != 0) { rawMaterials[9].voorraad -= (bt.howMuchToBuyNL * bt.coriander); }
                if (bt.orangePeel != 0) { rawMaterials[10].voorraad -= (bt.howMuchToBuyNL * bt.orangePeel); }
                bt.voorraad += PCL.Find(x => x.productieclusternaam == bt.prducl).typecapacity.Find(y => y.Key.beerName == bt.beerName).Value;
                totalcostQ0 += PCL.Find(x => x.productieclusternaam == bt.prducl).leasePrice;
                bt.producedforNL += bt.howMuchToBuyNL;
                bt.voorraad -= bt.howMuchToBuyNL;
            }
            if (bt.howMuchToBuyBE > bt.voorraad)
            {
                if (bt.malt != 0) { rawMaterials[0].voorraad -= (bt.howMuchToBuyBE * bt.malt); }
                if (bt.naturalWater != 0) { rawMaterials[1].voorraad -= (bt.howMuchToBuyBE * bt.naturalWater); }
                if (bt.hops != 0) { rawMaterials[2].voorraad -= (bt.howMuchToBuyBE * bt.hops); }
                if (bt.munichMalt != 0) { rawMaterials[3].voorraad -= (bt.howMuchToBuyBE * bt.munichMalt); }
                if (bt.caramelMalt != 0) { rawMaterials[4].voorraad -= (bt.howMuchToBuyBE * bt.caramelMalt); }
                if (bt.wheat != 0) { rawMaterials[5].voorraad -= (bt.howMuchToBuyBE * bt.wheat); }
                if (bt.rye != 0) { rawMaterials[6].voorraad -= (bt.howMuchToBuyBE * bt.rye); }
                if (bt.maize != 0) { rawMaterials[7].voorraad -= (bt.howMuchToBuyBE * bt.maize); }
                if (bt.caneSugar != 0) { rawMaterials[8].voorraad -= (bt.howMuchToBuyBE * bt.caneSugar); }
                if (bt.coriander != 0) { rawMaterials[9].voorraad -= (bt.howMuchToBuyBE * bt.coriander); }
                if (bt.orangePeel != 0) { rawMaterials[10].voorraad -= (bt.howMuchToBuyBE * bt.orangePeel); }
                bt.voorraad += PCL.Find(x => x.productieclusternaam == bt.prducl).typecapacity.Find(y => y.Key.beerName == bt.beerName).Value;
                totalcostQ0 += PCL.Find(x => x.productieclusternaam == bt.prducl).leasePrice;
                bt.producedforBE += bt.howMuchToBuyBE;
                bt.voorraad -= bt.howMuchToBuyBE;
            }
            if (bt.howMuchToBuySW > bt.voorraad)
            {
                if (bt.malt != 0) { rawMaterials[0].voorraad -= (bt.howMuchToBuySW * bt.malt); }
                if (bt.naturalWater != 0) { rawMaterials[1].voorraad -= (bt.howMuchToBuySW * bt.naturalWater); }
                if (bt.hops != 0) { rawMaterials[2].voorraad -= (bt.howMuchToBuySW * bt.hops); }
                if (bt.munichMalt != 0) { rawMaterials[3].voorraad -= (bt.howMuchToBuySW * bt.munichMalt); }
                if (bt.caramelMalt != 0) { rawMaterials[4].voorraad -= (bt.howMuchToBuySW * bt.caramelMalt); }
                if (bt.wheat != 0) { rawMaterials[5].voorraad -= (bt.howMuchToBuySW * bt.wheat); }
                if (bt.rye != 0) { rawMaterials[6].voorraad -= (bt.howMuchToBuySW * bt.rye); }
                if (bt.maize != 0) { rawMaterials[7].voorraad -= (bt.howMuchToBuySW * bt.maize); }
                if (bt.caneSugar != 0) { rawMaterials[8].voorraad -= (bt.howMuchToBuySW * bt.caneSugar); }
                if (bt.coriander != 0) { rawMaterials[9].voorraad -= (bt.howMuchToBuySW * bt.coriander); }
                if (bt.orangePeel != 0) { rawMaterials[10].voorraad -= (bt.howMuchToBuySW * bt.orangePeel); }
                bt.voorraad += PCL.Find(x => x.productieclusternaam == bt.prducl).typecapacity.Find(y => y.Key.beerName == bt.beerName).Value;
                totalcostQ0 += PCL.Find(x => x.productieclusternaam == bt.prducl).leasePrice;
                bt.producedforSW += bt.howMuchToBuySW;
                bt.voorraad -= bt.howMuchToBuySW;
            }
        }
        public void DoWeHaverawMaterials(RawMaterials rm, double btamount, BeerType bt)
        {
            if (rm.voorraad < (btamount * bt.howMuchToBuyNL))
            {
                goBuyRawMaterial(rm, btamount, bt);
            }
        }
        public void goBuyRawMaterial(RawMaterials rm, double btamount, BeerType bt)
        {
            while (rm.voorraad < (btamount * bt.howMuchToBuyNL))
            {
                rm.voorraad += rm.minOrderAmount;
                totalcostQ0 += (rm.minOrderAmount * rm.priceBuy);
                rm.amountBoughtThisQuarter += rm.minOrderAmount;
            }

        }


        static ExcelPackage package;
        static ExcelWorksheet workbook;
        public void PrettyPrintingResults()
        {
            package = new ExcelPackage(new MemoryStream());
            workbook = package.Workbook.Worksheets.Add("Results");



            //DOE EEN HELEBOEL EXCEL SHIT 
            saveAsExcelSheet();
        }

        public static void saveAsExcelSheet()
        {
            bool saved = false;
            while (!saved)
                try
                {
                    var outputFile = new FileStream("results.xlsx", FileMode.Create);
                    package.SaveAs(outputFile);
                    saved = true;
                }
                catch
                {
                    Console.WriteLine("Please close the file so it can be overwritten, press any key to try again");
                    Console.ReadLine();
                }
        }
    }
}
