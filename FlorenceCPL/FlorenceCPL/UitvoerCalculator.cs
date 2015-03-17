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
        public static IT it;
        public static InvoerParser ip = new InvoerParser();
        public int capital = 1000000;
        public int totaalproduced = 0;
        public double totalcostQ0 = 0;
        public UitvoerCalculator(List<RawMaterials> rw, List<BeerType> bt, List<ProductieClusters> pcl, IT newit)
        {
            it = newit;
            rawMaterials = rw;
            beertype = bt;
            PCL = pcl;
            DoWeHaveToProduce();
            ExtraCosts();
            foreach(BeerType b in bt)
            {
                totaalproduced += (b.producedforBE+b.producedforNL+b.producedforSW);
            }
            PrettyPrintingResults();
        }
        public void ExtraCosts()
        {
            StreamReader sr = new StreamReader("ExtraCosts.csv");
            string invoer = sr.ReadLine();
            string[] readedLine = invoer.Split(';');
            for (int i = 0; i < readedLine.Length; i++)
            {
                totalcostQ0 += int.Parse(readedLine[i]);
            }
            
        }
        public void DoWeHaveToProduce()
        {
            foreach (BeerType bt in beertype)
            {
                if ((bt.howMuchToBuyNL + bt.howMuchToBuyBE + bt.howMuchToBuySW) > bt.voorraad)
                {
                    ProduceBeer(bt);
                }
                else
                {
                    bt.producedforNL = bt.howMuchToBuyNL;
                    bt.producedforBE = bt.howMuchToBuyBE;
                    bt.producedforSW = bt.howMuchToBuySW;
                    bt.voorraad -= (bt.howMuchToBuyNL + bt.howMuchToBuyBE + bt.howMuchToBuySW);
                }
            }
        }
        public void ProduceBeer(BeerType bt)
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
            while (bt.howMuchToBuyNL + bt.howMuchToBuyBE + bt.howMuchToBuySW > bt.voorraad)
            {
                int GaIkNuProduceren = PCL.Find(x => x.productieclusternaam == bt.prducl).typecapacity.Find(y => y.Key.beerName == bt.beerName).Value;

                if (bt.malt != 0) { rawMaterials[0].voorraad -= (GaIkNuProduceren * bt.malt); }
                if (bt.naturalWater != 0) { rawMaterials[1].voorraad -= (GaIkNuProduceren * bt.naturalWater); }
                if (bt.hops != 0) { rawMaterials[2].voorraad -= (GaIkNuProduceren * bt.hops); }
                if (bt.munichMalt != 0) { rawMaterials[3].voorraad -= (GaIkNuProduceren * bt.munichMalt); }
                if (bt.caramelMalt != 0) { rawMaterials[4].voorraad -= (GaIkNuProduceren * bt.caramelMalt); }
                if (bt.wheat != 0) { rawMaterials[5].voorraad -= (GaIkNuProduceren * bt.wheat); }
                if (bt.rye != 0) { rawMaterials[6].voorraad -= (GaIkNuProduceren * bt.rye); }
                if (bt.maize != 0) { rawMaterials[7].voorraad -= (GaIkNuProduceren * bt.maize); }
                if (bt.caneSugar != 0) { rawMaterials[8].voorraad -= (GaIkNuProduceren * bt.caneSugar); }
                if (bt.coriander != 0) { rawMaterials[9].voorraad -= (GaIkNuProduceren * bt.coriander); }
                if (bt.orangePeel != 0) { rawMaterials[10].voorraad -= (GaIkNuProduceren * bt.orangePeel); }
                bt.voorraad += GaIkNuProduceren;
                totalcostQ0 += PCL.Find(x => x.productieclusternaam == bt.prducl).leasePrice;
            }
            bt.producedforNL = bt.howMuchToBuyNL;
            bt.producedforBE = bt.howMuchToBuyBE;
            bt.producedforSW = bt.howMuchToBuySW;
            bt.voorraad -= (bt.howMuchToBuyNL + bt.howMuchToBuyBE + bt.howMuchToBuySW);
        }

        public void DoWeHaverawMaterials(RawMaterials rm, double btamount, BeerType bt)
        {
            int Productie = PCL.Find(x => x.productieclusternaam == bt.prducl).typecapacity.Find(y => y.Key.beerName == bt.beerName).Value;
            double nodig = Math.Ceiling((bt.howMuchToBuyNL + bt.howMuchToBuyBE + bt.howMuchToBuySW)*1.0/Productie);
            while (rm.voorraad < ((nodig * Productie) * btamount))
            {
                goBuyRawMaterial(rm, btamount,bt);
            }
        }
        public void goBuyRawMaterial(RawMaterials rm, double btamount, BeerType bt)
        {
            int Productie = PCL.Find(x => x.productieclusternaam == bt.prducl).typecapacity.Find(y => y.Key.beerName == bt.beerName).Value;
            double nodig = Math.Ceiling((bt.howMuchToBuyNL + bt.howMuchToBuyBE + bt.howMuchToBuySW)*1.0/Productie);
            while(rm.voorraad<((nodig*Productie)*btamount))
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
            workbook.Cells[2, 2].Value = "TotaalCost";
            workbook.Cells[2, 3].Value = totalcostQ0;
            workbook.Cells[3, 2].Value = rawMaterials[0].materialName;
            workbook.Cells[3, 3].Value = rawMaterials[0].amountBoughtThisQuarter;
            workbook.Cells[4, 2].Value = beertype[0].beerName;
            workbook.Cells[4, 3].Value = beertype[0].producedforNL * beertype[0].transportPriceNL;


            workbook.Workbook.Worksheets.Add("Costs per Product");
            workbook.Workbook.Worksheets.Add("IT");


            workbook.Workbook.Worksheets["Costs per Product"].Cells.Style.Numberformat.Format = "0.00";



            // X-as van de excelsheet
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 2].Value = "Costs per Product";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 3].Value = "Produced for Netherlands";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 4].Value = "Produced for Belgium";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 5].Value = "Produced for Sweden";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 6].Value = "TransportCost to Netherlands";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 7].Value = "TransportCost to Belgium";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 8].Value = "TransportCost to Sweden";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 9].Value = "TransportCost from Netherlands";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 10].Value = "TransportCost from Belgium";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 11].Value = "TransportCost from Sweden";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 12].Value = "Opslagkosten na terugkomst";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 12].Value = "Gemiddeldemn kosten";

            //Y-as van de excelsheet
            workbook.Workbook.Worksheets["Costs per Product"].Cells[3, 2].Value = beertype[0].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[4, 2].Value = beertype[1].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[5, 2].Value = beertype[2].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[6, 2].Value = beertype[3].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[7, 2].Value = beertype[4].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[8, 2].Value = beertype[5].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[9, 2].Value = beertype[6].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[10, 2].Value = "Totaal";

            //Produced aantal voor Nederland
            workbook.Workbook.Worksheets["Costs per Product"].Cells[3, 3].Value = beertype[0].producedforNL;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[4, 3].Value = beertype[1].producedforNL;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[5, 3].Value = beertype[2].producedforNL;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[6, 3].Value = beertype[3].producedforNL;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[7, 3].Value = beertype[4].producedforNL;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[8, 3].Value = beertype[5].producedforNL;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[9, 3].Value = beertype[6].producedforNL;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[10, 3].Value = beertype[0].producedforNL + beertype[1].producedforNL + beertype[2].producedforNL + beertype[3].producedforNL + beertype[4].producedforNL + beertype[5].producedforNL + beertype[6].producedforNL;

            //Produced aantal voor Belgie
            workbook.Workbook.Worksheets["Costs per Product"].Cells[3, 4].Value = beertype[0].producedforBE;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[4, 4].Value = beertype[1].producedforBE;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[5, 4].Value = beertype[2].producedforBE;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[6, 4].Value = beertype[3].producedforBE;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[7, 4].Value = beertype[4].producedforBE;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[8, 4].Value = beertype[5].producedforBE;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[9, 4].Value = beertype[6].producedforBE;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[10, 4].Value = beertype[0].producedforBE + beertype[1].producedforBE + beertype[2].producedforBE + beertype[3].producedforBE + beertype[4].producedforBE + beertype[5].producedforBE + beertype[6].producedforBE;

            //Produced aantal voor Sweden
            workbook.Workbook.Worksheets["Costs per Product"].Cells[3, 5].Value = beertype[0].producedforSW;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[4, 5].Value = beertype[1].producedforSW;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[5, 5].Value = beertype[2].producedforSW;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[6, 5].Value = beertype[3].producedforSW;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[7, 5].Value = beertype[4].producedforSW;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[8, 5].Value = beertype[5].producedforSW;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[9, 5].Value = beertype[6].producedforSW;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[10, 5].Value = beertype[0].producedforSW + beertype[1].producedforSW + beertype[2].producedforSW + beertype[3].producedforSW + beertype[4].producedforSW + beertype[5].producedforSW + beertype[6].producedforSW;

            //Transportkosten voor Nederland
            workbook.Workbook.Worksheets["Costs per Product"].Cells[3, 6].Value = beertype[0].transportPriceNL * beertype[0].producedforNL;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[4, 6].Value = beertype[1].transportPriceNL * beertype[1].producedforNL;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[5, 6].Value = beertype[2].transportPriceNL * beertype[2].producedforNL;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[6, 6].Value = beertype[3].transportPriceNL * beertype[3].producedforNL;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[7, 6].Value = beertype[4].transportPriceNL * beertype[4].producedforNL;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[8, 6].Value = beertype[5].transportPriceNL * beertype[5].producedforNL;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[9, 6].Value = beertype[6].transportPriceNL * beertype[6].producedforNL;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[10, 6].Value = beertype[0].transportPriceNL * beertype[0].producedforNL + beertype[1].transportPriceNL * beertype[1].producedforNL + beertype[2].transportPriceNL * beertype[2].producedforNL + beertype[3].transportPriceNL * beertype[3].producedforNL + beertype[4].transportPriceNL * beertype[4].producedforNL + beertype[5].transportPriceNL * beertype[5].producedforNL + beertype[6].transportPriceNL * beertype[6].producedforNL;

            //Transportkosten voor Belgie
            workbook.Workbook.Worksheets["Costs per Product"].Cells[3, 7].Value = beertype[0].transportPriceBE * beertype[0].producedforBE;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[4, 7].Value = beertype[1].transportPriceBE * beertype[1].producedforBE;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[5, 7].Value = beertype[2].transportPriceBE * beertype[2].producedforBE;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[6, 7].Value = beertype[3].transportPriceBE * beertype[3].producedforBE;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[7, 7].Value = beertype[4].transportPriceBE * beertype[4].producedforBE;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[8, 7].Value = beertype[5].transportPriceBE * beertype[5].producedforBE;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[9, 7].Value = beertype[6].transportPriceBE * beertype[6].producedforBE;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[10, 7].Value = beertype[0].transportPriceBE * beertype[0].producedforBE + beertype[1].transportPriceBE * beertype[1].producedforBE + beertype[2].transportPriceBE * beertype[2].producedforBE + beertype[3].transportPriceBE * beertype[3].producedforBE + beertype[4].transportPriceBE * beertype[4].producedforBE + beertype[5].transportPriceBE * beertype[5].producedforBE + beertype[6].transportPriceBE * beertype[6].producedforBE;

            //Transportkosten voor Sweden
            workbook.Workbook.Worksheets["Costs per Product"].Cells[3, 8].Value = beertype[0].transportPriceSW * beertype[0].producedforSW;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[4, 8].Value = beertype[1].transportPriceSW * beertype[1].producedforSW;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[5, 8].Value = beertype[2].transportPriceSW * beertype[2].producedforSW;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[6, 8].Value = beertype[3].transportPriceSW * beertype[3].producedforSW;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[7, 8].Value = beertype[4].transportPriceSW * beertype[4].producedforSW;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[8, 8].Value = beertype[5].transportPriceSW * beertype[5].producedforSW;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[9, 8].Value = beertype[6].transportPriceSW * beertype[6].producedforSW;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[10, 8].Value = beertype[0].transportPriceSW * beertype[0].producedforSW + beertype[1].transportPriceSW * beertype[1].producedforSW + beertype[2].transportPriceSW * beertype[2].producedforSW + beertype[3].transportPriceSW * beertype[3].producedforSW + beertype[4].transportPriceSW * beertype[4].producedforSW + beertype[5].transportPriceSW * beertype[5].producedforSW + beertype[6].transportPriceSW * beertype[6].producedforSW;

            //Transportkosten terug voor Nederland
            workbook.Workbook.Worksheets["Costs per Product"].Cells[3, 9].Value = beertype[0].transportPriceNL * beertype[0].producedforNL * (1 - beertype[0].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[4, 9].Value = beertype[1].transportPriceNL * beertype[1].producedforNL * (1 - beertype[1].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[5, 9].Value = beertype[2].transportPriceNL * beertype[2].producedforNL * (1 - beertype[2].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[6, 9].Value = beertype[3].transportPriceNL * beertype[3].producedforNL * (1 - beertype[3].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[7, 9].Value = beertype[4].transportPriceNL * beertype[4].producedforNL * (1 - beertype[4].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[8, 9].Value = beertype[5].transportPriceNL * beertype[5].producedforNL * (1 - beertype[5].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[9, 9].Value = beertype[6].transportPriceNL * beertype[6].producedforNL * (1 - beertype[6].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[10, 9].Value = (beertype[0].transportPriceNL * (1 - beertype[0].VerkoopPercentage) * beertype[0].producedforNL + beertype[1].transportPriceNL * (1 - beertype[1].VerkoopPercentage) * beertype[1].producedforNL + beertype[2].transportPriceNL * (1 - beertype[2].VerkoopPercentage) * beertype[2].producedforNL + beertype[3].transportPriceNL * (1 - beertype[3].VerkoopPercentage) * beertype[3].producedforNL + beertype[4].transportPriceNL * (1 - beertype[4].VerkoopPercentage) * beertype[4].producedforNL + beertype[5].transportPriceNL * (1 - beertype[5].VerkoopPercentage) * beertype[5].producedforNL + beertype[6].transportPriceNL * (1 - beertype[6].VerkoopPercentage) * beertype[6].producedforNL);

            //Transportkosten terug voor Belgie
            workbook.Workbook.Worksheets["Costs per Product"].Cells[3, 10].Value = beertype[0].transportPriceBE * beertype[0].producedforBE * (1 - beertype[0].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[4, 10].Value = beertype[1].transportPriceBE * beertype[1].producedforBE * (1 - beertype[1].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[5, 10].Value = beertype[2].transportPriceBE * beertype[2].producedforBE * (1 - beertype[2].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[6, 10].Value = beertype[3].transportPriceBE * beertype[3].producedforBE * (1 - beertype[3].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[7, 10].Value = beertype[4].transportPriceBE * beertype[4].producedforBE * (1 - beertype[4].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[8, 10].Value = beertype[5].transportPriceBE * beertype[5].producedforBE * (1 - beertype[5].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[9, 10].Value = beertype[6].transportPriceBE * beertype[6].producedforBE * (1 - beertype[6].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[10, 10].Value = (beertype[0].transportPriceBE * (1 - beertype[0].VerkoopPercentage) * beertype[0].producedforBE + beertype[1].transportPriceBE * (1 - beertype[1].VerkoopPercentage) * beertype[1].producedforBE + beertype[2].transportPriceBE * (1 - beertype[2].VerkoopPercentage) * beertype[2].producedforBE + beertype[3].transportPriceBE * (1 - beertype[3].VerkoopPercentage) * beertype[3].producedforBE + beertype[4].transportPriceBE * (1 - beertype[4].VerkoopPercentage) * beertype[4].producedforBE + beertype[5].transportPriceBE * (1 - beertype[5].VerkoopPercentage) * beertype[5].producedforBE + beertype[6].transportPriceBE * (1 - beertype[6].VerkoopPercentage) * beertype[6].producedforBE);

            //Transportkosten terug voor Sweden
            workbook.Workbook.Worksheets["Costs per Product"].Cells[3, 11].Value = beertype[0].transportPriceSW * beertype[0].producedforSW * (1 - beertype[0].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[4, 11].Value = beertype[1].transportPriceSW * beertype[1].producedforSW * (1 - beertype[1].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[5, 11].Value = beertype[2].transportPriceSW * beertype[2].producedforSW * (1 - beertype[2].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[6, 11].Value = beertype[3].transportPriceSW * beertype[3].producedforSW * (1 - beertype[3].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[7, 11].Value = beertype[4].transportPriceSW * beertype[4].producedforSW * (1 - beertype[4].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[8, 11].Value = beertype[5].transportPriceSW * beertype[5].producedforSW * (1 - beertype[5].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[9, 11].Value = beertype[6].transportPriceSW * beertype[6].producedforSW * (1 - beertype[6].VerkoopPercentage);
            workbook.Workbook.Worksheets["Costs per Product"].Cells[10, 11].Value = (beertype[0].transportPriceSW * (1 - beertype[0].VerkoopPercentage) * beertype[0].producedforSW + beertype[1].transportPriceSW * (1 - beertype[1].VerkoopPercentage) * beertype[1].producedforSW + beertype[2].transportPriceSW * (1 - beertype[2].VerkoopPercentage) * beertype[2].producedforSW + beertype[3].transportPriceSW * (1 - beertype[3].VerkoopPercentage) * beertype[3].producedforSW + beertype[4].transportPriceSW * (1 - beertype[4].VerkoopPercentage) * beertype[4].producedforSW + beertype[5].transportPriceSW * (1 - beertype[5].VerkoopPercentage) * beertype[5].producedforSW + beertype[6].transportPriceSW * (1 - beertype[6].VerkoopPercentage) * beertype[6].producedforSW);

            //Opslagkosten na terugkomst totaal
            workbook.Workbook.Worksheets["Costs per Product"].Cells[3, 12].Value = (beertype[0].voorraad + ((beertype[0].producedforNL + beertype[0].producedforBE + beertype[0].producedforSW) * (1 - beertype[0].VerkoopPercentage)));
            workbook.Workbook.Worksheets["Costs per Product"].Cells[4, 12].Value = (beertype[1].voorraad + ((beertype[1].producedforNL + beertype[1].producedforBE + beertype[1].producedforSW) * (1 - beertype[1].VerkoopPercentage)));
            workbook.Workbook.Worksheets["Costs per Product"].Cells[5, 12].Value = (beertype[2].voorraad + ((beertype[2].producedforNL + beertype[2].producedforBE + beertype[2].producedforSW) * (1 - beertype[2].VerkoopPercentage)));
            workbook.Workbook.Worksheets["Costs per Product"].Cells[6, 12].Value = (beertype[3].voorraad + ((beertype[3].producedforNL + beertype[3].producedforBE + beertype[3].producedforSW) * (1 - beertype[3].VerkoopPercentage)));
            workbook.Workbook.Worksheets["Costs per Product"].Cells[7, 12].Value = (beertype[4].voorraad + ((beertype[4].producedforNL + beertype[4].producedforBE + beertype[4].producedforSW) * (1 - beertype[4].VerkoopPercentage)));
            workbook.Workbook.Worksheets["Costs per Product"].Cells[8, 12].Value = (beertype[5].voorraad + ((beertype[5].producedforNL + beertype[5].producedforBE + beertype[5].producedforSW) * (1 - beertype[5].VerkoopPercentage)));
            workbook.Workbook.Worksheets["Costs per Product"].Cells[9, 12].Value = (beertype[6].voorraad + ((beertype[6].producedforNL + beertype[6].producedforBE + beertype[6].producedforSW) * (1 - beertype[6].VerkoopPercentage)));
            workbook.Workbook.Worksheets["Costs per Product"].Cells[10, 12].Formula = "SUM(L3:L9)";

            //Kosten per biertie
            workbook.Workbook.Worksheets["Costs per Product"].Cells[3, 13].Value =(totalcostQ0/ totaalproduced)    ;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[4, 13].Value =(totalcostQ0 /totaalproduced)    ;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[5, 13].Value =(totalcostQ0 /totaalproduced)    ;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[6, 13].Value =(totalcostQ0 /totaalproduced)    ;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[7, 13].Value =(totalcostQ0 /totaalproduced)    ;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[8, 13].Value =(totalcostQ0 /totaalproduced)    ;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[9, 13].Value = (totalcostQ0 / totaalproduced);                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 
            workbook.Workbook.Worksheets["Costs per Product"].Cells[10,13].Value = "SOEPERTOL";
            //DOE EEN HELEBOEL EXCEL SHIT 


            //IT SHIZZLE
            workbook.Workbook.Worksheets["IT"].Cells[3, 3].Value = "IT Cost";
            workbook.Workbook.Worksheets["IT"].Cells[4, 3].Value = it.totalITCost;
            workbook.Workbook.Worksheets["IT"].Cells[3, 4].Value = "Ratio";
            workbook.Workbook.Worksheets["IT"].Cells[4, 4].Value = it.printRatioIT();




            saveAsExcelSheet();
        }

        public static void saveAsExcelSheet()
        {
            bool saved = false;
            while (!saved)
                try
                {
                    FileStream outputFile = new FileStream("results.xlsx", FileMode.Create);
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
