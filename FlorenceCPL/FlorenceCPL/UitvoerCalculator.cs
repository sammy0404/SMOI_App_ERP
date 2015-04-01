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
        public static Marketing marketing;
        public int capital = 1000000;
        public int totaalproduced = 0;
        public double totalcostQ0 = 0;
        public double projectcost = 0;
        public UitvoerCalculator(List<RawMaterials> rw, List<BeerType> bt, List<ProductieClusters> pcl, IT newit, Marketing mark)
        {
            it = newit;
            marketing = mark;
            rawMaterials = rw;
            beertype = bt;
            PCL = pcl;
            DoWeHaveToProduce();
            ExtraCosts();
            foreach (BeerType b in bt)
            {
                totaalproduced += (b.producedforBE + b.producedforNL + b.producedforSW);
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
                projectcost += int.Parse(readedLine[i]);
            }


        }
        public void DoWeHaveToProduce()
        {
            foreach (BeerType bt in beertype)
            {
                if ((bt.howMuchToBuyNL + bt.howMuchToBuyBE + bt.howMuchToBuySW + bt.howMuchToBuyFR) > bt.voorraad)
                {
                    ProduceBeer(bt);
                }
                else
                {
                    bt.producedforNL = bt.howMuchToBuyNL;
                    bt.producedforBE = bt.howMuchToBuyBE;
                    bt.producedforSW = bt.howMuchToBuySW;
                    bt.producedforFR = bt.howMuchToBuyFR;
                    bt.voorraad -= (bt.howMuchToBuyNL + bt.howMuchToBuyBE + bt.howMuchToBuySW + bt.howMuchToBuyFR);

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
            while (bt.howMuchToBuyNL + bt.howMuchToBuyBE + bt.howMuchToBuySW  +bt.howMuchToBuyFR> bt.voorraad)
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
            bt.producedforFR = bt.howMuchToBuyFR;
            bt.voorraad -= (bt.howMuchToBuyNL + bt.howMuchToBuyBE + bt.howMuchToBuySW + bt.howMuchToBuyFR);

        }

        public void DoWeHaverawMaterials(RawMaterials rm, double btamount, BeerType bt)
        {
            int Productie = PCL.Find(x => x.productieclusternaam == bt.prducl).typecapacity.Find(y => y.Key.beerName == bt.beerName).Value;
            double nodig = Math.Ceiling((bt.howMuchToBuyNL + bt.howMuchToBuyBE + bt.howMuchToBuySW + bt.howMuchToBuyFR) * 1.0 / Productie);
            while (rm.voorraad < ((nodig * Productie) * btamount))
            {
                goBuyRawMaterial(rm, btamount, bt);
            }
        }
        public void goBuyRawMaterial(RawMaterials rm, double btamount, BeerType bt)
        {
            int Productie = PCL.Find(x => x.productieclusternaam == bt.prducl).typecapacity.Find(y => y.Key.beerName == bt.beerName).Value;
            double nodig = Math.Ceiling((bt.howMuchToBuyNL + bt.howMuchToBuyBE + bt.howMuchToBuySW + bt.howMuchToBuyFR) * 1.0 / Productie);
            while (rm.voorraad < ((nodig * Productie) * btamount))
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
            double MaxnrProduced = 0.0;
            foreach (BeerType bt in beertype)
            {
                bt.CalculatePCLprijs(PCL);
                bt.totalproduced = bt.producedforBE + bt.producedforNL + bt.producedforSW + bt.howMuchToBuyFR;
                MaxnrProduced += totaalproduced;
            }
            
           

            package = new ExcelPackage(new MemoryStream());
            workbook = package.Workbook.Worksheets.Add("Results");
          

            workbook.Workbook.Worksheets.Add("RawMaterialss");
            workbook.Workbook.Worksheets.Add("Costs per Product");
            workbook.Workbook.Worksheets.Add("IT");
            workbook.Workbook.Worksheets.Add("Marketing");
            


            workbook.Workbook.Worksheets["Costs per Product"].Cells.Style.Numberformat.Format = "0.00";



            // X-as van de excelsheet
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 2].Value = "Costs per Product";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 3].Value = "Produced for Netherlands";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 4].Value = "Produced for Belgium";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 5].Value = "Produced for Sweden";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 6].Value = "Produced for France";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 7].Value = "TransportCost to Netherlands";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 8].Value = "TransportCost to Belgium";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 9].Value = "TransportCost to Sweden";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 10].Value = "TransportCost to France";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 11].Value = "TransportCost from Netherlands";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 12].Value = "TransportCost from Belgium";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 13].Value = "TransportCost from Sweden";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 14].Value = "TransportCost from France";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 15].Value = "Opslagkosten na terugkomst";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 16].Value = "Gemiddeld per bier NL";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 17].Value = "Gemiddeld per bier BE";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 18].Value = "Gemiddeld per bier SW";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 19].Value = "Gemiddeld per bier FR";
            workbook.Workbook.Worksheets["Costs per Product"].Cells[2, 20].Value = "Voorraad";
           

            //Y-as van de excelsheet

            workbook.Workbook.Worksheets["Costs per Product"].Cells[3, 2].Value = beertype[0].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[4, 2].Value = beertype[1].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[5, 2].Value = beertype[2].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[6, 2].Value = beertype[3].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[7, 2].Value = beertype[4].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[8, 2].Value = beertype[5].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[9, 2].Value = beertype[6].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[10, 2].Value = beertype[7].beerName;
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 2].Value = "Totaal";

            //Produced aantal voor Nederland
            for (int i = 3; i <= 10; i++)
            {
                workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 3].Value = beertype[i - 3].producedforNL;
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 3].Formula = "SUM(C3:C10)";
            //Produced aantal voor Belgie
            for (int i = 3; i <= 10; i++)
            {
                workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 4].Value = beertype[i - 3].producedforBE;
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 4].Formula = "SUM(D3:D10)";

            //Produced aantal voor Sweden
            for (int i = 3; i <= 10; i++)
            {
                workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 5].Value = beertype[i - 3].producedforSW;
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 5].Formula = "SUM(E3:E10)";
            //Produced aantal voor Frankrijk
            for (int i = 3; i <= 10; i++)
            {
                workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 6].Value = beertype[i - 3].producedforFR;
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 6].Formula = "SUM(F3:F10)";

            //Transportkosten voor Nederland
            for (int i = 3; i <= 10; i++)
            {
                workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 7].Value = beertype[i-3].transportPriceNL * beertype[i-3].producedforNL;
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 7].Formula = "SUM(G3:G10)";
          
            //Transportkosten voor Belgie
            for (int i = 3; i <= 10; i++)
            {
                workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 8].Value = beertype[i - 3].transportPriceBE * beertype[i - 3].producedforBE;
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 8].Formula = "SUM(H3:H10)";
     
            //Transportkosten voor Sweden
            for (int i = 3; i <= 10; i++)
            {
                workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 9].Value = beertype[i - 3].transportPriceSW * beertype[i - 3].producedforSW;
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 9].Formula = "SUM(I3:I10)";
            //Transportkosten voor Frankrijk
            for (int i = 3; i <= 10; i++)
            {
                workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 10].Value = beertype[i - 3].transportPriceFR * beertype[i - 3].producedforFR;
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 10].Formula = "SUM(J3:J10)";
            //Transportkosten terug voor Nederland
            for (int i = 3; i <= 10; i++)
            {
                workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 11].Value = beertype[i-3].transportPriceNL * beertype[i-3].producedforNL * (1 - beertype[i-3].VerkoopPercentage);
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 11].Formula = "SUM(K3:K10)";
        
            //Transportkosten terug voor Belgie
            for (int i = 3; i <= 10; i++)
            {
                workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 12].Value = beertype[i - 3].transportPriceBE * beertype[i - 3].producedforBE * (1 - beertype[i - 3].VerkoopPercentage);
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 12].Formula = "SUM(L3:L10)";
          
            //Transportkosten terug voor Sweden
            for (int i = 3; i <= 10; i++)
            {
                workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 13].Value = beertype[i - 3].transportPriceSW * beertype[i - 3].producedforSW * (1 - beertype[i - 3].VerkoopPercentage);
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 13].Formula = "SUM(M3:M10)";
            //Transportkosten terug voor Frankrijk
            for (int i = 3; i <= 10; i++)
            {
                workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 14].Value = beertype[i - 3].transportPriceFR * beertype[i - 3].producedforFR * (1 - beertype[i - 3].VerkoopPercentage);
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 14].Formula = "SUM(N3:N10)";

            //Opslagkosten na terugkomst totaal
            for (int i = 3; i <= 10; i++)
            {
                workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 15].Value = (beertype[i-3].voorraad + ((beertype[i - 3].producedforNL + beertype[i - 3].producedforBE + beertype[i - 3].producedforSW + beertype[i - 3].producedforFR) * (1 - beertype[i - 3].VerkoopPercentage)));
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 15].Formula = "SUM(O3:O10)";
        
    


            //kosten per biertje Nederland
            for (int i = 3; i <= 10; i++)
            {
                if (beertype[i - 3].totalproduced != 0)
                {
                    if (beertype[i - 3].producedforNL != 0)
                    {
                        workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 16].Value = ((projectcost * (beertype[i - 3].totalproduced / MaxnrProduced)) / beertype[i - 3].totalproduced) + (beertype[i - 3].priceperBeer) + ((beertype[i - 3].transportPriceNL) * (2 - beertype[i - 3].VerkoopPercentage)) + (marketing.totaalNed / beertype[i - 3].producedforNL) + ((it.totalITCost * (beertype[i - 3].totalproduced / MaxnrProduced)) / beertype[i - 3].totalproduced) + (beertype[i - 3].clusterprijs) + (1 - beertype[i - 3].VerkoopPercentage);
                    }
                }
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 16].Value = "No info needed";
            //kosten per biertje Belgium
            for (int i = 3; i <= 10; i++)
            {
                if (beertype[i - 3].totalproduced != 0)
                {
                    if (beertype[i - 3].producedforBE != 0)
                    {
                        workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 17].Value = ((projectcost * (beertype[i - 3].totalproduced / MaxnrProduced)) / beertype[i - 3].totalproduced) + (beertype[i - 3].priceperBeer) + ((beertype[i - 3].transportPriceBE) * (2 - beertype[i - 3].VerkoopPercentage)) + (marketing.totaalBel / beertype[i - 3].producedforSW) + ((it.totalITCost * (beertype[i - 3].totalproduced / MaxnrProduced)) / beertype[i - 3].totalproduced) + (beertype[i - 3].clusterprijs) + (1 - beertype[i - 3].VerkoopPercentage);
                    }
                }
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 17].Value = "No info needed";
            //kosten per biertje Sweden
            for (int i = 3; i <= 10; i++)
            {
                if (beertype[i - 3].totalproduced != 0)
                {
                    if (beertype[i - 3].producedforSW != 0)
                    {
                        workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 18].Value = ((projectcost * (beertype[i - 3].totalproduced / MaxnrProduced)) / beertype[i - 3].totalproduced) + (beertype[i - 3].priceperBeer) + ((beertype[i - 3].transportPriceSW) * (2 - beertype[i - 3].VerkoopPercentage)) + (marketing.totaalSW / beertype[i - 3].producedforNL) + ((it.totalITCost * (beertype[i - 3].totalproduced / MaxnrProduced)) / beertype[i - 3].totalproduced) + (beertype[i - 3].clusterprijs) + (1 - beertype[i - 3].VerkoopPercentage);
                    }
                }
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 18].Value = "No info needed";
            //kosten per biertje Frankrijk
            for (int i = 3; i <= 10; i++)
            {
                if (beertype[i - 3].totalproduced != 0)
                {
                    if (beertype[i - 3].producedforFR != 0)
                    {
                        workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 19].Value = ((projectcost * (beertype[i - 3].totalproduced / MaxnrProduced)) / beertype[i - 3].totalproduced) + (beertype[i - 3].priceperBeer) + ((beertype[i - 3].transportPriceFR) * (2 - beertype[i - 3].VerkoopPercentage)) + (marketing.totaalFR / beertype[i - 3].producedforNL) + ((it.totalITCost * (beertype[i - 3].totalproduced / MaxnrProduced)) / beertype[i - 3].totalproduced) + (beertype[i - 3].clusterprijs) + (1 - beertype[i - 3].VerkoopPercentage);
                    }
                }
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 19].Value = "No info needed";

            //print voorraad
            for (int i = 3; i <= 10; i++)
            {
                workbook.Workbook.Worksheets["Costs per Product"].Cells[i, 20].Value = beertype[i - 3].voorraad;
            }
            workbook.Workbook.Worksheets["Costs per Product"].Cells[11, 20].Formula = "SUM(O3:O10)";

            //DOE EEN HELEBOEL EXCEL SHIT 


            //IT SHIZZLE
            workbook.Workbook.Worksheets["IT"].Cells[3, 3].Value = "IT Cost";
            workbook.Workbook.Worksheets["IT"].Cells[4, 3].Value = it.totalITCost;
            workbook.Workbook.Worksheets["IT"].Cells[3, 4].Value = "Ratio";
            workbook.Workbook.Worksheets["IT"].Cells[4, 4].Value = it.printRatioIT();

            //Marketing SHIZZLE
            workbook.Workbook.Worksheets["Marketing"].Cells[3, 3].Value = "Marketing Cost";
            workbook.Workbook.Worksheets["Marketing"].Cells[4, 3].Value = marketing.totaalMarketingCost;
            workbook.Workbook.Worksheets["Marketing"].Cells[3, 4].Value = "Cost Netherlands";
            workbook.Workbook.Worksheets["Marketing"].Cells[4, 4].Value = marketing.totaalNed;
            workbook.Workbook.Worksheets["Marketing"].Cells[3, 5].Value = "Cost Belgium";
            workbook.Workbook.Worksheets["Marketing"].Cells[4, 5].Value = marketing.totaalBel;
            workbook.Workbook.Worksheets["Marketing"].Cells[3, 6].Value = "Cost Sweden";
            workbook.Workbook.Worksheets["Marketing"].Cells[4, 6].Value = marketing.totaalSW;
            workbook.Workbook.Worksheets["Marketing"].Cells[3, 6].Value = "Cost France";
            workbook.Workbook.Worksheets["Marketing"].Cells[4, 6].Value = marketing.totaalFR;

            //RawMaterials SHIZZLE
            //X-axis
            for (int i = 3; i <= 13; i++)
            {
                workbook.Workbook.Worksheets["RawMaterialss"].Cells[i, 3].Value = rawMaterials[i-3].materialName;
            }
            

            workbook.Workbook.Worksheets["RawMaterialss"].Cells[2, 4].Value = "Price";
            workbook.Workbook.Worksheets["RawMaterialss"].Cells[2, 5].Value = "Voorraad";
            workbook.Workbook.Worksheets["RawMaterialss"].Cells[2, 6].Value = "Stock Cost";

            //price voor price ingekochte spullen
            for (int i = 3; i <= 13; i++)
            {
                workbook.Workbook.Worksheets["RawMaterialss"].Cells[i, 4].Value = rawMaterials[i-3].amountBoughtThisQuarter * rawMaterials[i-3].priceBuy;
            }
            workbook.Workbook.Worksheets["RawMaterialss"].Cells[18, 4].Formula = "SUM(D3:E17)";
            //price voor voorraad
            for (int i = 3; i <= 13; i++)
            {
                workbook.Workbook.Worksheets["RawMaterialss"].Cells[i, 5].Value = rawMaterials[i - 3].voorraad;
            }
            workbook.Workbook.Worksheets["RawMaterialss"].Cells[18, 5].Formula = "SUM(E3:E17)";
            //price voor stockCost
            for (int i = 3; i <= 13; i++)
            {
                workbook.Workbook.Worksheets["RawMaterialss"].Cells[i, 6].Value = rawMaterials[i - 3].voorraad * rawMaterials[i - 3].priceSupply;
            }

            workbook.Workbook.Worksheets["RawMaterialss"].Cells[18, 6].Formula = "SUM(F3:F17)";

            workbook.Cells[2, 2].Value = "Omzet";
            for (int i = 3; i <= 9; i++)
            {
                workbook.Workbook.Worksheets["Results"].Cells[i, 3].Value = beertype[i - 3].beerName;
            }
            workbook.Workbook.Worksheets["Results"].Cells[2, 4].Value = "Netherlands";
            workbook.Workbook.Worksheets["Results"].Cells[2, 5].Value = "Belgium";
            workbook.Workbook.Worksheets["Results"].Cells[2, 6].Value = "Sweden";
            workbook.Workbook.Worksheets["Results"].Cells[2, 7].Value = "France";
            workbook.Workbook.Worksheets["Results"].Cells[2, 8].Value = "He Transportkosten";
            workbook.Workbook.Worksheets["Results"].Cells[2, 9].Value = "Te Transportkosten";
            workbook.Workbook.Worksheets["Results"].Cells[2, 10].Value = "LeasePrijsPCL";
            workbook.Workbook.Worksheets["Results"].Cells[2, 11].Value = "MaintenancePCL";

            workbook.Workbook.Worksheets["Results"].Cells[3, 4].Value = beertype[0].producedforNL * beertype[0].VerkoopPercentage * 10;
            workbook.Workbook.Worksheets["Results"].Cells[4, 4].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[5, 4].Value = beertype[2].producedforNL * beertype[2].VerkoopPercentage * 10;
            workbook.Workbook.Worksheets["Results"].Cells[6, 4].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[7, 4].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[8, 4].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[9, 4].Value = beertype[6].producedforNL * beertype[6].VerkoopPercentage * 10;

            workbook.Workbook.Worksheets["Results"].Cells[3, 5].Value = beertype[0].producedforBE * beertype[0].VerkoopPercentage * 10;
            workbook.Workbook.Worksheets["Results"].Cells[4, 5].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[5, 5].Value = beertype[2].producedforBE * beertype[2].VerkoopPercentage * 10;
            workbook.Workbook.Worksheets["Results"].Cells[6, 5].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[7, 5].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[8, 5].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[9, 5].Value = beertype[6].producedforBE * beertype[6].VerkoopPercentage * 10;

            workbook.Workbook.Worksheets["Results"].Cells[3, 6].Value = beertype[0].producedforSW * beertype[0].VerkoopPercentage * 10;
            workbook.Workbook.Worksheets["Results"].Cells[4, 6].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[5, 6].Value = beertype[2].producedforSW * beertype[2].VerkoopPercentage * 10;
            workbook.Workbook.Worksheets["Results"].Cells[6, 6].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[7, 6].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[8, 6].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[9, 6].Value = beertype[6].producedforSW * beertype[6].VerkoopPercentage * 10;

            workbook.Workbook.Worksheets["Results"].Cells[3, 7].Value = beertype[0].producedforFR * beertype[0].VerkoopPercentage * 10;
            workbook.Workbook.Worksheets["Results"].Cells[4, 7].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[5, 7].Value = beertype[2].producedforFR * beertype[2].VerkoopPercentage * 10;
            workbook.Workbook.Worksheets["Results"].Cells[6, 7].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[7, 7].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[8, 7].Value = "0";
            workbook.Workbook.Worksheets["Results"].Cells[9, 7].Value = beertype[6].producedforFR * beertype[2].VerkoopPercentage * 10;

            workbook.Workbook.Worksheets["Results"].Cells[3, 8].Formula = "=SOM('Costs per Product'!G11:J11)";
            workbook.Workbook.Worksheets["Results"].Cells[3, 9].Formula = "=SOM('Costs per Product'!K11:N11)";
            int totaalPCL = 0;
            foreach (ProductieClusters pl in PCL)
            {
                totaalPCL += pl.timesneeded* pl.leasePrice;
            }
            workbook.Workbook.Worksheets["Results"].Cells[3, 10].Value = (totaalPCL / 125) * 100;
            workbook.Workbook.Worksheets["Results"].Cells[3, 11].Value = (totaalPCL / 125) * 25;

            workbook.Workbook.Worksheets["Results"].Cells[11, 4].Value = "Totaal Omzet";
            workbook.Workbook.Worksheets["Results"].Cells[11, 5].Formula = "SUM(D3:D9) + SUM(E3:E9) +SUM(F3:F9)+ SUM(G3:G9)";





















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
