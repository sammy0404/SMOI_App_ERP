using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FlorenceCPL
{
    class BeerType
    {
        public string beerName;
        public double malt;
        public double naturalWater;
        public double hops;
        public double munichMalt;
        public double caramelMalt;
        public double wheat;
        public double rye;
        public double maize;
        public double caneSugar;
        public double coriander;
        public double orangePeel;
        public int voorraad;
        public double transportPriceNL;
        public double transportPriceSW;
        public double transportPriceBE;
        public double transportPriceFR;
        public int howMuchToBuyNL;
        public int howMuchToBuySW;
        public int howMuchToBuyBE;
        public int howMuchToBuyFR;
        public int producedforNL;
        public int producedforSW;
        public int producedforBE;
        public int producedforFR;
        public string prducl;
        public double VerkoopPercentage;
        public double priceperBeer;
        public int totalproduced;
        public double clusterprijs;
        public List<RawMaterials> rw;
        public List<ProductieClusters> PCL;

        public BeerType(string[] invoer, List<RawMaterials> raw)
        {
            rw = raw;
            beerName = invoer[0];
            malt = double.Parse(invoer[1]);
            naturalWater = double.Parse(invoer[2]);
            hops = double.Parse(invoer[3]);
            munichMalt = double.Parse(invoer[4]);
            caramelMalt = double.Parse(invoer[5]);
            wheat = double.Parse(invoer[6]);
            rye = double.Parse(invoer[7]);
            maize = double.Parse(invoer[8]);
            caneSugar = double.Parse(invoer[9]);
            coriander = double.Parse(invoer[10]);
            orangePeel = double.Parse(invoer[11]);
            transportPriceNL = double.Parse(invoer[12]);
            transportPriceBE = double.Parse(invoer[13]);
            transportPriceSW = double.Parse(invoer[14]);
            transportPriceFR = double.Parse(invoer[15]);
            VerkoopPercentage = double.Parse(invoer[16]);
            priceperBeer = 0;
            clusterprijs = 0;
            voorraad = int.Parse(invoer[17]);
            if (malt != 0)
            {
                priceperBeer += (malt * rw[0].priceBuy);
            }
            if (naturalWater != 0)
            {
                priceperBeer += (naturalWater * rw[1].priceBuy);
            }
            if (hops != 0)
            {
                priceperBeer += (hops * rw[2].priceBuy);
            }
            if (munichMalt != 0)
            {
                priceperBeer += (munichMalt * rw[3].priceBuy);
            }
            if (caramelMalt != 0)
            {
                priceperBeer += (caramelMalt * rw[4].priceBuy);
            }
            if (wheat != 0)
            {
                priceperBeer += (wheat * rw[5].priceBuy);
            }
            if (rye != 0)
            {
                priceperBeer += (rye * rw[6].priceBuy);
            }
            if (maize != 0)
            {
                priceperBeer += (maize * rw[7].priceBuy);
            }
            if (caneSugar != 0)
            {
                priceperBeer += (caneSugar * rw[8].priceBuy);
            }
            if (coriander != 0)
            {
                priceperBeer += (coriander * rw[9].priceBuy);
            }
            if (orangePeel != 0)
            {
                priceperBeer += (orangePeel * rw[10].priceBuy);
            }


        }
        public void CalculatePCLprijs(List<ProductieClusters> pcl)
        {
            PCL = pcl;
            clusterprijs = ((PCL.Find(x => x.productieclusternaam == prducl).totalPersonalCost) + (PCL.Find(x=>x.productieclusternaam == prducl).leasePrice))/(PCL.Find(x=> x.productieclusternaam == prducl).typecapacity[0].Value) ;

        }
    }

}
