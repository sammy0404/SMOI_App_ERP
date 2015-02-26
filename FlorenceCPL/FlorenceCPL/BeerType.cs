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
        public int howMuchToBuyNL;
        public int howMuchToBuySW;
        public int howMuchToBuyBE;
        public int producedforNL;
        public int producedforSW;
        public int producedforBE;
        public string prducl;

        public BeerType(string[] invoer)
        {
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
            voorraad = 0;
        }
        
    }
}