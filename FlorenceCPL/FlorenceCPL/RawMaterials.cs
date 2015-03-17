using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FlorenceCPL
{

    class RawMaterials
    {
        public string materialName;
        public double priceBuy;
        public double priceSupply;
        public double voorraad;
        public int minOrderAmount;
        public double bulkPrice;
        public int amountBoughtThisQuarter;

        public RawMaterials(string[] invoer)
        {
            materialName = invoer[0];
            priceBuy = double.Parse(invoer[1]);
            priceSupply = double.Parse(invoer[2]);
            minOrderAmount = int.Parse(invoer[3]);
            bulkPrice = double.Parse(invoer[4]);
            voorraad = 0;
            amountBoughtThisQuarter = 0;
        }
    }
}
