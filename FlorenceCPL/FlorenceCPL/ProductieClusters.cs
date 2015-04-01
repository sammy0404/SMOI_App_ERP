using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FlorenceCPL
{
    class ProductieClusters
    {
        public string productieclusternaam;
        public int leasePrice;
        public int totalPersonalCost;
        public int timesneeded;
        public int highprodgen;
        public int lowprodgen;
        public int highprodspec;
        public int lowprodspec;

        public int amountused;
        public List<KeyValuePair<BeerType, int>> typecapacity;
        public ProductieClusters(string[] invoer)
        {
            productieclusternaam = invoer[0];
            leasePrice = (int)(((int.Parse(invoer[3]))*1.25)*1.1);
     //       totalPersonalCost = int.Parse(invoer[4]);
            typecapacity = new List<KeyValuePair<BeerType, int>>();
        }
        public void addTypeCapacity(BeerType bt, string cap)
        {
            KeyValuePair<BeerType, int> kvp = new KeyValuePair<BeerType, int>(bt, int.Parse(cap));
            typecapacity.Add(kvp);
        }
    }

}
