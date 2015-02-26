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
        public List<KeyValuePair<BeerType, int>> typecapacity;
        public ProductieClusters(string[] invoer)
        {
            productieclusternaam = invoer[0];
            leasePrice = int.Parse(invoer[3]);
            typecapacity = new List<KeyValuePair<BeerType, int>>();
        }
        public void addTypeCapacity(BeerType bt, string cap)
        {
            KeyValuePair<BeerType, int> kvp = new KeyValuePair<BeerType, int>(bt, int.Parse(cap));
            typecapacity.Add(kvp);
        }
    }

}
