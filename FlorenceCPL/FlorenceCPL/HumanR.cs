using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FlorenceCPL
{
    class HumanR
    {
        public int lowprodgen;
        public int highprodgen;
        public int lowprodspec;
        public int highprodspec;
        public int lowadm;
        public int highadm;
        public double costlowprodgen;
        public double costhighprodgen;
        public double costlowprodspec;
        public double costhighprodspec;
        public double costlowadm;
        public double costhighadm;
        public double totalcost;


        public HumanR(string[] invoer, Marketing mt)
        {
            lowadm = (int)(mt.belban + mt.belcom + mt.frban + mt.frcom + mt.netban + mt.netcom + mt.swban + mt.swcom);
            

            totalcost = 0;
            costlowprodgen = double.Parse(invoer[0]);
            costhighprodgen =double.Parse(invoer[1]);
            costlowprodspec = double.Parse(invoer[2]);
            costhighprodspec = double.Parse(invoer[3]);
            costlowadm = double.Parse(invoer[4]);
            costhighadm = double.Parse(invoer[5]);
            highadm = int.Parse(invoer[6]);




        }

        public double totalcostPersonal()
        {


            return totalcost;
        }
    }
}
