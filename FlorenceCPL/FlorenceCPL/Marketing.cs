using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FlorenceCPL
{
    class Marketing
    {
        public double netban;
        public double belban;
        public double swban;
        public double frban;
        public double netcom;
        public double belcom;
        public double swcom;
        public double frcom;
        public double costnetban;
        public double costbelban;
        public double costswban;
        public double costfrban;
        public double costnetcom;
        public double costbelcom;
        public double costswcom;
        public double costfrcom;
        public double totaalNed;
        public double totaalBel;
        public double totaalSW;
        public double totaalFR;
        public double totaalMarketingCost;
        public Marketing(string [] invoer)
        {
            netban = double.Parse(invoer[0]);
            belban = double.Parse(invoer[1]);
            swban = double.Parse(invoer[2]);
            frban = double.Parse(invoer[3]);
            netcom = double.Parse(invoer[4]);
            belcom = double.Parse(invoer[5]);
            swcom = double.Parse(invoer[6]);
            frcom = double.Parse(invoer[7]);
            costnetban = double.Parse(invoer[8]);
            costbelban = double.Parse(invoer[9]);
            costswban = double.Parse(invoer[10]);
            costfrban = double.Parse(invoer[11]);
            costnetcom = double.Parse(invoer[12]);
            costbelcom = double.Parse(invoer[13]);
            costswcom = double.Parse(invoer[14]);
            costfrcom = double.Parse(invoer[15]);
            totaalNed += (costnetban * netban) + (costnetcom * netcom);
            totaalSW += (costswban * swban) + (costswcom * swcom);
            totaalBel += (costbelban * belban) + (costbelcom * belcom);
            totaalFR += (costfrban * frban) + (costfrcom * frcom);
            totaalMarketingCost = (totaalBel + totaalNed + totaalSW + totaalFR);
        }
    }
}
