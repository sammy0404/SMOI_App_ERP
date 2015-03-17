using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FlorenceCPL
{
    class IT
    {
        public int admPersonal;
        public int aantalClients;
        public int aantalServers;
        public double costClient;
        public double costServer;
        public int previousClients;
        public int previousServers;
        public int tellerclient;
        public double totalITCost;
        public IT(string[] invoer)
        {
            admPersonal = int.Parse(invoer[0]);
            aantalClients = int.Parse(invoer[1]);
            costClient = double.Parse(invoer[2]);
            costServer = double.Parse(invoer[3]);
            previousClients = int.Parse(invoer[4]);
            previousServers = int.Parse(invoer[5]);
            totalITCost = 0;
            aantalServers = 0;
            tellerclient = aantalClients;
            CalculateCost();
        }
        public void CalculateCost()
        {
            if (previousClients - aantalClients < 0)
            {
                totalITCost += ((aantalClients - previousClients)*costClient); 
            }
            
            while (tellerclient > 0)
            {
                tellerclient -= 20;
                aantalServers += 1;
            }
            if (previousServers - aantalServers < 0)
            {
                totalITCost += ((aantalServers - previousServers) * costServer);
            }
        }
        public string printRatioIT()
        {
            double ratio = (20.0 + tellerclient) / 20.0;
            return ratio.ToString();
        }
    }
}
