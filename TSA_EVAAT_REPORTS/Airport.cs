using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Airport_Asset
{
    class Airport
    {
        private string AirportName;
        private string AirportAbbv;
        private DateTime AirportDate;
        private ArrayList AssetNames;
        private List<Asset> AssetList;
        private List<Asset> IEDAssetList;
        private List<Asset> VBIEDAssetList;
        private List<Terminal> TerminalList;
        private List<Target> Top5TargetList;
        private string FSDName;
        private string TSSEName;
        private string Duration;
        private string Address;
        private string City;
        private Figure LayoutMap;
        private string category;
        private ArrayList TerminalIds;
        private ArrayList TerminalNames;
        private string TSSEOverview;
        
        private string[, ,] RATScores; // socre[assettype, terminal, scoretype]
        private string[,] RASScores; // socre[assettype, scoretype]

        private string[,] SituationalChartData; // data[terminal, eventtype]

        public void setAirportName(string val)
        {
            AirportName = val;
        }

        public string getAirportName()
        {
            return AirportName;
        }

        public string getCategory()
        {
            return category;
        }

        public void setAirportAbbv(string val)
        {
            AirportAbbv = val;
        }

        public void setAirportAddr(string val)
        {
            Address = val;
        }

        public void setCity(string val)
        {
            City = val;
        }

        public void setFSDName(string val)
        {
            FSDName = val;
        }

        public void setTSSEName(string val)
        {
            TSSEName = val;
        }
        public void setCategory(string val)
        {
            category = val;
        }
        public void setDuration(string val)
        {
            Duration = val;
        }

        public void setRAScoresTerminal(string[, ,] val)
        {
            RATScores = val;
        }

        public void setRAScoresSupporting(string[,] val)
        {
            RASScores = val;
        }

        public string getAirportAbbv()
        {
            return AirportAbbv;
        }

        public string getAirportAddr()
        {
            return Address;
        }

        public string getAirportCity()
        {
            return City;
        }

        public void setAirportDate(DateTime dt)
        {
            AirportDate = dt;
        }

        public DateTime getAirportDate()
        {
            return AirportDate;
        }

        public void setAssetNames(ArrayList val)
        {
            AssetNames = val;
        }

        public ArrayList getAssetNames()
        {
            return AssetNames;
        }

        public void setIEDAssetList(List<Asset> val)
        {
            IEDAssetList = val;
        }

        public void setVBIEDAssetList(List<Asset> val)
        {
            VBIEDAssetList = val;
        }

        public void setAssetList(List<Asset> val)
        {
            AssetList = val;
        }

        public List<Asset> getAssetList()
        {
            return AssetList;
        }

        public List<Asset> getIEDAssetList()
        {
            return IEDAssetList;
        }

        public List<Asset> getVBIEDAssetList()
        {
            return VBIEDAssetList;
        }

        public string getFSDName()
        {
            return FSDName;
        }

        public string getDuration()
        {
            return Duration;
        }

        public string getTSSEName()
        {
            return TSSEName;
        }

        public void setLayoutMap(Figure mp)
        {
            LayoutMap = mp;
        }

        public Figure getLayoutMap()
        {
            return LayoutMap;
        }
        public string[, ,] getRASScoresTerminal()
        {
            return RATScores;
        }

        public string[,] getRAScoresSupporting()
        {
            return RASScores;
        }

        public void setTerminalIds(ArrayList val)
        {
            TerminalIds = val;
        }

        public void setTerminalNames(ArrayList val)
        {
            TerminalNames = val;
        }

        public ArrayList getTerminalIds()
        {
            return TerminalIds;
        }

        public ArrayList getTerminalNames()
        {
            return TerminalNames;
        }

        public void setTerminalList(List<Terminal> val)
        {
            TerminalList = val;
        }

        public List<Terminal> getTerminalList()
        {
            return TerminalList;

        }

        public void setSituationalChartData(string[,] val) {

            SituationalChartData = val;

        }

        public string[,] getSituationalChartData()
        {

            return SituationalChartData;

        }

        public void setTop5TargetList(List<Target> val)
        {
            Top5TargetList = val;

        }

        public List<Target> getTopTargetList()
        {
            return Top5TargetList;

        }

        public void setTSSEOverview(string val)
        {
            TSSEOverview = val;
        }

        public string getTSSEOverview()
        {
            return TSSEOverview;
        }

    }
}
