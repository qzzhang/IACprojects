using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Airport_Asset
{
    class Asset
    {

        private string AssetName;
        private ArrayList TargetNames;
        private ArrayList Options;
        private List<Target> TargetList;
        private Figure AssetMap;
        private Int16 TerminalId;
        private string TerminalName;
        private Int16 Scenarios;
        private Int16 ScenariosDetected;
        private Int16 ScenariosDetectedPct;
        private Int16 Situations;
        private Int16 SituationsDetected;
        private Int16 SituationsDetectedPct;
        private string TSSEOverview;

        public void setAssetMap(Figure mp)
        {
            AssetMap = mp;
        }

        public void setTerminalId(Int16 val)
        {
            TerminalId = val;
        }

        public void setTerminalName(string val)
        {
            TerminalName = val;
        }

        public Int16 getTerminalId()
        {
            return TerminalId;
        }
        public string getTerminalName()
        {
            return TerminalName;
        }
        public Figure getAssetMap()
        {
            return AssetMap;
        }

        public void setAssetName(string val)
        {
            AssetName = val;
        }

        public string getAssetName()
        {
            return AssetName;
        }

        public void setTargetNames(ArrayList val)
        {
            TargetNames = val;
        }

        public ArrayList getTargetNames()
        {
            return TargetNames;
        }

        public void setOptions(ArrayList val)
        {
            Options = val;
        }

        public ArrayList getOptions()
        {
            return Options;
        }

        public void setTargetList(List<Target> val)
        {
            TargetList = val;
        }

        public List<Target> getTargetList()
        {
            return TargetList;
        }

        public void setScenarios(Int16 val)
        {
            Scenarios = val;
        }

        public void setScenariosDetected(Int16 val)
        {
            ScenariosDetected = val;
        }

        public void setScenariosDetectedPct(Int16 val)
        {
            ScenariosDetectedPct = val;
        }

        public Int16 getScenarios()
        {
            return Scenarios;
        }

        public Int16 getScenariosDetected()
        {
            return ScenariosDetected;
        }

        public Int16 getScenariosDetectedPct()
        {
            double pct = 0;
            Int16 ret = 0;
            if (Scenarios > 0)
            {
                pct = (double)ScenariosDetected / Scenarios;
                pct = Math.Round(pct * 100.00, 0);
            }
            ret = (Int16)pct;
            return ret;
        }

        public void setSituations(Int16 val)
        {
            Situations = val;
        }

        public void setSituationsDetected(Int16 val)
        {
            SituationsDetected = val;
        }

        public void setSituationsDetectedPct(Int16 val)
        {
            SituationsDetectedPct = val;
        }

        public Int16 getSituations()
        {
            return Situations;
        }

        public Int16 getSituationsDetected()
        {
            return SituationsDetected;
        }

        public Int16 getSituationsDetectedPct()
        {
            double pct = 0;
            Int16 ret = 0;
            if (Situations > 0)
            {
                pct = (double)SituationsDetected / Situations;
                pct = Math.Round(pct * 100.00, 0);
            }
            ret = (Int16)pct;
            return ret;
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
