using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Airport_Asset {
    class ThreatS {
        private int Threat_Id;
        private int Minute_Id;
        private int Detect_Id;
        private string ThreatDate;

        public void setThreat_Id(int val)
        {
            Threat_Id = val;
        }

        public int getThreat_Id()
        {
            return Threat_Id;
        }

        public void setMinute_Id(int val)
        {
            Minute_Id = val;
        }

        public int getMinute_Id()
        {
            return Minute_Id;
        }

        public void setDetect_Id(int val)
        {
            Detect_Id = val;
        }

        public int getDetect_Id()
        {
            return Detect_Id;
        }

        public void setThreatDate(string val) {
            ThreatDate = val;
        }

        public string getThreatDate() {
            return ThreatDate;
        }


    }
}
