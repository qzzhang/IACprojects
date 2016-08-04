using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Airport_Asset {
    class DeterM {

        private int numTeam;
        private int numCCTV;    // 1  - CCTV
        private int numSec;     // 2  - Security
        private int numLight;   // 3  - Lighting
        private int numEmp;     // 4  - Airport Employee
        private int numWall;    // 5  - Blast Walls
        private int numGlass;   // 6  - Glass Mitigation
        private int numRand;    // 7  - Random Security
        private int numBarr;    // 8  - Barriers
        private int numPerm;    // 9  - Parimeter
        private int numSign;    // 10 - Signage
        private int numPublic;  // 11 - General Public


        public void setNumTeam(int val) {
            numTeam = val;
        }


        public int getNumTeam() {
            return numTeam;
        }


        public void setNumCCTV(int val) {
            numCCTV = val;
        }


        public int getNumCCTV() {
            double pct = 0;
            if (numTeam > 0)
            {
                pct = (double)numCCTV / numTeam;
                pct = Math.Round(pct * 100.00, 0);
            }
            return (int) pct;
           // return numCCTV;
        }


        public void setNumSec(int val) {
            numSec = val;
        }


        public int getNumSec() {
            double pct = 0;
            if (numTeam > 0)
            {
                pct = (double)numSec / numTeam;
                pct = Math.Round(pct * 100.00, 0);
            }
            return (int)pct;
            //return numSec;
        }


        public void setNumLight(int val) {
            numLight = val;
        }


        public int getNumLight() {
            double pct = 0;
            if (numTeam > 0)
            {
                pct = (double)numLight / numTeam;
                pct = Math.Round(pct * 100.00, 0);
            }
            return (int)pct;
            //return numLight;
        }


        public void setNumEmp(int val) {
            numEmp = val;
        }


        public int getNumEmp() {
            double pct = 0;
            if (numTeam > 0)
            {
                pct = (double)numEmp / numTeam;
                pct = Math.Round(pct * 100.00, 0);
            }
            return (int)pct;
            //return numEmp;
        }


        public void setNumWall(int val) {
            numWall = val;
        }


        public int getNumWall() {
            double pct = 0;
            if (numTeam > 0)
            {
                pct = (double)numWall / numTeam;
                pct = Math.Round(pct * 100.00, 0);
            }
            return (int)pct;
            //return numWall;
        }


        public void setNumGlass(int val) {
            numGlass = val;
        }


        public int getNumGlass() {
            double pct = 0;
            if (numTeam > 0)
            {
                pct = (double)numGlass / numTeam;
                pct = Math.Round(pct * 100.00, 0);
            }
            return (int)pct;
            //return numGlass;
        }


        public void setNumRand(int val) {
            numRand = val;
        }


        public int getNumRand() {
            double pct = 0;
            if (numTeam > 0)
            {
                pct = (double)numRand / numTeam;
                pct = Math.Round(pct * 100.00, 0);
            }
            return (int)pct;
            //return numRand;
        }


        public void setNumBarr(int val) {
            numBarr = val;
        }


        public int getNumBarr() {
            double pct = 0;
            if (numTeam > 0)
            {
                pct = (double)numBarr / numTeam;
                pct = Math.Round(pct * 100.00, 0);
            }
            return (int)pct;
            //return numBarr;
        }


        public void setNumPerm(int val) {
            numPerm = val;
        }


        public int getNumPerm() {
            double pct = 0;
            if (numTeam > 0)
            {
                pct = (double)numPerm / numTeam;
                pct = Math.Round(pct * 100.00, 0);
            }
            return (int)pct;
            //return numPerm;
        }


        public void setNumSign(int val) {
            numSign = val;
        }


        public int getNumSign() {
            double pct = 0;
            if (numTeam > 0)
            {
                pct = (double)numSign / numTeam;
                pct = Math.Round(pct * 100.00, 0);
            }
            return (int)pct;
            //return numSign;
        }


        public void setNumPublic(int val) {
            numPublic = val;
        }


        public int getNumPublic() {
            double pct = 0;
            if (numTeam > 0)
            {
                pct = (double)numPublic / numTeam;
                pct = Math.Round(pct * 100.00, 0);
            }
            return (int)pct;
            //return numPublic;
        }


    }
}
