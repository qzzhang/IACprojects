using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Airport_Asset {
    

    class Target {

        private string Target_Name;
        private string Justify;
        private ArrayList Observation;
        private ArrayList Options;
        private Figure TargetMap;
        private List<Figure> TargetFigure; 
        private DeterM Deter;
        private List<ThreatS> Threat;
        private string[] SituationalChartData; // data[eventstype]
        private string TerminalName;
        private string AssetName;

        public void setTarget_Name(string tn) {
            Target_Name = tn;
        }

        public string getTarget_Name() {
            return Target_Name;
        }

        public void setJustify(string jt) {
            Justify = jt;
        }

        public string getJustify() {
            return Justify;
        }

        public void setObservation(ArrayList val) {
            Observation = val;
        }

        public void setOptions(ArrayList val)
        {
            Options = val;
        }

        public ArrayList getObservation()
        {
            return Observation;
        }

        public ArrayList getOptions()
        {
            return Options;
        }

        public void setTargetMap(Figure mp)
        {
            TargetMap = mp;
        }

        public Figure getTargetMap() {
            return TargetMap;
        }

        public void setTargetFigure(List<Figure> fig) {
            TargetFigure = fig;
        }

        public List<Figure> getTargetFigure() {
            return TargetFigure;
        }

        public void setDeter(DeterM val) {
            Deter = val;
        }

        public DeterM getDeter() {
            return Deter;
        }

        public void setThreat(List<ThreatS> val) {
            Threat = val;
        }

        public List<ThreatS> getThreat() {
            return Threat;
        }
        public void setSituationalChartData(string[] val)
        {

            SituationalChartData = val;

        }

        public string[] getSituationalChartData()
        {

            return SituationalChartData;

        }

        public void setTerminalName(string val)
        {
            TerminalName = val;
        }

        public string getTerminalName()
        {
            return TerminalName ;
        }

        public void setAssetName(string val)
        {
            AssetName = val;
        }

        public string getAssetName()
        {
            return AssetName;
        }
    }
}
