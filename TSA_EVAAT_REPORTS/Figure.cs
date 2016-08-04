using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Airport_Asset {


    class Figure {

        private byte[] tfig;
        private string fileName;

        public void setFigure(byte[] fg) {
            tfig = fg;
        }

        public byte[] getFigure() {
            return tfig;
        }

        public void setFileName(string val) {
            fileName = val;
        }

        public string getFileName() {
            return fileName;
        }

    }
}
