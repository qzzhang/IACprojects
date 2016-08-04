using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Airport_Asset {
    class Config {

        public Config() {
        }


        public string getValue(string var) {

            string rtnval = "";

            try {
                StreamReader sr = new StreamReader(Globals.config_file);
                string line;
                while ((line = sr.ReadLine()) != null) {
                    string[] words = Regex.Split(line," = ");
                    if (words[0] == var) {
                        rtnval = words[1];
                    }
                }

            } catch (Exception ex) {
                Console.WriteLine("File could not read: " + ex.Message);
            }
            return rtnval;
        }



    }
}
