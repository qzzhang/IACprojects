using System;
using System.Data;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;

namespace Airport_Asset
{
    class DBManager
    {

        static OleDbConnection conn1;	// Used to populate Airport Object
        static OleDbConnection conn2;	// Used to populate Asset Object

        string g_OraConStr;
        static long VisitId;
        static long AirportId;
        public DBManager()
        {
            Config cfg = new Config();

            g_OraConStr = "Provider=OraOLEDB.Oracle;" +
                          "Data Source=(DESCRIPTION = (ADDRESS_LIST = (ADDRESS = " +
                          "(PROTOCOL = TCP)(HOST = " + cfg.getValue("ohost") + ")(PORT = 1521)) )" +
                          "(CONNECT_DATA = (SERVER = DEDICATED) (SERVICE_NAME = " + cfg.getValue("orasource") + ")));" +
                          "User ID=" + cfg.getValue("userid") + ";" +
                          "Password=" + CryptorEngine.Decrypt((String)cfg.getValue("passwd"), true) + ";";

            

        }

        // Open the Database Manager
        public void openDB()
        {
            
            // Used to populate Airport Object
            conn1 = new OleDbConnection(g_OraConStr);

            // Used to populate Asset Object
            conn2 = new OleDbConnection(g_OraConStr);

            conn1.Close();
            conn2.Close();

            conn1.Open();
            conn2.Open();
        }

        // Close the Database Manager
        public void closeDB()
        {
            conn1.Close();
            conn2.Close();
        }


        public byte[] fillBLOB()
        {
            return null;
        }


        /* *****************************************
         * Populate the Airport Object
         *  - Get Name and Abbrev
         *  - Set the Airport Consider Date
         *  - Populate the Asset Name List
         *  - Populate Asset List
         * 
         * *****************************************/
        public Airport fillAirport(long vid)
        {

            VisitId = vid;
            Airport airp = new Airport();
            long airportId = 0;
            try
            {
                /////////////////////////////////////////
                // Populate Name, Abbrev, and Date
                /////////////////////////////////////////
                OleDbCommand srcCmd;
                OleDbDataReader srcRead;
                string strSQL = "Select a.ID, a.NAME, a.ABBREVIATION, a.STREET, s.state, a.CITY, a.ZIP, c.category, v.tsse_overview" +
                                " From AIRPORTS a, VISITS v, STATES s, CATEGORIES c" +
                                " where a.id = v.facilityid and visitid=" + vid +
                                "   and s.id = a.stateid" +
                                "   and c.categoryid = nvl(a.categoryid,2)";

                srcCmd = new OleDbCommand(strSQL, conn1);
                srcCmd.CommandType = CommandType.Text;

                srcRead = srcCmd.ExecuteReader();

                while (srcRead.Read())
                {
                    airp.setAirportName(srcRead["NAME"].ToString());

                    airp.setAirportAbbv(srcRead["ABBREVIATION"].ToString());
                    airp.setAirportAddr(srcRead["STREET"].ToString() + ", " + srcRead["CITY"].ToString() + " " + srcRead["STATE"].ToString() + " "
                                        + srcRead["ZIP"].ToString());
                    airp.setCity(srcRead["CITY"].ToString());
                    airportId = Int32.Parse(srcRead["ID"].ToString());
                    airp.setCategory(srcRead["CATEGORY"].ToString());
                    airp.setTSSEOverview(srcRead["TSSE_OVERVIEW"].ToString());

                    AirportId = airportId;
                }
                srcRead.Close();
                srcCmd.Dispose();

                ////////////////////////////////////////////////
                // TEMPORARY SETTING THE AIRPORT CONSIDER DATE
                ///////////////////////////////////////////////
                airp.setAirportDate(DateTime.Now);


                ///////////////////////////////////
                // Populate Terminals List
                ///////////////////////////////////
                ArrayList idLst = new ArrayList();
                ArrayList nameLst = new ArrayList();

                strSQL = "SELECT id, terminal" +
                         " FROM terminals" +
                         " WHERE AIRPORTID=" + airportId +
                         " ORDER BY terminal";

                srcCmd = new OleDbCommand(strSQL, conn1);
                srcCmd.CommandType = CommandType.Text;
                srcRead = srcCmd.ExecuteReader();

                while (srcRead.Read())
                {
                    idLst.Add(srcRead["ID"].ToString());
                    nameLst.Add(srcRead["TERMINAL"].ToString());

                }
                srcRead.Close();
                srcCmd.Dispose();
                airp.setTerminalIds(idLst);
                airp.setTerminalNames(nameLst);

                ///////////////////////////////////
                // Populate Asset Name List
                ///////////////////////////////////
                ArrayList astLst = new ArrayList();
                strSQL = "SELECT ASSET_NAME" +
                         " FROM vw_asset_overview" +
                         " WHERE VISITID=" + vid +
                         " ORDER BY ASSET_NAME";

                srcCmd = new OleDbCommand(strSQL, conn1);
                srcCmd.CommandType = CommandType.Text;
                srcRead = srcCmd.ExecuteReader();
                while (srcRead.Read())
                {
                    astLst.Add(srcRead["ASSET_NAME"].ToString());
                }
                srcRead.Close();
                srcCmd.Dispose();
                airp.setAssetNames(astLst);
                ///////////////////////////
                // FILL UP THE ASSET LIST
                //////////////////////////
                airp.setAssetList(fillAsset(vid));

                ///////////////////////////
                // FILL UP THE IED ASSET LIST
                //////////////////////////
                airp.setIEDAssetList(fillIEDAsset(vid));

                airp.setVBIEDAssetList(fillVBIEDAsset(vid));

                /////////////////////////////////////////
                // Populate FSDName, TSS-E Name
                /////////////////////////////////////////
                strSQL = "Select u.first_name, u.last_name " +
                          " From VISITS v, EVAT_USERS u" +
                         " where visitid=:vid " +
                         "   and u.EVAT_USER_NAME = v.FSD";
                srcCmd = new OleDbCommand(strSQL, conn1);
                srcCmd.CommandType = CommandType.Text;

                srcCmd.Parameters.AddWithValue("vid", vid);

                srcRead = srcCmd.ExecuteReader();

                while (srcRead.Read())
                {
                    airp.setFSDName(srcRead["FIRST_NAME"].ToString() + " " + srcRead["LAST_NAME"].ToString());
                }
                srcRead.Close();
                srcCmd.Dispose();

                strSQL = "Select u.first_name, u.last_name " +
                          " From VISITS v, EVAT_USERS u" +
                         " where visitid=:vid " +
                         "   and u.EVAT_USER_NAME = v.TSSE";
                srcCmd = new OleDbCommand(strSQL, conn1);
                srcCmd.CommandType = CommandType.Text;

                srcCmd.Parameters.AddWithValue("vid", vid);

                srcRead = srcCmd.ExecuteReader();

                while (srcRead.Read())
                {
                    airp.setTSSEName(srcRead["FIRST_NAME"].ToString() + " " + srcRead["LAST_NAME"].ToString());
                }
                srcRead.Close();
                srcCmd.Dispose();

                /////////////////////////////////////////
                // Populate Duration
                /////////////////////////////////////////
                strSQL = "Select to_char(DataentryStarted, 'Mon YYYY') from_dt, to_char(DataentryCompleted,'Mon YYYY') to_dt " +
                          " From SITEVISITINFO" +
                         " where visitid=:vid " +
                         "   and dataentryStarted is not null and dataentrycompleted is not null ";
                srcCmd = new OleDbCommand(strSQL, conn1);
                srcCmd.CommandType = CommandType.Text;

                srcCmd.Parameters.AddWithValue("vid", vid);

                srcRead = srcCmd.ExecuteReader();

                while (srcRead.Read())
                {
                    airp.setDuration(srcRead["FROM_DT"].ToString() + " to " + srcRead["TO_DT"].ToString());
                    airp.setAirportDate(DateTime.Parse(srcRead["TO_DT"].ToString()));
                }
                srcRead.Close();
                srcCmd.Dispose();

                Figure fg = fillLayoutFigure(airportId);

                airp.setLayoutMap(fg);

                strSQL = @"begin
                          delete from EVAAT_RPT_SCORES where airport_id = ?;
                          PKG_AT_SCORING.GET_AT_SCORE(0, ?, 0, -1, 0, 'VBIED');
                          PKG_AT_SCORING.GET_AT_SCORE(0, ?, 0, -1, 0, 'IED');
                          null;
                        end;";

                srcCmd = new OleDbCommand(strSQL, conn1);
                srcCmd.CommandType = CommandType.Text;
                srcCmd.Parameters.AddWithValue("airpid", airportId);
                srcCmd.Parameters.AddWithValue("vid", vid);
                srcCmd.Parameters.AddWithValue("vid", vid);
                srcCmd.ExecuteNonQuery();//This query takes quite a long time to finish
                strSQL = @"SELECT nvl(TERMINAL_ID,0) TERMINAL_ID, SCORE_TYPE, round(least( greatest(nvl(ied_score,0),1),  100), 0) IED_SCORE, 
                                                   round(least( greatest(nvl(vbied_score,0),1),100), 0) VBIED_SCORE
                         FROM EVAAT_RPT_SCORES 
                        WHERE airport_id  =  " + airportId +
                         @" and nvl(TERMINAL_ID,0) != 0
                        and score_type in (102, 103, 104, 105, 106, 107)
                        order by score_type";

                srcCmd = new OleDbCommand(strSQL, conn1);
                srcCmd.CommandType = CommandType.Text;

                int i = 0;
                int j = -1;

                //string terminalId = "";
                string[] term_arr = new string[10]; //used to store the terminal ids
                string temp_term = "";
                string[, ,] scores = new string[6, 10, 2];

                srcRead = srcCmd.ExecuteReader();

                while (srcRead.Read())
                {//Here is where the "Index was outside the bounds of the array." error message was thrown from, because the scores[6,10,2] reached to scores[0,10,2]
                    temp_term = srcRead["TERMINAL_ID"].ToString();
                    if (Array.Exists(term_arr, element => element == temp_term))
                    {
                        j = Array.FindIndex(term_arr, element => element == temp_term);
                        switch (srcRead["SCORE_TYPE"].ToString())
                        {
                            case "102":
                                i = 0;
                                break;
                            case "103":
                                i = 1;
                                break;
                            case "104":
                                i = 2;
                                break;
                            case "105":
                                i = 3;
                                break;
                            case "106":
                                i = 4;
                                break;
                            case "107":
                                i = 5;
                                break;
                            default:
                                break;
                         }
                    }
                    else
                    //if (terminalId != srcRead["TERMINAL_ID"].ToString())
                    {
                        j++;
                        i = 0;
                        //terminalId = srcRead["TERMINAL_ID"].ToString();
                        term_arr[j] = temp_term;
                    }
                    scores[i, j, 0] = srcRead["IED_SCORE"].ToString();
                    scores[i, j, 1] = srcRead["VBIED_SCORE"].ToString();
                    //i++;
                }
                srcRead.Close();

                airp.setRAScoresTerminal(scores);

                string[,] scores2 = new string[6, 2];

                strSQL = @"SELECT nvl(TERMINAL_ID,0) TERMINAL_ID, SCORE_TYPE, round(least( greatest(nvl(ied_score,0),1),  100), 0) IED_SCORE, 
                                                   round(least( greatest(nvl(vbied_score,0),1),100), 0) VBIED_SCORE
                         FROM EVAAT_RPT_SCORES 
                        WHERE airport_id  =  " + airportId +
                         @" and nvl(TERMINAL_ID,0) = 0 
                        and score_type in (108, 109, 110, 111, 112)
                        order by score_type";

                srcCmd = new OleDbCommand(strSQL, conn1);
                srcCmd.CommandType = CommandType.Text;

                srcRead = srcCmd.ExecuteReader();

                i = 0;
                while (srcRead.Read())
                {
                    scores2[i, 0] = srcRead["IED_SCORE"].ToString();
                    scores2[i, 1] = srcRead["VBIED_SCORE"].ToString();
                    i++;

                }
                srcRead.Close();

                airp.setRAScoresSupporting(scores2);

                string[,] sitNums = new string[10, 4];

                strSQL = @"select b.terminalid, count(1) total,
                            count(decode(detect_hour_id,1, 1, null)) high,
                            count(decode(detect_hour_id,2, 1, null)) medium,
                            count(decode(detect_hour_id,3, 1, null)) low
                        from tbl_detect_behave a, tbl_target_area b
                        where a.target_area_id = b.target_area_id
                        and b.visitid = :visitid
                        group by b.terminalid";

                srcCmd = new OleDbCommand(strSQL, conn1);
                srcCmd.CommandType = CommandType.Text;
                srcCmd.Parameters.AddWithValue("visitid", vid);

                try
                {
                    srcRead = srcCmd.ExecuteReader();

                    string termid = "";
                    string prev_termid = "";
                    i = 0;
                    while (srcRead.Read())
                    {
                        termid = srcRead["TERMINALID"].ToString();
                        if (i > 0 && termid != prev_termid) i++;

                        sitNums[i, 0] = srcRead["TOTAL"].ToString();
                        sitNums[i, 1] = srcRead["LOW"].ToString();
                        sitNums[i, 2] = srcRead["MEDIUM"].ToString();
                        sitNums[i, 3] = srcRead["HIGH"].ToString();

                        prev_termid = termid;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: db query: " + ex.Message);
                }
                srcRead.Close();

                airp.setSituationalChartData(sitNums);

                //set top 5 targets...Query modified by QZ on July 1, 2014

                strSQL = @"select * from (select distinct b.id terminalid, b.terminal, a.target_area_id, c.asset_id, c.description asset_name,
                              a.ta_name target_name, d.attractiveness
                        from tbl_target_area a, terminals b, lu_asset c, 
                             (select visitid, assetid, target_area_id, attractiveness
                                from (select visitid, assetid, target_area_id, attractiveness
                                        from tbl_attractiveness
                                       where visitid = :vid 
                                    order by attractiveness desc)   
                              )d 
                        where a.visitid = d.visitid
                          and b.id = a.terminalid
                          and c.asset_id = a.asset_id
                          and a.target_area_id = d.target_area_id
                        order by d.attractiveness desc) where rownum<6";

                srcCmd = new OleDbCommand(strSQL, conn1);
                srcCmd.CommandType = CommandType.Text;
                srcCmd.Parameters.AddWithValue("vid", vid);

                srcRead = srcCmd.ExecuteReader();
                List<Target> tgtList = new List<Target>();
                i = 0;
                while (srcRead.Read())
                {

                    Target trg = new Target();
                    string ast = srcRead["ASSET_NAME"].ToString();                       // Get Asset Name
                    string tan = srcRead["TARGET_NAME"].ToString();                     // Get the Target Name
                    long tgid = long.Parse(srcRead["TARGET_AREA_ID"].ToString());   // Get the TARGET_AREA_ID
                    long tid = long.Parse(srcRead["TERMINALID"].ToString());
                    trg.setTarget_Name(tan);                          // Set the Target Name
                    trg.setOptions(fillOptions(tgid));                     // get and Set the ArrayList of Observations
                    trg.setTerminalName(srcRead["TERMINAL"].ToString());
                    trg.setAssetName(srcRead["ASSET_NAME"].ToString());

                    string[] sitNumsT = new string[4];

                    strSQL = @"select count(1) total,
                            count(decode(detect_hour_id,1, 1, null)) high,
                            count(decode(detect_hour_id,2, 1, null)) medium,
                            count(decode(detect_hour_id,3, 1, null)) low
                        from tbl_detect_behave a, tbl_target_area b
                        where a.target_area_id = b.target_area_id
                        and b.visitid = :visitid
                        and b.target_area_id = :tgtid";

                    srcCmd = new OleDbCommand(strSQL, conn1);
                    srcCmd.CommandType = CommandType.Text;
                    srcCmd.Parameters.AddWithValue("visitid", vid);
                    srcCmd.Parameters.AddWithValue("tgtid", tgid);

                    OleDbDataReader srcRead2;
                    srcRead2 = srcCmd.ExecuteReader();
                    i = 0;
                    while (srcRead2.Read())
                    {
                        sitNumsT[0] = srcRead2["TOTAL"].ToString();
                        sitNumsT[1] = srcRead2["LOW"].ToString();
                        sitNumsT[2] = srcRead2["MEDIUM"].ToString();
                        sitNumsT[3] = srcRead2["HIGH"].ToString();
                    }
                    srcRead2.Close();

                    trg.setSituationalChartData(sitNumsT);

                    tgtList.Add(trg);

                }
                srcRead.Close();

                airp.setTop5TargetList(tgtList);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            return airp;
        }




        /* *****************************************
         *  Fill the Asset List
         * 
         *  For each Asset
         *  * Get the Asset Name
         *  * Get the Terminal Id
         *  * Get the Asset Id
         *  * Get the List of Target Names
         *  * Get the Asset Option List
         *  * Fill the Asset's Target List
         * 
         * 
         * *****************************************/
        public List<Asset> fillAsset(long vid)
        {

            List<Asset> asl = new List<Asset>();
            //////////////////////////////////////////////////////
            //   First Get Each Asset ID based on the VisitId   //
            //////////////////////////////////////////////////////
            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT ASSET_NAME, nvl(TERMINALID,0) TERMINALID, TERMINAL, nvl(ASSET_ID,0) asset_id" +
                            "  FROM vw_asset_overview" +
                            " WHERE VISITID=" + vid +
                            " ORDER BY ASSET_NAME";

            srcCmd = new OleDbCommand(strSQL, conn2);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {
                // For Each Asset
                Asset ast = new Asset();
                ast.setAssetName(srcRead["ASSET_NAME"].ToString());         // Get the Asset Name
                long tid = long.Parse(srcRead["TERMINALID"].ToString());    // Get the Terminal Id
                long asid = long.Parse(srcRead["ASSET_ID"].ToString());     // Get the Asset Id
                ast.setTargetNames(fillTargetNames(tid, asid, vid));        // Get the List of Asset Target Names using VISITID and ASSET_ID
                ast.setOptions(fillOptions(tid, vid));                      // Get the set of Options using TERMINALID and VISITID
                ast.setTargetList(fillTargetList(tid, asid, vid));          // Get the List of Target Options
                ast.setAssetMap(fillAssetMapFigure(asid, tid, vid));        // Get the Asset Map - figure
                ast.setTerminalId((Int16)tid);
                ast.setTerminalName(srcRead["TERMINAL"].ToString());

                ast.setScenarios(getAssetScenarios(tid, asid));
                ast.setScenariosDetected(getAssetScenariosDetected(tid, asid));

                ast.setSituations(getAssetSituations(tid, asid));
                ast.setSituationsDetected(getAssetSituationsDetected(tid, asid));

                asl.Add(ast);                                               // Add this Asset to the Asset List
            }
            srcRead.Close();
            srcCmd.Dispose();

            return asl;                                                     // return the Asset List
        }

        private Int16 getAssetScenarios(long terminalId, long astId)
        {

            Int16 scenCount = 0;

            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT count(1) scenCount" +
                            "  FROM tbl_target_area_scen a, tbl_target_area b" +
                            " WHERE b.terminalid = " + terminalId.ToString() +
                            "   AND b.asset_id = " + astId.ToString() +
                            "   AND b.visitid = " + VisitId.ToString() +
                            "   AND a.target_area_id = b.target_area_id";

            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;

            srcRead = srcCmd.ExecuteReader();

            int i = 0;
            while (srcRead.Read())
            {
                i++;
                scenCount = Int16.Parse(srcRead[0].ToString());

            }
            srcRead.Close();


            return scenCount;

        }

        private Int16 getAssetScenariosDetected(long terminalId, long astId)
        {

            Int16 scenCount = 0;

            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT count(*) scenCount" +
                            "  FROM tbl_target_area_scen a, tbl_target_area b" +
                            " WHERE b.terminalid = " + terminalId +
                            "   AND b.asset_id = " + astId +
                            "   AND b.visitid = " + VisitId +
                            "   AND a.target_area_id = b.target_area_id" +
                            "   AND a.detect_minute_id in (1,2) ";

            srcCmd = new OleDbCommand(strSQL, conn2);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            if (srcRead.Read())
            {
                scenCount = Int16.Parse(srcRead["SCENCOUNT"].ToString());
            }
            srcRead.Close();
            srcCmd.Dispose();
            return scenCount;

        }

        private Int16 getAssetSituations(long terminalId, long astId)
        {

            Int16 scenCount = 0;

            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT count(1) sCount" +
                            "  FROM tbl_detect_behave a, tbl_target_area b" +
                            " WHERE b.terminalid = " + terminalId +
                            "   AND b.asset_id = " + astId +
                            "   AND b.visitid = " + VisitId +
                            "   AND a.target_area_id = b.target_area_id";


            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;

            srcRead = srcCmd.ExecuteReader();

            int i = 0;
            while (srcRead.Read())
            {
                i++;
                scenCount = Int16.Parse(srcRead[0].ToString());

            }
            srcRead.Close();


            return scenCount;

        }

        private Int16 getAssetSituationsDetected(long terminalId, long astId)
        {

            Int16 scenCount = 0;

            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT count(*) sCount" +
                            "  FROM tbl_detect_behave a, tbl_target_area b" +
                            " WHERE b.terminalid = " + terminalId +
                            "   AND b.asset_id = " + astId +
                            "   AND b.visitid = " + VisitId +
                            "   AND a.target_area_id = b.target_area_id" +
                            "   AND a.detect_hour_id in (1,2,3,4) ";

            srcCmd = new OleDbCommand(strSQL, conn2);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            if (srcRead.Read())
            {
                scenCount = Int16.Parse(srcRead[0].ToString());
            }
            srcRead.Close();
            srcCmd.Dispose();
            return scenCount;

        }
        public List<Asset> fillIEDAsset(long vid)
        {

            List<Asset> asl = new List<Asset>();
            //////////////////////////////////////////////////////
            //   First Get Each Asset ID based on the VisitId   //
            //////////////////////////////////////////////////////
            OleDbCommand srcCmd;
            OleDbDataReader srcRead;


            string strSQL = @"select va.asset_name, nvl(va.terminalid, 0) terminalid, va.terminal, nvl(va.asset_id,0) asset_id, sc.ied_score, ast.asset_ovr_vw
                                from vw_asset_overview va, tbl_asset_map ast,
                                    (select terminal_id, score_type asset_id, round(ied_score,2) ied_score
                                        from (
                                        select * from evaat_rpt_scores
                                        where airport_id = :airp
                                          and round(ied_score,2) > 0 and terminal_id is not null
                                        order by round(ied_score,2) desc)
                                    where rownum < 6) sc
                                where  nvl(sc.terminal_id,0) = nvl(va.terminalid,0)
                                  and sc.asset_id = va.asset_id
                                  and va.visitid = :vid 
                                  and ast.visitid = va.visitid
                                  and nvl(va.terminalid,0) = nvl(ast.terminalid,0)
                                  and ast.asset_id = va.asset_id
                                order by ied_score desc";

            srcCmd = new OleDbCommand(strSQL, conn2);
            srcCmd.CommandType = CommandType.Text;
            srcCmd.Parameters.AddWithValue("airp", AirportId);
            srcCmd.Parameters.AddWithValue("Vid", vid);

            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {
                // For Each Asset
                Asset ast = new Asset();
                ast.setAssetName(srcRead["ASSET_NAME"].ToString());         // Get the Asset Name
                long tid = long.Parse(srcRead["TERMINALID"].ToString());    // Get the Terminal Id
                long asid = long.Parse(srcRead["ASSET_ID"].ToString());     // Get the Asset Id
                ast.setTerminalName(srcRead["TERMINAL"].ToString());
                ast.setTargetNames(fillTargetNames(tid, asid, vid));             // Get the List of Asset Target Names using VISITID and ASSET_ID
                ast.setOptions(fillOptions(tid, vid));                      // Get the set of Options using TERMINALID and VISITID
                ast.setTargetList(fillTargetList(tid, asid, vid));               // Get the List of Target Options
                ast.setAssetMap(fillAssetMapFigure(asid, tid, vid));                     // Get the Asset Map - figure
                ast.setTSSEOverview(srcRead["ASSET_OVR_VW"].ToString());
                asl.Add(ast);                                               // Add this Asset to the Asset List
            }
            srcRead.Close();
            srcCmd.Dispose();

            return asl;                                                     // return the Asset List
        }

        public List<Asset> fillVBIEDAsset(long vid)
        {

            List<Asset> asl = new List<Asset>();
            //////////////////////////////////////////////////////
            //   First Get Each Asset ID based on the VisitId   //
            //////////////////////////////////////////////////////
            OleDbCommand srcCmd;
            OleDbDataReader srcRead;

            string strSQL = @"select va.asset_name, nvl(va.terminalid, 0) terminalid, va.terminal, nvl(va.asset_id,0) asset_id
                                from vw_asset_overview va,
                                    (select terminal_id, score_type asset_id, round(vbied_score,2) vbied_score
                                        from (
                                        select * from evaat_rpt_scores
                                        where airport_id = :airp
                                          and round(vbied_score,2) > 0 and terminal_id is not null
                                        order by round(vbied_score,2) desc)
                                    where rownum < 6) sc
                                where  nvl(sc.terminal_id,0) = nvl(va.terminalid,0)
                                  and sc.asset_id = va.asset_id
                                  and va.visitid = :vid
                                order by vbied_score desc ";

            srcCmd = new OleDbCommand(strSQL, conn2);
            srcCmd.CommandType = CommandType.Text;
            srcCmd.Parameters.AddWithValue("airp", AirportId);
            srcCmd.Parameters.AddWithValue("Vid", vid);
            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {
                // For Each Asset
                Asset ast = new Asset();
                ast.setAssetName(srcRead["ASSET_NAME"].ToString());         // Get the Asset Name
                long tid = long.Parse(srcRead["TERMINALID"].ToString());    // Get the Terminal Id
                long asid = long.Parse(srcRead["ASSET_ID"].ToString());     // Get the Asset Id
                ast.setTerminalName(srcRead["TERMINAL"].ToString());
                ast.setTargetNames(fillTargetNames(tid, asid, vid));             // Get the List of Asset Target Names using VISITID and ASSET_ID
                ast.setOptions(fillOptions(tid, vid));                      // Get the set of Options using TERMINALID and VISITID
                ast.setTargetList(fillTargetList(tid, asid, vid));               // Get the List of Target Options
                ast.setAssetMap(fillAssetMapFigure(asid, tid, vid));                     // Get the Asset Map - figure
                asl.Add(ast);                                               // Add this Asset to the Asset List
            }
            srcRead.Close();
            srcCmd.Dispose();

            return asl;                                                     // return the Asset List
        }





        /* *****************************************
         *  Get the List of Target Names
         * 
         *   Populate an ArrayList with each of the
         *   Target Names
         * 
         * *****************************************/
        public ArrayList fillTargetNames(long tid, long asid, long vid)
        {
            ArrayList arl = new ArrayList();
            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT ASSET, TA_NAME" +
                            " FROM vw_target_area" +
                            " WHERE TERMINALID = " + tid +
                            "   AND ASSET_ID=" + asid + " AND VISITID=" + vid;

            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {
                string ass = srcRead["ASSET"].ToString();
                string trg = srcRead["TA_NAME"].ToString();
                arl.Add(ass + " - " + trg);
            }
            srcRead.Close();
            srcCmd.Dispose();

            return arl;
        }


        /* *****************************************
         * Fill Up the Target List
         * 
         *  For Each Target
         *  * Get the Asset Name
         *  * Get the Target Name
         *  * Get the Target Id
         *  * Get the Target Justification
         *  * Get Target Justification
         *  * Get Target Observations
         *  * Get Target's DeterM
         *  * Get Target List for ThreatS
         * 
         * *****************************************/
        public List<Target> fillTargetList(long terminalId, long asid, long vid)
        {
            List<Target> tgl = new List<Target>();
            ///////////////////////////
            // Populate Each Target
            //////////////////////////
            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT ASSET, TA_NAME, nvl(TARGET_AREA_ID,0) target_area_id, TA_JUSTIFICATION, NVL(TERMINALID,0) TERMINALID" +
                            " FROM vw_target_area" +
                            " WHERE TERMINALID = " + terminalId +
                            "   AND ASSET_ID=" + asid +
                            " AND VISITID=" + vid;

            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {                                            // for each Target
                Target trg = new Target();
                string ast = srcRead["ASSET"].ToString();                       // Get Asset Name
                string tan = srcRead["TA_NAME"].ToString();                     // Get the Target Name
                long tgid = long.Parse(srcRead["TARGET_AREA_ID"].ToString());   // Get the TARGET_AREA_ID
                long tid = long.Parse(srcRead["TERMINALID"].ToString());
                //                trg.setTarget_Name(ast + " - " + tan);                          // Set the Target Name
                trg.setTarget_Name(tan);                          // Set the Target Name
                trg.setJustify(srcRead["TA_JUSTIFICATION"].ToString());         // get and Set the Target Justification
                trg.setObservation(fillObservations(tgid));                     // get and Set the ArrayList of Observations
                trg.setOptions(fillOptions(tgid));                     // get and Set the ArrayList of Observations
                trg.setDeter(fillDeterM(tgid));                                 // get and Set the DeterM for the Target
                trg.setThreat(fillThreatS(tgid));                               // get and Set the ThreatS for the Target
                trg.setTargetMap(fillMapFigure(tid, tgid));                     // get and Set the Target Map for the Target
                trg.setTargetFigure(fillImageFigure(tid, tgid));
                tgl.Add(trg);
            }
            srcRead.Close();
            srcCmd.Dispose();

            return tgl;
        }



        /* *****************************************
         *  Populate the Target's Observations
         * 
         * 
         * *****************************************/
        public ArrayList fillObservations(long tid)
        {
            ArrayList obl = new ArrayList();

            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT OB_POINT_DESCRIPTION" +
                            " FROM TBL_OB_POINT" +
                            " WHERE TARGET_AREA_ID=" + tid;

            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {
                obl.Add(srcRead["OB_POINT_DESCRIPTION"].ToString());
            }
            srcRead.Close();
            srcCmd.Dispose();
            return obl;
        }

        /* *****************************************
         *  Populate the Target's Options for consideration
         * 
         * *****************************************/
        public ArrayList fillOptions(long tid)
        {
            ArrayList obl = new ArrayList();

            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT DESCRIPTION" +
                            " FROM vw_ofc" +
                            " WHERE TARGET_AREA_ID=" + tid;

            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {
                obl.Add(srcRead["DESCRIPTION"].ToString());
            }
            srcRead.Close();
            srcCmd.Dispose();
            return obl;
        }

        /* *****************************************
         *  Populate the Asset's Options
         * 
         *  - Get the Description Field base on the
         *    terminal Id and Visit Id
         * 
         * *****************************************/
        public ArrayList fillOptions(long tid, long vid)
        {
            ArrayList olst = new ArrayList();

            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT DESCRIPTION" +
                            "  FROM vw_ofc" +
                            " WHERE VISITID=" + vid +
                            "   AND TERMINALID=" + tid +
                //"   AND HEADING_NAME='Options for Consideration'" +//commented this line by QZ on July 16, 2014 because there is no record in vw_ofc having the ND HEADING_NAME='Options for Consideration'
                            " ORDER BY DESCRIPTION";

            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {
                olst.Add(srcRead["DESCRIPTION"].ToString());
            }
            srcRead.Close();
            srcCmd.Dispose();
            return olst;
        }


        /* *****************************************
         *  Populate the DeterM Object
         * 
         * - The the total Team Count
         * - For Each DETERRENCE_ID
         *   *  Get the count
         * 
         * *****************************************/
        public DeterM fillDeterM(long tid)
        {
            DeterM det = new DeterM();
            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT COUNT (DISTINCT(TARGET_TEAM_ID)) AS CNT" +
                            " FROM TBL_DETER_MEASURE" +
                            " WHERE TARGET_AREA_ID =" + tid;

            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {
                det.setNumTeam(int.Parse(srcRead["CNT"].ToString()));
            }
            srcRead.Close();
            srcCmd.Dispose();

            for (int id = 1; id < 12; id++)
            {
                strSQL = "SELECT COUNT (DISTINCT(TARGET_TEAM_ID)) AS CNT" +
                          " FROM TBL_DETER_MEASURE" +
                          " WHERE TARGET_AREA_ID = " + tid +
                          " AND ((IS_IED is not NULL) OR (IS_VBIED is not NULL))" +
                          " AND DETERRENCE_ID =" + id;

                srcCmd = new OleDbCommand(strSQL, conn1);
                srcCmd.CommandType = CommandType.Text;
                srcRead = srcCmd.ExecuteReader();

                while (srcRead.Read())
                {
                    switch (id)
                    {
                        case 1:
                            det.setNumCCTV(int.Parse(srcRead["CNT"].ToString()));   // CCTV
                            break;

                        case 2:
                            det.setNumSec(int.Parse(srcRead["CNT"].ToString()));    // Security
                            break;

                        case 3:
                            det.setNumLight(int.Parse(srcRead["CNT"].ToString()));  // Lighting
                            break;

                        case 4:
                            det.setNumEmp(int.Parse(srcRead["CNT"].ToString()));    // Airport Employee
                            break;

                        case 5:
                            det.setNumWall(int.Parse(srcRead["CNT"].ToString()));   // Blast Walls
                            break;

                        case 6:
                            det.setNumGlass(int.Parse(srcRead["CNT"].ToString()));  // Glass Mitigation
                            break;

                        case 7:
                            det.setNumRand(int.Parse(srcRead["CNT"].ToString()));   // Random Security
                            break;

                        case 8:
                            det.setNumBarr(int.Parse(srcRead["CNT"].ToString()));   // Barriers
                            break;

                        case 9:
                            det.setNumPerm(int.Parse(srcRead["CNT"].ToString()));   // Parimeter
                            break;

                        case 10:
                            det.setNumSign(int.Parse(srcRead["CNT"].ToString()));   // Signage
                            break;

                        case 11:
                            det.setNumPublic(int.Parse(srcRead["CNT"].ToString())); // General Public
                            break;

                        default:
                            break;
                    }
                }
                srcRead.Close();
                srcCmd.Dispose();
            }

            return det;
        }



        /* *****************************************
         *  Populate ThreatS List
         *  
         *  Based on the Target_Area_Id get all
         *  cases:
         *   - DETECT_THREAT_ID
         *   - DETECT_MINUTE_ID
         *   - DETECT_MEASURE_ID
         *   - SCEN_DATETIME
         * 
         * *****************************************/
        public List<ThreatS> fillThreatS(long tid)
        {
            List<ThreatS> thrLst = new List<ThreatS>();
            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT to_char(SCEN_DATETIME, 'MM/DD/YY HH24:MI') SCEN_DATETIME, DETECT_THREAT_ID, DETECT_MINUTE_ID, DETECT_MEASURE_ID" +
                            " FROM TBL_TARGET_AREA_SCEN" +
                            " WHERE TARGET_AREA_ID =" + tid +
                            " AND SCEN_DATETIME IS NOT NULL AND DETECT_THREAT_ID IS NOT NULL AND DETECT_MINUTE_ID IS NOT NULL AND DETECT_MEASURE_ID IS NOT NULL";

            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {
                ThreatS ths = new ThreatS();
                ths.setThreat_Id(int.Parse(srcRead["DETECT_THREAT_ID"].ToString()));
                ths.setMinute_Id(int.Parse(srcRead["DETECT_MINUTE_ID"].ToString()));
                string detId = srcRead["DETECT_MEASURE_ID"].ToString();
                if (detId.Length > 0)
                    ths.setDetect_Id(int.Parse(detId));
                else
                    ths.setDetect_Id(-1);
                //                ths.setThreatDate(((DateTime)srcRead["SCEN_DATETIME"]).ToString("MM/dd/yyyy"));
                ths.setThreatDate(srcRead["SCEN_DATETIME"].ToString());
                thrLst.Add(ths);
            }
            srcRead.Close();
            srcCmd.Dispose();
            return thrLst;
        }








        /* **********************************
         * 
         * 
         * 
         * 
         * **********************************/
        public List<string> fillDetMeasure()
        {

            List<string> lst = new List<string>();
            int id;
            string Desc;

            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT DETECT_MEASURE_ID, DESCRIPTION" +
                            " FROM LU_DETECT_MEASURE";

            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {
                id = int.Parse(srcRead["DETECT_MEASURE_ID"].ToString());
                Desc = srcRead["DESCRIPTION"].ToString();
                lst.Insert(id - 1, Desc);
            }
            srcRead.Close();
            srcCmd.Dispose();

            return lst;
        }



        /* **********************************
         * 
         * 
         * 
         * 
         * **********************************/
        public List<Figure> fillImageFigure(long TerminalId, long TargId)
        {
            List<Figure> lfg = new List<Figure>();

            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT TARGET_AREA_BLOB, 'TargetAreaPhoto.jpg' FILE_NAME" +
                            " FROM TBL_TARGET_AREA" +
                            " WHERE TARGET_AREA_ID=" + TargId +
                            " AND TARGET_AREA_BLOB IS NOT NULL";

            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {
                Figure rtn = new Figure();
                rtn.setFileName(srcRead["FILE_NAME"].ToString());
                if (!Convert.IsDBNull(srcRead["TARGET_AREA_BLOB"]))
                    rtn.setFigure((byte[])srcRead["TARGET_AREA_BLOB"]);
                else
                    rtn.setFigure(new byte[0]);
                lfg.Add(rtn);
            }
            srcRead.Close();
            srcCmd.Dispose();

            return lfg;
        }


        /* **********************************
         * 
         * 
         * 
         * 
         * **********************************/
        public Figure fillMapFigure(long AssetId, long TargId)
        {
            Figure rtn = new Figure();

            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT BLAST_BLOB, BLAST_FILENAME" +
                            " FROM TBL_TARGET_AREA" +
                            " WHERE TARGET_AREA_ID=" + TargId;

            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {
                rtn.setFileName(srcRead["BLAST_FILENAME"].ToString());
                if(!Convert.IsDBNull(srcRead["BLAST_BLOB"]))
                    rtn.setFigure((byte[])srcRead["BLAST_BLOB"]);
                else
                    rtn.setFigure(new byte[0]);
            }
            srcRead.Close();
            srcCmd.Dispose();

            return rtn;
        }


        /* **********************************
         * 
         * 
         * 
         * 
         * **********************************/
        public Figure fillLayoutFigure(long AirportId)
        {
            Figure rtn = new Figure();

            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT TERMINAL_LAYOUT, 'AirportLayout.jpg' FILE_NAME" +
                            " FROM AIRPORTS" +
                            " WHERE ID=" + AirportId +
                            " AND TERMINAL_LAYOUT IS NOT NULL";

            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {
                rtn.setFileName(srcRead["FILE_NAME"].ToString());
                if (!Convert.IsDBNull(srcRead["TERMINAL_LAYOUT"]))
                    rtn.setFigure((byte[])srcRead["TERMINAL_LAYOUT"]);
                else
                    rtn.setFigure(new byte[0]);
            }
            srcRead.Close();
            srcCmd.Dispose();

            return rtn;
        }


        /* ************************************
         * Fill Asset Map Figures
         * 
         * - Get the BLOBs based on
         *   o TERMINALID
         *   o ASSET_ID
         *   o VISITID
         * 
         * ************************************/
        public Figure fillAssetMapFigure(long aid, long tid, long visitid)
        {
            Figure rtn = new Figure();
            rtn.setFileName("");

            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT BLOB_FILE, FILE_NAME" +
                            " FROM TBL_ASSET_MAP" +
                            " WHERE TERMINALID=" + tid +
                            " AND ASSET_ID=" + aid +
                            " AND VISITID=" + visitid +
                            " AND BLOB_FILE IS NOT NULL";

            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;
            srcRead = srcCmd.ExecuteReader();

            while (srcRead.Read())
            {
                rtn.setFileName(srcRead["FILE_NAME"].ToString());
                if (!Convert.IsDBNull(srcRead["BLOB_FILE"]))
                    rtn.setFigure((byte[])srcRead["BLOB_FILE"]);
                else
                    rtn.setFigure(new byte[0]);
            }
            srcRead.Close();
            srcCmd.Dispose();

            return rtn;
        }

        public void putRptDB(string product_key, string FileNamePath, long vid)
        {
            string FileName = FileNamePath.Substring(FileNamePath.LastIndexOf('\\') + 1, FileNamePath.Length - FileNamePath.LastIndexOf('\\') - 1);

            FileStream fs = new FileStream(FileNamePath, FileMode.OpenOrCreate, FileAccess.Read);
            Byte[] MyData = new Byte[fs.Length];
            fs.Read(MyData, 0, (int)fs.Length);
            fs.Close();

            string strSQL = "DELETE FROM TBL_PRODUCTS" +
                            " WHERE PRODUCT_TYPE = '" + product_key + "'" +
                            "   AND VISITID =" + vid;

            OleDbCommand updCommand = new OleDbCommand(strSQL, conn1);

            try
            {
                updCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }


            strSQL = "INSERT INTO TBL_PRODUCTS" +
                     " (PRODUCT_TYPE, BLOB_FILE, FILE_NAME, VISITID)" +
                     " VALUES (?, ?, ?, ?)";

            updCommand = new OleDbCommand(strSQL, conn1);
            updCommand.Parameters.Add("PRODUCT_KEY", OleDbType.VarChar).Value = product_key;
            updCommand.Parameters.Add("BLOB_FILE", OleDbType.Binary).Value = MyData;
            updCommand.Parameters.Add("FILE_NAME", OleDbType.VarChar).Value = FileName;
            updCommand.Parameters.Add("VISITID", OleDbType.Integer).Value = vid;

            try
            {
                updCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            updCommand.Dispose();
        }

        //Added by QZ just for inserting updated template files into Oracle table TBL_PRODUCTS
        public void putTemplateDB(string product_key, string FileNamePath)
        {
            string FileName = FileNamePath.Substring(FileNamePath.LastIndexOf('\\') + 1, FileNamePath.Length - FileNamePath.LastIndexOf('\\') - 1);

            FileStream fs = new FileStream(FileNamePath, FileMode.OpenOrCreate, FileAccess.Read);
            Byte[] MyData = new Byte[fs.Length];
            fs.Read(MyData, 0, (int)fs.Length);
            fs.Close();

            string strSQL = "DELETE FROM TBL_PRODUCTS" +
                            " WHERE PRODUCT_TYPE = '" + product_key + "'";

            OleDbCommand updCommand = new OleDbCommand(strSQL, conn1);

            try
            {
                updCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            strSQL = "INSERT INTO TBL_PRODUCTS" +
                     " (PRODUCT_TYPE, BLOB_FILE, FILE_NAME)" +
                     " VALUES (?, ?, ?)";

            updCommand = new OleDbCommand(strSQL, conn1);
            updCommand.Parameters.Add("PRODUCT_KEY", OleDbType.VarChar).Value = product_key;
            updCommand.Parameters.Add("BLOB_FILE", OleDbType.Binary).Value = MyData;
            updCommand.Parameters.Add("FILE_NAME", OleDbType.VarChar).Value = FileName;

            try
            {
                updCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            updCommand.Dispose();
        }
        public void DownloadTemplate(string templateName, string filename)
        {
            OleDbCommand srcCmd;
            OleDbDataReader srcRead;
            string strSQL = "SELECT  BLOB_FILE, FILE_NAME FROM TBL_PRODUCTS where PRODUCT_TYPE = :templateName";

            srcCmd = new OleDbCommand(strSQL, conn1);
            srcCmd.CommandType = CommandType.Text;
            srcCmd.Parameters.AddWithValue("templateName", templateName);
            //*Commented out by QZ on July 7, 2014 so that I could use the local updated copy of the template files
            srcRead = srcCmd.ExecuteReader();
            Byte[] b = null;
            while (srcRead.Read())
            {
                //rtn.setFileName(srcRead["FILE_NAME"].ToString());
                b = (byte[])srcRead["BLOB_FILE"];
            }
            srcRead.Close();
            srcCmd.Dispose();


            try
            {
                System.IO.File.Delete(filename);
            }
            catch (System.IO.IOException)
            {
                //DO NOTHING
            }

            System.IO.FileStream fs = new System.IO.FileStream(filename, System.IO.FileMode.Create, System.IO.FileAccess.Write);

            fs.Write(b, 0, b.Length);
           
            fs.Close();
            //*/
        }
    }
}
