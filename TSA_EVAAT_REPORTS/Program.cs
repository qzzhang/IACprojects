using System;
using System.Data;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
//using Microsoft.Office.Interop.Excel;

namespace Airport_Asset
{
    class Program
    {

        // The Program's properties
        static long VisitId;        // Visit Id - key to Facility Data - populated from command line
        static string Template;     // Path and Filename of the Template used as base to create document - populated from command line
        static string Output;       // Path to the Output Directory - populated from the command line
        static string ReportType = "";
        static DBManager dbm = null;
        static StreamWriter logFile = null;
        static Airport airPort;

        static Microsoft.Office.Interop.Word.ApplicationClass objWord;
        static Microsoft.Office.Interop.Word.Document objWordDoc;
        static Microsoft.Office.Interop.Word.Document objWordTempl;

        static int FigureNum;       // Counter to keep Track of the Figure Num - Used in the Figure Captions and in the Text

        // static List<string> detMeas;        // List of String Descriptions for each DETECT_MEASURE_ID
        // be used when populating the Threat Scenario Tables
        const int startThreatDataRow = 3;   // The Starting row in the Threat Scenario Tables where data will added
        //private static Range rng;


        /* ***************************************************
         *  Main() - Start Method for Word Generation
         * ***************************************************
         * - Populate properties from Command Line
         * - Populate the Airport Object From Database
         * - Debug to check Airport Object Data
         * - Populate the Document (using the Airport Object)
         * 
         * ****************************************************/
        [STAThread]
        static void Main(string[] args)
        {
            Airport airp;
            try
            {
                
                Output = args[0];                          // 1: Path to the Output Folder
                VisitId = long.Parse(args[1].ToString());   // 2: Visit Id value

                logFile = new StreamWriter(Output + "\\LogFile.txt");
                
                if (args.Length >= 3)
                {
                    ReportType = args[2];
                }

                dbm = new DBManager();
                dbm.openDB();

                airp = dbm.fillAirport(VisitId);    // Populate the Airport Object - will have all the data for the Report

                airPort = airp;

                if (ReportType == "" || ReportType == "FULL")
                {
                    FullReport_populateDocument(airp);
                }

                if (ReportType == "" || ReportType == "STAKEHOLDER")
                {
                    Stakeholder_PopulateDocument(airp);
                }

                if (ReportType == "" || ReportType == "EXECSUMMARY")
                {                   
                    ExecSummary_populateDocument(airp);
                }
                dbm.closeDB();
                Console.WriteLine("Done: Press Enter. ");
                //Console.Read();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Main: " + ex.Message);
                dbm.closeDB();
                quit_and_close();
            }
        }


        static void log(string message)
        {
            Console.WriteLine(message);
            logFile.WriteLine(message);
            logFile.Flush();

        }
        static bool open_document(string fileName, string template_key)
        {
            object fileNameObj, templObj;
            object novalue = System.Reflection.Missing.Value;
            object isVisible = true;
            object readOnly = false;
            object noPrompt = false;

            object originalFormat = Type.Missing;
            object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdSaveChanges;
            object NoSaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
            object routeDocument = Type.Missing;

            bool retval = false;

            try
            {
                dbm.DownloadTemplate(template_key, Template);

                fileNameObj = (Object)fileName;
                templObj = (Object)Template;

                System.IO.Directory.CreateDirectory(Output);
                System.IO.File.Copy(Template, fileName, true);

                // Opening the Word Document
                objWord = new Microsoft.Office.Interop.Word.ApplicationClass();

                objWordDoc = objWord.Documents.Open(ref fileNameObj, ref novalue, ref readOnly, ref novalue, ref novalue,
                                                        ref novalue, ref novalue, ref novalue, ref novalue, ref novalue,
                                                        ref novalue, ref isVisible, ref novalue, ref novalue, ref novalue, ref novalue);

                objWordTempl = objWord.Documents.Open(ref templObj, ref novalue, ref readOnly, ref novalue, ref novalue,
                                                        ref novalue, ref novalue, ref novalue, ref novalue, ref novalue,
                                                        ref novalue, ref isVisible, ref novalue, ref novalue, ref novalue, ref novalue);
                retval = true;
            }
            catch (Exception fileopenEx)
            {
                Console.WriteLine("Error Opening documents: " + fileopenEx.Message);
                dbm.closeDB();
            }
            return retval;
        }

        static void quit_and_close()
        {
            object NoSaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
            object novalue = System.Reflection.Missing.Value;

            //This objWordTemp1 could be null when quit_and_close() is called from, e.g. Main() after open_document() fails, so I add the if check here--QZ on June 10, 2014
            if( objWordTempl != null )
                objWordTempl.Close(NoSaveChanges);
            if (objWordDoc != null)
                objWordDoc.Close(NoSaveChanges);

            objWord.Quit(ref novalue, ref novalue, ref novalue);

        }


        static void save_and_close(string fileName, string product_key)
        {
            object novalue = System.Reflection.Missing.Value;
            object isVisible = true;
            object readOnly = false;
            object noPrompt = false;

            object originalFormat = Type.Missing;
            object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdSaveChanges;
            object NoSaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
            object routeDocument = Type.Missing;


            //Remove blank pages:

            Paragraphs paragraphs = objWordDoc.Paragraphs;
            foreach (Paragraph paragraph in paragraphs)
            {
                if (paragraph.Range.Words.Count == 0)
                {
                    paragraph.Range.Select();
                    objWord.Selection.Delete();
                }
            }

            //Added by QZ on July 8, 2014 to clean up unnecessary bookmarks.
            if (objWordDoc.Bookmarks.Exists("A_TargetBlock"))
                objWordDoc.Bookmarks["A_TargetBlock"].Delete();
            if (objWordDoc.Bookmarks.Exists("B_TargetBlock"))
                objWordDoc.Bookmarks["B_TargetBlock"].Delete();


            // Update the Contents Sections:

            for( int ci =1; ci <= objWordDoc.TablesOfContents.Count; ci++ )
            {
                objWordDoc.TablesOfContents[ci].Update();
                objWordDoc.TablesOfContents[ci].Range.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                objWordDoc.TablesOfContents[ci].UpdatePageNumbers();
            }
            for( int fi =1; fi <= objWordDoc.TablesOfFigures.Count; fi++ )
            {
                objWordDoc.TablesOfFigures[fi].Update();
                objWordDoc.TablesOfFigures[fi].Range.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                objWordDoc.TablesOfFigures[fi].UpdatePageNumbers();
            }
            /*Tested by QZ
            object oTrueValue = true;
            object cstart = objWordDoc.Content.End - 1;
            object missing = System.Type.Missing;
            Range rangeForTOC = objWordDoc.Range(ref cstart, ref missing);
            TableOfContents toc = objWordDoc.TablesOfContents.Add(rangeForTOC,
                ref oTrueValue, ref missing, ref missing,
                ref missing, ref missing, ref oTrueValue,
                ref oTrueValue, ref oTrueValue, ref oTrueValue,
                ref oTrueValue, ref oTrueValue);
            toc.Update();
            Range rngTOC = toc.Range;
            rngTOC.Font.Size = 10;
            rngTOC.Font.Name = "Georgia";
            */
            Shapes shps = objWordDoc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes;

            foreach (Shape shp in shps)
            {
                if (shp.Name.Contains("WaterMark"))
                {
                    shp.Delete();
                }
            }

            System.Windows.Forms.Clipboard.Clear();

            // Add Password to Document
             objWordDoc.Password = "TVC8*tvc";

            // Save and Close the Document
            objWordTempl.Close(NoSaveChanges,Type.Missing, Type.Missing);
            objWordDoc.Close(saveChanges,Type.Missing, Type.Missing);

            objWord.Quit(ref novalue, ref novalue, ref novalue);

            dbm.putRptDB(product_key, fileName, VisitId);

            Console.WriteLine("Done");
        }


        static void Stakeholder_PopulateDocument(Airport airp)
        {
            string fileName;
            FigureNum = 1;
            try
            {
                fileName = Output + '\\' + airp.getAirportName() + " Stakeholder Report.docx";

                Template = Output + "\\evaat_stakeholder_template.docx";

                string template_key = "EVAAT_STAKEHOLDER_REPORT_TEMPLATE";

                if (open_document(fileName, template_key))
                {
                    // Populate the Title Page
                    populateTitle(airp);

                    // Populate the Introduction
                    Stakeholder_populateIntroduction(airp);

                    Populate_Top5_By_Attractiveness();

                    //// Populate IED Threat Assets
                    Stakeholder_populateIEDThreatAssetList(airp.getIEDAssetList());

                    Stakeholder_populateVBIEDThreatAssetList(airp.getVBIEDAssetList());

                    string product_key = "EVAAT_STAKEHOLDER_REPORT";
                    save_and_close(fileName, product_key);
                }
                else
                {
                    Console.WriteLine("Error Opening document: " + fileName + " or template " + template_key);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Stakeholder_populateDocument: " + ex.Message);
                quit_and_close();
            }
        }


        static void ExecSummary_populateDocument(Airport airp)
        {

            string fileName;
            FigureNum = 1;
            fileName = Output + '\\' + airp.getAirportName() + " EXEC SUMMARY.docx";

            Template = Output + "\\EVAAT EXEC SUMMARY Template.docx";

            string template_key = "EVAAT_EXEC_SUMMARY_TEMPLATE";

            if (open_document(fileName, template_key))
            {
                // Populate the Title Page
                populateTitle(airp);

                // Populate the Introduction
                ExecSummary_populateIntroduction(airp);

                Populate_Top5_By_Attractiveness();

                populate_options_for_consideration();

                string product_key = "EVAAT_EXEC_SUMMARY";
                save_and_close(fileName, product_key);
            }
            else
            {
                Console.WriteLine("Error Opening document: " + fileName + " or template " + template_key);
            }

        }



        /* **************************************************
         * Populate the Document
         * 
         * This is the beginning method used to populate the
         * Final Word Document
         * 
         * - Set up the Output File
         * - Open the word Document
         * - Populate the Title Page
         * - Populate the Introduction
         * - Populate the Asset Information
         * - Update the Contents Sections
         * - Add Password to Document
         * - Save and Close the Document
         * 
         * *************************************************/
        static void FullReport_populateDocument(Airport airp)
        {

            string fileName = "";

            FigureNum = 1;
            try
            {
                fileName = Output + '\\' + airp.getAirportName() + " Full Report.docx";

                Template = Output + "\\evaat_full_template.docx";

                string template_key = "EVAAT_FULL_REPORT_TEMPLATE";

                if (open_document(fileName, template_key))
                {
                    // Populate the Title Page
                    populateTitle(airp);

                    // Populate the Introduction
                    populateIntroduction(airp);

                    Populate_Top5_By_Attractiveness();

                    //// Populate IED Threat Assets
                    populateIEDThreatAssetList(airp.getIEDAssetList());

                    populateVBIEDThreatAssetList(airp.getVBIEDAssetList());

                    populateAllAssetList(airp.getAssetNames());

                    populateAppendix("A", airp);

                    populateAppendix("B", airp);

                    string product_key = "EVAAT_FULL_REPORT";
                    save_and_close(fileName, product_key);
                }
                else
                {
                    Console.WriteLine("Error Opening document: " + fileName + " or template " + template_key);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: FullReport_populateDocument: " + ex.Message);
                quit_and_close();
            }
        }


        private static void populate_options_for_consideration()
        {

            Range blkRng = objWordDoc.Bookmarks["options_block"].Range;

            blkRng.Copy();

            foreach (Target tgt in airPort.getTopTargetList())
            {
                blkRng.Bookmarks["options_target_name"].Range.Text = tgt.getTerminalName() + " - " + tgt.getAssetName() + " - " + tgt.getTarget_Name();

                string optlist = "";
                foreach (string opt in tgt.getOptions())
                {
                    if (optlist.Length > 0) optlist += "\n" + opt;
                    else optlist = opt;
                }
                if (optlist == "") optlist = "No options";
                blkRng.Bookmarks["options"].Range.Text = optlist;

                blkRng.SetRange(blkRng.End, blkRng.End);
                blkRng.Paste();
            }

            objWordDoc.Bookmarks["options_block"].Delete();//added this line by QZ on July 3, 2014

            blkRng.Text = "";

        }

        static void Populate_Top5_By_Attractiveness()
        {
            Range rng = objWordDoc.Bookmarks["Table_Top5Targets"].Range;

            Table tbl = rng.Tables[1];

            tbl.AllowPageBreaks = false;

            ArrayList ast = new ArrayList();

            string astName = "";
            string mostAttr = "";
            int row = 1;
            foreach (Target tgt in airPort.getTopTargetList())
            {
                if (row == 1)
                {
                    mostAttr = tgt.getTarget_Name() + " in " + tgt.getTerminalName() + " - " + tgt.getAssetName();
                }

                row++;
                tbl.Cell(row, 1).Range.Text = tgt.getTerminalName();
                tbl.Cell(row, 2).Range.Text = tgt.getAssetName();
                tbl.Cell(row, 3).Range.Text = tgt.getTarget_Name();
                tbl.Cell(row, 4).Range.Text = "X";

            }

            tbl.Columns.AutoFit();

            rng = objWordDoc.Bookmarks["MostAttactive"].Range;
            rng.Text = mostAttr;

            // charts
            // group by asset
            List<Target> tgtList = airPort.getTopTargetList();
            ast.Clear();
            for (int i = 0; i < tgtList.Count; i++)
            {
                astName = tgtList[i].getTerminalName() + " - " + tgtList[i].getAssetName();
                int found = 0;
                for (int j = 0; j < ast.Count; j++)
                {
                    if ((string)ast[j] == astName) { found = 1; break; }
                }

                if (found == 0)
                {
                    ast.Add(astName);
                }

            }

            rng = objWordDoc.Bookmarks["Target_Situational_Chart_Tmpl"].Range;
            rng.Copy();
            int figNum = FigureNum - 1;
            int count = 0;

            string figRange = "";

            foreach (string asset in ast)
            {
                count++;
                List<Target> tgtList2 = new List<Target>();

                foreach (Target tgt in airPort.getTopTargetList())
                {
                    if (asset == tgt.getTerminalName() + " - " + tgt.getAssetName())
                    {
                        tgtList2.Add(tgt);
                    }
                }
                figNum++;
                if (count == 1)
                    figRange = figNum.ToString();
                if (count > 1 && count < ast.Count)
                {
                    figRange += ", " + figNum.ToString();
                }

                if (count > 1 && count == ast.Count)
                {
                    figRange += " and " + figNum.ToString();
                }

                //rng.Bookmarks["FigureNum"].Range.Text = figNum.ToString();
                //rng.Bookmarks["AssetName"].Range.Text = asset;

                addTargetSituationalGraph(asset, tgtList2, rng);

                objWordDoc.Bookmarks["Target_Situational_Chart_Tmpl"].Delete();

                rng.SetRange(rng.End, rng.End);
                rng.Paste();

            }

            rng.Text = "";
    
            objWordDoc.Bookmarks["FigureRange"].Range.Text = figRange;
        }

        /* ***************************************************************
         *  Populate the Title Page
         *  - For each Book Mark 
         *    * extract the Value from the Airport Object
         *    * add the value to the Book Mark Range
         *    
         * ***************************************************************/
        static void populateTitle(Airport airp)
        {
            try
            {
                Microsoft.Office.Interop.Word.Bookmarks bpBookMarks = objWordDoc.Bookmarks;
                foreach (Microsoft.Office.Interop.Word.Bookmark bpBookMark in bpBookMarks)
                {
                    Range rng = bpBookMark.Range;
                    rng.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                    switch (bpBookMark.Name.ToUpper())
                    {
                        case "AIRPORTNAME":
                            rng.Text = airp.getAirportName();
                            break;
                        case "AIRABBREV":
                            rng.Text = airp.getAirportAbbv();
                            break;
                        case "AIRABBREV1"://added by QZ on July 7, 2014
                            rng.Text = airp.getAirportAbbv();
                            break;
                        case "CONDUCTED_DATES":
                            rng.Text = airp.getDuration();
                            break;
                        case "FSD_NAME":
                            rng.Text = airp.getFSDName();
                            break;
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("populationTitle Error: " + ex.Message);
            }
        }


/* ***************************************************************
 * Populate the Introduction
 * - This is only the List of Assets
 * - Put the List of Assets into a string
 * - The string must be formatted so that the last entry does
 *   not have a newLine at the end.
 * - Insert the string into the AssetList Book Mark
 * 
 * ***************************************************************/
        static void Stakeholder_populateIntroduction(Airport airp)
        {
            objWordDoc.Bookmarks["TSSE_Name"].Range.Text = airp.getTSSEName();

            objWordDoc.Bookmarks["EVAAT_DATES"].Range.Text = airp.getDuration();

            objWordDoc.Bookmarks["AIRABBREV2"].Range.Text = airp.getAirportAbbv();

            string[, ,] scores = airp.getRASScoresTerminal();

            insertTableRAT(ref scores);

            string[,] scores2 = airp.getRAScoresSupporting();

            insertTableRAS(ref scores2);

            Range rng = objWordDoc.Bookmarks["Airport_Situational_Chart_Tmpl"].Range;

            addSituationalGraph(airp, rng);

            //objWordDoc.Bookmarks["TSSE_Overview"].Range.HighlightColorIndex = WdColorIndex.wdGray25;//quote Lori Eaton "it doesn’t have to be in grey and italicized once pulled."
            objWordDoc.Bookmarks["TSSE_Overview"].Range.Italic = 0;
            objWordDoc.Bookmarks["TSSE_Overview"].Range.Text = airp.getTSSEOverview();

        }

        static void ExecSummary_populateIntroduction(Airport airp)
        {

            objWordDoc.Bookmarks["AIRABBREV2"].Range.Text = airp.getAirportAbbv();

            string[, ,] scores = airp.getRASScoresTerminal();

            insertTableRAT(ref scores);

            string[,] scores2 = airp.getRAScoresSupporting();

            insertTableRAS(ref scores2);

            Range rng = objWordDoc.Bookmarks["Airport_Situational_Chart_Tmpl"].Range;

            addSituationalGraph(airp, rng);

        }

        /* ***************************************************************
         * Populate the Introduction
         * - This is only the List of Assets
         * - Put the List of Assets into a string
         * - The string must be formatted so that the last entry does
         *   not have a newLine at the end.
         * - Insert the string into the AssetList Book Mark
         * 
         * ***************************************************************/
        static void populateIntroduction(Airport airp)
        {
            objWordDoc.Bookmarks["TSSE_Name"].Range.Text = airp.getTSSEName();

            objWordDoc.Bookmarks["EVAAT_DATES"].Range.Text = airp.getDuration();

            string assList = "";
            foreach (string ast in airp.getAssetNames())
            {
                if (assList.Length == 0)
                {
                    assList += ast;
                }
                else
                {
                    assList += Environment.NewLine + ast;

                }
            }
            objWordDoc.Bookmarks["VISITED_AREAS"].Range.Text = assList;

            //objWordDoc.Bookmarks["TSSE_Overview"].Range.HighlightColorIndex = WdColorIndex.wdGray25;//quote Lori Eaton "it doesn’t have to be in grey and italicized once pulled."
            objWordDoc.Bookmarks["TSSE_Overview"].Range.Italic = 0;
            objWordDoc.Bookmarks["TSSE_Overview"].Range.Text = airp.getTSSEOverview();


            Find fnd = objWordDoc.ActiveWindow.Selection.Find;

            fnd.ClearFormatting();
            fnd.Replacement.ClearFormatting();
            fnd.Forward = true;
            fnd.Wrap = WdFindWrap.wdFindContinue;

            fnd.Text = "[##Airport Abbreviation##]";
            fnd.Replacement.Text = airp.getAirportAbbv();

            fnd.Execute(Replace: WdReplace.wdReplaceAll);

            fnd.ClearFormatting();
            fnd.Replacement.ClearFormatting();
            fnd.Forward = true;
            fnd.Wrap = WdFindWrap.wdFindContinue;

            fnd.Text = "[##Airport Address##]";
            fnd.Replacement.Text = airp.getAirportAddr();

            fnd.Execute(Replace: WdReplace.wdReplaceAll);

            fnd.ClearFormatting();
            fnd.Replacement.ClearFormatting();
            fnd.Forward = true;
            fnd.Wrap = WdFindWrap.wdFindContinue;

            fnd.Text = "[##Airport City##]";
            fnd.Replacement.Text = airp.getAirportCity();
            fnd.Replacement.Highlight = 1;//added by QZ on July 7, 2014
            //fnd.Highlight = 1;

            fnd.Execute(Replace: WdReplace.wdReplaceAll);

            objWordDoc.Bookmarks["Category"].Range.Text = airp.getCategory();

            Range rng = objWordDoc.Bookmarks["Airport_Layout_Image"].Range;
            // rng.set_Style(ref normaltype);                     // - Set the Style - Normal
            rng.Text = "";
            addFigure(airp.getLayoutMap(), " Airport / Terminal Layout", rng);
            
            string[, ,] scores = airp.getRASScoresTerminal();

            insertTableRAT(ref scores);

            string[,] scores2 = airp.getRAScoresSupporting();

            insertTableRAS(ref scores2);

            rng = objWordDoc.Bookmarks["Airport_Situational_Chart_Tmpl"].Range;

            addSituationalGraph(airp, rng);
        }


        static void populateIEDThreatAssetList(List<Asset> al)
        {
            try
            {
                string astList = "";
                string astName = "";
                int i = 1;
                foreach (Asset ast in al)
                {
                    astName = ast.getAssetName();
                    if (astList.Length == 0)
                    {
                        astList += astName;
                    }
                    else
                    {
                        astList += Environment.NewLine + astName;
                    }
                }
                objWordDoc.Bookmarks["Top5_Threats_IED"].Range.Text = astList;

                Range srcRng = objWordTempl.Bookmarks["AssetBlock"].Range;

                Range astRange = objWordDoc.Bookmarks["AssetBlock"].Range;

                objWordDoc.Bookmarks["AssetBlock"].Delete();

                int alcount = al.Count;
                i = 0;

                foreach (Asset ast in al)
                {
                    i++;
                    srcRng.Copy();
                    astRange.Paste();

                    populateAssetBlock("IED", ast, astRange, i);

                    astRange.Bookmarks["AssetBlock"].Delete();

                    astRange.SetRange(astRange.End, astRange.End);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("populateIEDThreatAssetList error: " + ex.Message);
            }
        }

        static void populateVBIEDThreatAssetList(List<Asset> al)
        {
            try
            {
                string astList = "";
                string astName = "";
                int i = 1;
                foreach (Asset ast in al)
                {
                    astName = ast.getAssetName();
                    if (astList.Length == 0)
                    {
                        astList += astName;
                    }
                    else
                    {
                        astList += Environment.NewLine + astName;
                    }
                }

                astList += Environment.NewLine;
                astList += Environment.NewLine;

                Range rng;
                rng = objWordDoc.Bookmarks["Top5_Threats_VBIED"].Range;
                rng.Text = astList;

                //rng.HighlightColorIndex = WdColorIndex.wdNoHighlight;

                rng.ListFormat.ApplyNumberDefault();
                //rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                Range srcRng = objWordTempl.Bookmarks["AssetBlockVBIED"].Range;

                Range astRange = objWordDoc.Bookmarks["AssetBlockVBIED"].Range;

                objWordDoc.Bookmarks["AssetBlockVBIED"].Delete();

                int alcount = al.Count;
                i = 0;

                foreach (Asset ast in al)
                {
                    i++;
                    srcRng.Copy();
                    astRange.Paste();

                    populateAssetBlock("VBIED", ast, astRange, i);

                    astRange.Bookmarks["AssetBlockVBIED"].Delete();

                    astRange.SetRange(astRange.End, astRange.End);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("populateVBIEDThreatAssetList error: " + ex.Message);
            }
        }


        static void populateAllAssetList(ArrayList al)
        {
            try
            {
                string astList = "";
                string astName = "";
                foreach (string ast in al)
                {
                    astName = ast;
                    if (astList.Length == 0)
                    {
                        astList += astName;
                    }
                    else
                    {
                        astList += Environment.NewLine + astName;

                    }
                }

                astList += Environment.NewLine;
                astList += Environment.NewLine;

                Range rng;
                rng = objWordDoc.Bookmarks["AllAssets"].Range;
                rng.Text = astList;

                rng.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                rng.ListFormat.ApplyBulletDefault();


            }
            catch (Exception ex)
            {
                Console.WriteLine("populateAllAssetList error: " + ex.Message);
            }
        }


        static void Stakeholder_populateIEDThreatAssetList(List<Asset> al)
        {
            try
            {
                string astList = "";
                string astName = "";
                int i = 1;
                foreach (Asset ast in al)
                {
                    astName = ast.getAssetName();
                    if (astList.Length == 0)
                    {
                        astList += astName;
                    }
                    else
                    {
                        astList += Environment.NewLine + astName;
                    }
                }
                objWordDoc.Bookmarks["Top5_Threats_IED"].Range.Text = astList;

                Range srcRng = objWordTempl.Bookmarks["AssetBlock"].Range;

                Range astRange = objWordDoc.Bookmarks["AssetBlock"].Range;

                objWordDoc.Bookmarks["AssetBlock"].Delete();

                int alcount = al.Count;
                i = 0;

                foreach (Asset ast in al)
                {
                    i++;
                    srcRng.Copy();
                    astRange.Paste();

                    Stakeholder_populateAssetBlock("IED", ast, astRange, i);

                    astRange.Bookmarks["AssetBlock"].Delete();

                    astRange.SetRange(astRange.End, astRange.End);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Stakeholder_populateIEDThreatAssetList error: " + ex.Message);
            }
        }

        static void Stakeholder_populateVBIEDThreatAssetList(List<Asset> al)
        {
            try
            {
                string astList = "";
                string astName = "";
                int i = 1;
                foreach (Asset ast in al)
                {
                    astName = ast.getAssetName();
                    if (astList.Length == 0)
                    {
                        astList += astName;
                    }
                    else
                    {
                        astList += Environment.NewLine + astName;
                    }
                }
                astList += Environment.NewLine;
                astList += Environment.NewLine;

                objWordDoc.Bookmarks["Top5_Threats_VBIED"].Range.Text = astList;

                Range srcRng = objWordTempl.Bookmarks["AssetBlockVBIED"].Range;

                Range astRange = objWordDoc.Bookmarks["AssetBlockVBIED"].Range;

                objWordDoc.Bookmarks["AssetBlockVBIED"].Delete();

                int alcount = al.Count;
                i = 0;

                foreach (Asset ast in al)
                {
                    i++;
                    srcRng.Copy();
                    astRange.Paste();

                    Stakeholder_populateAssetBlock("VBIED", ast, astRange, i);

                    astRange.Bookmarks["AssetBlockVBIED"].Delete();

                    astRange.SetRange(astRange.End, astRange.End);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Stakeholder_populateVBIEDThreatAssetList error: " + ex.Message);
            }
        }

        static void populateAssetBlock(string riskType, Asset ast, Range InRng, int num)
        {
            Range rng;

            string assetName = ast.getAssetName();
            string list = "";

            if (riskType == "IED")
            {

                InRng.Bookmarks["IED_Threat_1_Title"].Range.Text = "13." + num.ToString() + " " + assetName;
                InRng.Bookmarks["IED_Asset_Name"].Range.Text = assetName;

                rng = InRng.Bookmarks["IED_Target_List"].Range;
                rng.Text = "";
                list = "";
                foreach (string tgt in ast.getTargetNames())
                {
                    list += tgt + Environment.NewLine;
                }
                rng.Text = list;
                rng.ListFormat.ApplyBulletDefault();
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                rng = InRng.Bookmarks.get_Item("FigureNum1").Range;
                rng.Text = FigureNum.ToString();

                rng = InRng.Bookmarks["IED_Asset_Map"].Range;
                rng.Text = "";
                addFigure(ast.getAssetMap(), assetName + " Target Areas", rng);

                InRng.Bookmarks["Asset_TSSE_Overview_IED"].Range.Text = "";//ast.getTSSEOverview();//QZ commented this part out per Lori's suggestion of removing the grey paragraph.
            }

            if (riskType == "VBIED")
            {
                InRng.Bookmarks["VBIED_Threat_1_Title"].Range.Text = "14." + num.ToString() + " " + assetName;
                InRng.Bookmarks["VBIED_Asset_Name"].Range.Text = assetName;

                rng = InRng.Bookmarks["VBIED_Target_List"].Range;
                rng.Text = "";
                list = "";
                foreach (string tgt in ast.getTargetNames())
                {
                    list += tgt + Environment.NewLine;
                }
                rng.Text = list;
                rng.ListFormat.ApplyBulletDefault();
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                rng = InRng.Bookmarks.get_Item("FigureNum2").Range;
                rng.Text = FigureNum.ToString();

                rng = InRng.Bookmarks["VBIED_Asset_Map"].Range;
                rng.Text = "";
                addFigure(ast.getAssetMap(), assetName + " Target Areas", rng);

                InRng.Bookmarks["Asset_TSSE_Overview_VBIED"].Range.Text = "";// ast.getTSSEOverview();//QZ commented this part out per Lori's suggestion of removing the grey paragraph.
            }
        }

        static void Stakeholder_populateAssetBlock(string riskType, Asset ast, Range InRng, int num)
        {
            Range rng;

            string assetName = ast.getAssetName();
            string list = "";

            if (riskType == "IED")
            {
                InRng.Bookmarks["IED_Threat_1_Title"].Range.Text = "10." + num.ToString() + " " + assetName;
                InRng.Bookmarks["IED_Asset_Name"].Range.Text = assetName;

                rng = InRng.Bookmarks["IED_Target_List"].Range;
                rng.Text = "";
                list = "";
                foreach (string tgt in ast.getTargetNames())
                {
                    list += tgt + Environment.NewLine;
                }
                rng.Text = list;
                rng.ListFormat.ApplyBulletDefault();
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                rng = InRng.Bookmarks.get_Item("FigureNum1").Range;
                rng.Text = FigureNum.ToString();

                rng = InRng.Bookmarks["IED_Asset_Map"].Range;
                rng.Text = "";
                addFigure(ast.getAssetMap(), assetName + " Target Areas", rng);

                InRng.Bookmarks["TSSE_Notes_IED"].Range.Text = "";// ast.getTSSEOverview();//QZ commented this part out per Lori's suggestion of removing the grey paragraph.
            }

            if (riskType == "VBIED")
            {
                InRng.Bookmarks["VBIED_Threat_1_Title"].Range.Text = "11." + num.ToString() + " " + assetName;
                InRng.Bookmarks["VBIED_Asset_Name"].Range.Text = assetName;

                rng = InRng.Bookmarks["VBIED_1_Target_List"].Range;
                rng.Text = "";
                list = "";
                foreach (string tgt in ast.getTargetNames())
                {
                    list += tgt + Environment.NewLine;
                }
                rng.Text = list;
                rng.ListFormat.ApplyBulletDefault();
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                rng = InRng.Bookmarks.get_Item("FigureNum2").Range;
                rng.Text = FigureNum.ToString();

                rng = InRng.Bookmarks["VBIED_Asset_Map"].Range;
                rng.Text = "";
                addFigure(ast.getAssetMap(), assetName + " Target Areas", rng);

                InRng.Bookmarks["TSSE_Notes_VBIED"].Range.Text = "";// ast.getTSSEOverview();//QZ commented this part out per Lori's suggestion of removing the grey paragraph.ast.getTSSEOverview();
            }

            Stakeholder_populate_target_area_group(riskType, ast, num);
        }


        static void populate_target_area_group(string appendix, List<Asset> al)
        {
            string tgtBlock = "";
            string riskType = "";

            if (appendix == "A") riskType = "IED";
            if (appendix == "B") riskType = "VBIED";

            if (riskType == "IED")
            {
                tgtBlock = "IED_Target_Area_Block";
            }

            if (riskType == "VBIED")
            {
                tgtBlock = "VBIED_Target_Area_Block";
            }

            Range srcRng = objWordTempl.Bookmarks[tgtBlock].Range;

            Range tgtRange = objWordDoc.Bookmarks[tgtBlock].Range;

            objWordDoc.Bookmarks[tgtBlock].Delete();

            srcRng.Copy();
            int tgtNum = 0;
            int astNum = 0;
            foreach (Asset ast in al)
            {
                astNum++;
                tgtNum = 0;
                foreach (Target tgt in ast.getTargetList())
                {
                    tgtRange.Paste();

                    tgtNum++;
                    populate_target_area(riskType, ast.getTerminalName(), tgt, tgtRange, astNum, tgtNum);
                    tgtRange.Bookmarks[tgtBlock].Delete();
                    tgtRange.SetRange(tgtRange.End, tgtRange.End);
                }
            }
        }

        static void Stakeholder_populate_target_area_group(string riskType, Asset ast, int astNum)
        {
            string tgtBlock = "";

            if (riskType == "IED")
            {
                tgtBlock = "IED_Target_Area_Block";
            }

            if (riskType == "VBIED")
            {
                tgtBlock = "VBIED_Target_Area_Block";
            }

            Range srcRng = objWordTempl.Bookmarks[tgtBlock].Range;

            Range tgtRange = objWordDoc.Bookmarks[tgtBlock].Range;

            objWordDoc.Bookmarks[tgtBlock].Delete();

            srcRng.Copy();
            int tgtNum = 0;
            foreach (Target tgt in ast.getTargetList())
            {
                tgtRange.Paste();

                tgtNum++;
                Stakeholder_populate_target_area(riskType, tgt, tgtRange, astNum, tgtNum);
                tgtRange.Bookmarks[tgtBlock].Delete();
                tgtRange.SetRange(tgtRange.End, tgtRange.End);
            }
        }

        static void populate_target_area(string riskType, string TerminalName, Target tgt, Range tgtR, int astNum, int tgtNum)
        {
            Range rng;
            object normaltype = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleNormal;


            if (riskType == "IED")
            {
                rng = tgtR.Bookmarks["IED_Target_Area_Title_1"].Range;
                rng.Text = "A." + astNum + "." + tgtNum + " " +
                            TerminalName + " - " + tgt.getTarget_Name();
                //added by QZ on July 8, 2014 so that the TableOfContents can pick up this entry
                rng.set_Style( WdBuiltinStyle.wdStyleHeading2 );
                rng.Font.Bold = 1;

                rng = tgtR.Bookmarks["IED_Target_Area_Blast_Image_1"].Range;
                rng.Text = "";
                addFigure(tgt.getTargetMap(), "Target Blast Image", rng);
                //rng.set_Style(normaltype);


                rng = tgtR.Tables[1].Cell(1, 1).Range.Bookmarks["IED_Target_Area_Justify_1"].Range;
                rng.Text = tgt.getJustify();
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                tgtR.Bookmarks["FigureNum3"].Range.Text = FigureNum.ToString();

                rng = tgtR.Bookmarks["IED_Target_Threat_Chart_1"].Range;
                addDeterGraph(tgt.getTarget_Name(), tgt.getDeter(), rng);
                tgtR.Bookmarks["IED_Target_Threat_Chart_1"].Delete();


                rng = tgtR.Bookmarks["IED_Target_Area_Image_1"].Range;
                rng.Text = "";
                addFigure2(tgt.getTargetFigure()[0], "Target Area Image", rng);


                rng = tgtR.Bookmarks["IED_1_OFC_List"].Range;
                rng.Text = GetOptions(tgt);
                rng.ListFormat.ApplyBulletDefault();
            }

            if (riskType == "VBIED")
            {
                rng = tgtR.Bookmarks["VBIED_Target_Area_Title_1"].Range;
                rng.Text = "B." + astNum + "." + tgtNum + " " +
                           TerminalName + " - " + tgt.getTarget_Name();
                //added by QZ on July 8, 2014 so that the TableOfContents can pick up this entry
                rng.set_Style(WdBuiltinStyle.wdStyleHeading2);
                //rng.set_Style(normaltype);
                rng.Font.Bold = 1;

                rng = tgtR.Bookmarks["VBIED_Target_Area_Blast_Image_1"].Range;
                rng.Text = "";
                addFigure(tgt.getTargetMap(), "Target Blast Image", rng);
                //rng.set_Style(normaltype);

                rng = tgtR.Tables[1].Cell(1, 1).Range.Bookmarks["VBIED_Target_Area_Justify_1"].Range;
                rng.Text = tgt.getJustify();
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                tgtR.Bookmarks["FigureNum4"].Range.Text = FigureNum.ToString();

                rng = tgtR.Bookmarks["VBIED_Target_Threat_Chart_1"].Range;
                addDeterGraph(tgt.getTarget_Name(), tgt.getDeter(), rng);
                tgtR.Bookmarks["VBIED_Target_Threat_Chart_1"].Delete();

                rng = tgtR.Bookmarks["VBIED_Target_Area_Image_1"].Range;
                rng.Text = "";
                addFigure2(tgt.getTargetFigure()[0], "Target Area Image", rng);

                rng = tgtR.Bookmarks["VBIED_1_OFC_List"].Range;
                rng.Text = GetOptions(tgt);
                rng.ListFormat.ApplyBulletDefault();
            }          
        }

        static void Stakeholder_populate_target_area(string riskType, Target tgt, Range tgtR, int astNum, int tgtNum)
        {
            Range rng;
            object normaltype = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleNormal; 

            if (riskType == "IED")
            {
                rng = tgtR.Bookmarks["IED_Target_Area_Title_1"].Range;
                rng.Text = "10." + astNum + "." + tgtNum + " " + tgt.getTarget_Name();
                //added by QZ on July 8, 2014 so that the TableOfContents can pick up this entry
                rng.set_Style(WdBuiltinStyle.wdStyleHeading2);
                rng.Font.Bold = 1;

                rng = tgtR.Bookmarks["IED_Target_Area_Blast_Image_1"].Range;
                rng.Text = "";
                addFigure(tgt.getTargetMap(), "Target Blast Image", rng);

                rng = tgtR.Tables[1].Cell(1, 1).Range.Bookmarks["IED_Target_Area_Justify_1"].Range;
                rng.Text = tgt.getJustify();
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                tgtR.Bookmarks["FigureNum3"].Range.Text = FigureNum.ToString();

                rng = tgtR.Bookmarks["IED_Target_Threat_Chart_1"].Range;
                addDeterGraph(tgt.getTarget_Name(), tgt.getDeter(), rng);
                tgtR.Bookmarks["IED_Target_Threat_Chart_1"].Delete();

                rng = tgtR.Bookmarks["IED_Target_Area_Image_1"].Range;
                rng.Text = "";
                addFigure2(tgt.getTargetFigure()[0], "Target Area Image", rng);


                rng = tgtR.Bookmarks["IED_1_OFC_List"].Range;
                rng.Text = GetOptions(tgt);
                rng.ListFormat.ApplyBulletDefault();
            }

            if (riskType == "VBIED")
            {
                rng = tgtR.Bookmarks["VBIED_Target_Area_Title_1"].Range;
                rng.Text = "11." + astNum + "." + tgtNum + " " + tgt.getTarget_Name();
                //added by QZ on July 8, 2014 so that the TableOfContents can pick up this entry
                rng.set_Style(WdBuiltinStyle.wdStyleHeading2);
                rng.Font.Bold = 1;

                rng = tgtR.Bookmarks["VBIED_Target_Area_Blast_Image_1"].Range;
                rng.Text = "";
                addFigure(tgt.getTargetMap(), "Target Blast Image", rng);
                //rng.set_Style(normaltype);

                rng = tgtR.Tables[1].Cell(1, 1).Range.Bookmarks["VBIED_Target_Area_Justify_1"].Range;
                rng.Text = tgt.getJustify();
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                tgtR.Bookmarks["FigureNum4"].Range.Text = FigureNum.ToString();

                rng = tgtR.Bookmarks["VBIED_Target_Threat_Chart_1"].Range;
                addDeterGraph(tgt.getTarget_Name(), tgt.getDeter(), rng);
                tgtR.Bookmarks["VBIED_Target_Threat_Chart_1"].Delete();

                rng = tgtR.Bookmarks["VBIED_Target_Area_Image_1"].Range;
                rng.Text = "";
                addFigure2(tgt.getTargetFigure()[0], "Target Area Image", rng);

                rng = tgtR.Bookmarks["VBIED_1_OFC_List"].Range;
                rng.Text = GetOptions(tgt);
                rng.ListFormat.ApplyBulletDefault();
            }
        }


        static string GetOptions(Target tgt)
        {
            string opts = "";


            foreach (string opt in tgt.getOptions())
            {
                opts += opt + "\n";
            }

            if (opts == "") opts = "None\n";
            return opts;
        }


        static void addFigure(Figure fig, string cap, Range rng)
        {
            try
            {
                rng.InsertParagraphAfter();
                if (fig.getFileName() == null) return;
                // Write the Figure "BLOB" to a temporary file
                string imgLink = Output + "\\" + fig.getFileName();
                FileStream fs = new FileStream(imgLink, FileMode.Create);
                BinaryWriter w = new BinaryWriter(fs);

                w.Write(fig.getFigure());
                w.Close();
                fs.Close();

                /* Insert the Image (which was written to a Temporary File) into the Document  */
                object tr = true;
                object fa = false;
                //clean up the image area--QZ
                foreach (Microsoft.Office.Interop.Word.InlineShape shp in rng.InlineShapes)
                    shp.Delete();

                Microsoft.Office.Interop.Word.InlineShape ils =
                    rng.InlineShapes.AddPicture(imgLink, ref fa, ref tr, System.Reflection.Missing.Value);

                // Set the Border around the Image to Red
                ils.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                ils.Range.Borders.OutsideColor = WdColor.wdColorRed;
                ils.Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth150pt;
                ils.Width = 460;//QZ
                ils.Height = 260;

                /* Insert the Figure's Caption */
                WdCaptionLabelID test = WdCaptionLabelID.wdCaptionFigure;
                //object CapTitle = cap;

                ils.Range.InsertCaption(test, " " + cap, System.Reflection.Missing.Value, WdCaptionPosition.wdCaptionPositionBelow, System.Reflection.Missing.Value);
                ils.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                //added the following two lines by QZ on July 8, 2014 to avoid the InLineShapes being inserted into the TableOfFigures
                Shape s = ils.ConvertToShape();
                s.WrapFormat.Type = WdWrapType.wdWrapTopBottom;
                ils.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                FigureNum++;

                /* Delete the Temporary File */
                System.IO.File.Delete(imgLink);
                rng.InsertParagraphAfter();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in addFigure: " + ex.Message);
            }
        }

        static void addFigure2(Figure fig, string cap, Range rng)
        {
            rng.InsertParagraphAfter();
            if (fig.getFileName() == null) return;
            // Write the Figure "BLOB" to a temporary file
            string imgLink = Output + "\\" + fig.getFileName();
            FileStream fs = new FileStream(imgLink, FileMode.Create);
            BinaryWriter w = new BinaryWriter(fs);

            w.Write(fig.getFigure());
            w.Close();
            fs.Close();

            /* Insert the Image (which was written to a Temporary File) into the Document  */
            object tr = true;
            object fa = false;
            Microsoft.Office.Interop.Word.InlineShape ils =
                rng.InlineShapes.AddPicture(imgLink, ref fa, ref tr, System.Reflection.Missing.Value);

            // Set the Border around the Image to Black
            ils.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            ils.Range.Borders.OutsideColor = WdColor.wdColorBlack;
            ils.Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth150pt;
            ils.Width = 460;
            ils.Height = 260;

            /* Insert the Figure's Caption */
            WdCaptionLabelID test = WdCaptionLabelID.wdCaptionFigure;
            //object CapTitle = cap;
            ils.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            ils.Range.InsertCaption(test, " " + cap, System.Reflection.Missing.Value, WdCaptionPosition.wdCaptionPositionBelow, System.Reflection.Missing.Value);
            //added the following two lines by QZ on July 8, 2014 to avoid the InLineShapes being inserted into the TableOfFigures
            Shape s = ils.ConvertToShape();
            s.WrapFormat.Type = WdWrapType.wdWrapTopBottom;
            FigureNum++;

            /* Delete the Temporary File */
            System.IO.File.Delete(imgLink);
            rng.InsertParagraphAfter();
        }

        static WdColorIndex getRAColor(int score)
        {
            WdColorIndex clr = WdColorIndex.wdGray25;

            if (score <= 29)
            {
                clr = WdColorIndex.wdGreen;
            }

            if (score >= 30 && score <= 69)
            {
                clr = WdColorIndex.wdYellow;
            }
            if (score >= 70)
            {
                clr = WdColorIndex.wdRed;
            }

            return clr;
        }

        static void insertTableRAT(ref string[, ,] scores)
        {
            Range rng;
            rng = objWordDoc.Bookmarks["Table_RAST"].Range;

            Table tbl = rng.Tables[1];
            tbl.AllowPageBreaks = false;//added by QZ on July 1, 2014

            while (tbl.Columns.Count > 3)
            {
                tbl.Columns.Last.Delete();
            }

            tbl.Cell(1, 2).Range.Text = airPort.getTerminalNames()[0].ToString(); // "Terminal 1";
            tbl.Cell(2, 2).Range.Text = "IED";
            tbl.Cell(2, 3).Range.Text = "VBIED";

            for (int i = 0; i < 6; i++)
            {
                tbl.Cell(3 + i, 2).Range.Text = scores[i, 0, 0];
                tbl.Cell(3 + i, 2).Range.Shading.BackgroundPatternColorIndex = getRAColor(Int32.Parse(scores[i, 0, 0]));

                if (i > 1)
                {
                    tbl.Cell(3 + i, 3).Range.Text = "N/A";
                }
                else
                {

                    tbl.Cell(3 + i, 3).Range.Text = scores[i, 0, 1];
                    tbl.Cell(3 + i, 3).Range.Shading.BackgroundPatternColorIndex = getRAColor(Int32.Parse(scores[i, 0, 1]));

                }
            }

            //rng.Copy();//commented out by QZ on June 23, 2014

            int numT = airPort.getTerminalNames().Count;
            int numC = 0;
            int maxT = 3;
            int currT = 0;

            int count = 0, kount = 0;

            while (numC < numT)
            {
                currT = maxT;

                if ((numT - numC) >= maxT)
                {
                    currT = maxT;
                }

                if ((numT - numC) < maxT)
                {
                    currT = numT - numC;
                }

                if (currT > numT) currT = numT;

                numC += currT;

                for (int j = 1; j < currT; j++)
                {

                    tbl.Columns.Add(Type.Missing);

                    tbl.Cell(1, tbl.Columns.Count).Range.Text = airPort.getTerminalNames()[j].ToString();//Terminals 2...(currT-1)
                    tbl.Cell(1, tbl.Columns.Count).Range.Font.Bold = 1;

                    for (int i = 0; i < 6; i++)
                    {
                        tbl.Cell(3 + i, tbl.Columns.Count).Range.Text = "";
                        tbl.Cell(3 + i, tbl.Columns.Count).Range.Shading.BackgroundPatternColorIndex = WdColorIndex.wdNoHighlight;

                    }

                    tbl.Cell(2, tbl.Columns.Count).Range.Text = "IED";

                    for (int i = 0; i < 6 && j < currT; i++)
                    {
                        tbl.Cell(3 + i, tbl.Columns.Count).Range.Text = scores[i, j, 0];
                        tbl.Cell(3 + i, tbl.Columns.Count).Range.Shading.BackgroundPatternColorIndex = getRAColor(Int32.Parse(scores[i, j, 0]));

                    }

                    tbl.Columns.Add(Type.Missing);

                    for (int i = 0; i < 6; i++)
                    {
                        tbl.Cell(3 + i, tbl.Columns.Count).Range.Text = "";
                        tbl.Cell(3 + i, tbl.Columns.Count).Range.Shading.BackgroundPatternColorIndex = WdColorIndex.wdNoHighlight;

                    }
                    tbl.Cell(2, tbl.Columns.Count).Range.Text = "VBIED";

                    for (int i = 0; i < 6 && j < currT; i++)
                    {
                        tbl.Cell(3 + i, tbl.Columns.Count).Range.Text = scores[i, j, 1];
                        tbl.Cell(3 + i, tbl.Columns.Count).Range.Shading.BackgroundPatternColorIndex = getRAColor(Int32.Parse(scores[i, j, 1]));

                    }
                }
                mergeCellCapRow(tbl, numT);

                if ((numT - numC) > 0)
                {
                    count = tbl.Rows[1].Cells.Count;

                    kount = (count - 1) / 2;

                    for (int i = 0; i < kount; i++)
                    {

                        count = tbl.Rows[1].Cells.Count - i;

                        tbl.Rows[1].Cells[count].Range.Text = "";
                        tbl.Rows[1].Cells[count - 1].Merge(tbl.Rows[1].Cells[count]);

                    }

                    tbl.Rows[1].Cells[1].Merge(tbl.Rows[2].Cells[1]);


                    rng.InsertParagraphAfter();

                    rng.SetRange(rng.End, rng.End);

                    rng.Paste();

                    tbl = rng.Tables[1];
                }
            }
            rng.Copy();//Added by QZ on June 23, 2014, for the call rng.Paste() later on

            // if there not enough room for national/regional columns,
            //  create a new table.
            if (numC + 2 > maxT)
            {
                rng.InsertParagraphAfter();

                rng.SetRange(rng.End, rng.End);

                rng.Paste();

                tbl = rng.Tables[1];
                tbl.AllowPageBreaks = false;//added by QZ on July 1, 2014

                tbl.Cell(1+8, 2).Range.Text = "National";
                tbl.Cell(1+8, 2).Range.Font.Bold = 1;

                for (int i = 0; i < 6; i++)
                {
                    tbl.Cell(3+8 + i, 2).Range.Text = "";
                    tbl.Cell(3+8 + i, 2).Range.Shading.BackgroundPatternColorIndex = WdColorIndex.wdNoHighlight;

                    tbl.Cell(3+8 + i, 3).Range.Text = "";
                    tbl.Cell(3+8 + i, 3).Range.Shading.BackgroundPatternColorIndex = WdColorIndex.wdNoHighlight;
                }

                //tbl.Columns.Add(Type.Missing);
                //tbl.Cell(1, tbl.Columns.Count).Range.Text = "Regional";
                //tbl.Cell(1, tbl.Columns.Count).Range.Font.Bold = 1;
                //tbl.Cell(2, tbl.Columns.Count).Range.Text = "IED";
                tbl.Cell(1+8, 3).Range.Text = "Regional";
                tbl.Cell(1+8, 3).Range.Font.Bold = 1;
                for (int i = 0; i < 6; i++)
                {
                    tbl.Cell(3+8 + i, 4).Range.Text = "";
                    tbl.Cell(3+8 + i, 4).Range.Shading.BackgroundPatternColorIndex = WdColorIndex.wdNoHighlight;

                    tbl.Cell(3+8 + i, 5).Range.Text = "";
                    tbl.Cell(3+8 + i, 5).Range.Shading.BackgroundPatternColorIndex = WdColorIndex.wdNoHighlight;
                }               

                //tbl.Cell(2, 4).Range.Text = "IED";
                //tbl.Columns.Add(Type.Missing);
                //tbl.Cell(2, tbl.Columns.Count).Range.Text = "VBIED";
            }

            else    // use the last table to add columns
            {
                for (int j = 1; j < 3; j++)
                {

                    tbl.Columns.Add(Type.Missing);

                    if (j == 1)
                    {
                        tbl.Cell(1, tbl.Columns.Count).Range.Text = "National";
                        tbl.Cell(1, tbl.Columns.Count).Range.Font.Bold = 1;
                    }
                    else if (j == 2)
                    {
                        tbl.Cell(1, tbl.Columns.Count).Range.Text = "Regional";
                        tbl.Cell(1, tbl.Columns.Count).Range.Font.Bold = 1;
                    }
                    tbl.Cell(2, tbl.Columns.Count).Range.Text = "IED";


                    tbl.Columns.Add(Type.Missing);
                    tbl.Cell(2, tbl.Columns.Count).Range.Text = "VBIED";

                }
                mergeCellCapRow(tbl, 2);
            }

            /*
            count = tbl.Rows[1].Cells.Count;

            kount = (count - 1) / 2;

            for (int i = 0; i < kount; i++)
            {

                count = tbl.Rows[1].Cells.Count - i;

                tbl.Rows[1].Cells[count].Range.Text = "";
                tbl.Rows[1].Cells[count - 1].Merge(tbl.Rows[1].Cells[count]);

            }

            tbl.Rows[1].Cells[1].Merge(tbl.Rows[2].Cells[1]);
            */
        }

        //This function was added by QZ on July 1, 2014
        static void mergeCellCapRow(Table tbl, int cols)
        {
            int tcnt = tbl.Rows[1].Cells.Count;
            int kount = (tcnt - 1) / 2;
            int cnt = 0;

            for (int i = 0; i < kount && cnt < cols; i++)
            {
               tcnt = tbl.Rows[1].Cells.Count - i;
               tbl.Rows[1].Cells[tcnt].Range.Text = "";
               tbl.Rows[1].Cells[tcnt - 1].Merge(tbl.Rows[1].Cells[tcnt]);
               cnt++;
            }
        }

        static void insertTableRAS(ref string[,] scores)
        {

            Range rng;
            rng = objWordDoc.Bookmarks["Table_RASS"].Range;

            Table tbl = rng.Tables[1];
            tbl.AllowPageBreaks = false;//added by QZ on July 1, 2014
            tbl.Cell(1, 2).Range.Text = airPort.getAirportAbbv();

            for (int i = 0; i < 5; i++)
            {
                tbl.Cell(3 + i, 2).Range.Text = scores[i, 0];
                tbl.Cell(3 + i, 2).Range.Shading.BackgroundPatternColorIndex = getRAColor(Int32.Parse(scores[i, 0]));

                tbl.Cell(3 + i, 3).Range.Text = scores[i, 1];
                tbl.Cell(3 + i, 3).Range.Shading.BackgroundPatternColorIndex = getRAColor(Int32.Parse(scores[i, 1]));
            }
            tbl.AllowPageBreaks = false;//added by QZ on July 1, 2014
        }



        /* ****************************************
         * Get the Row Number
         * This is to determine how many rows are
         * needed in the Threat Scenero Table
         * Since there are two different sets, IED
         *   and VIED, which run in parallel columns
         *   find out which has the greater number
         *   of entries and that will be the number
         *   of rows needed
         * 
         * ****************************************/
        static int getRowNum(List<ThreatS> lst)
        {
            int rtnval = 0;
            int iedCnt = 0;
            int vedCnt = 0;

            foreach (ThreatS ts in lst)
            {
                if (ts.getThreat_Id() == 1)
                {
                    iedCnt++;
                }
                else if (ts.getThreat_Id() == 2)
                {
                    vedCnt++;
                }
            }

            if (iedCnt > vedCnt)
            {
                rtnval = iedCnt;
            }
            else
            {
                rtnval = vedCnt;
            }
            return rtnval;
        }



        /* **************************************************
         * Get the Threat Color
         * - return the color based on the Detect_Minute_Id
         * 
         * **************************************************/
        static Microsoft.Office.Interop.Word.WdColor getThreatColor(int tid)
        {
            Microsoft.Office.Interop.Word.WdColor clr = Microsoft.Office.Interop.Word.WdColor.wdColorWhite;

            switch (tid)
            {
                case 1:
                    clr = Microsoft.Office.Interop.Word.WdColor.wdColorLightGreen;
                    break;
                case 2:
                    clr = Microsoft.Office.Interop.Word.WdColor.wdColorLightYellow;
                    break;
                case 3:
                    clr = Microsoft.Office.Interop.Word.WdColor.wdColorLightOrange;
                    break;
                case 4:
                case -1:
                    clr = Microsoft.Office.Interop.Word.WdColor.wdColorRed;
                    break;
            }

            return clr;
        }


        /* ***********************************************************************
         * Add the Deterrence Measure Graph into the Document for this Target Area
         * 
         *  Parameters: trgName is the Name of the Target Area (used in Caption)
         *              det is the DeterM Object used to populate the Graph
         *              rng is the Range used to position graph within Document
         * 
         * ***********************************************************************/
        static void addDeterGraph(string trgName, DeterM det, Microsoft.Office.Interop.Word.Range rng)
        {
            // Set Up the Word Positioning for Chart
            object novalue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Range wrdRng = rng;
            Microsoft.Office.Interop.Word.InlineShape objShape = rng.InlineShapes[1];
            Chart xlchart = rng.InlineShapes[1].Chart;
            xlchart.ChartData.Activate();

            xlchart.ChartStyle = 4;
            xlchart.HasLegend = false;
            xlchart.HasTitle = false;

            // Set the Labels for X and Y Axis
            Microsoft.Office.Interop.Word.Axis xAxis =
                                (Microsoft.Office.Interop.Word.Axis)xlchart.Axes(Microsoft.Office.Interop.Word.XlAxisType.xlCategory,
                                                                                 Microsoft.Office.Interop.Word.XlAxisGroup.xlPrimary);
            xAxis.HasTitle = true;
            xAxis.AxisTitle.Text = "Deterrence Measures";

            Microsoft.Office.Interop.Word.Axis yAxis =
                                (Microsoft.Office.Interop.Word.Axis)xlchart.Axes(Microsoft.Office.Interop.Word.XlAxisType.xlValue,
                                                                                 Microsoft.Office.Interop.Word.XlAxisGroup.xlPrimary);
            yAxis.MaximumScale = 100;
            yAxis.HasTitle = true;
            yAxis.AxisTitle.Text = "Percentage of AT Teams influenced by Deterrence Measure";


            // Set Up Excel Datasheet for Chart
            Microsoft.Office.Interop.Excel.Workbook wb = (Microsoft.Office.Interop.Excel.Workbook)xlchart.ChartData.Workbook;
            wb.Application.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized;
            //wb.Windows[1].Visible = false;
            //wb.Application.Visible = false;

            Microsoft.Office.Interop.Excel.Worksheet dataSheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range tblRng = dataSheet.get_Range("A1", "B12");
            Microsoft.Office.Interop.Excel.ListObject tbl = dataSheet.ListObjects[1];
            tbl.Resize(tblRng);

            try
            {
                // Populate the Chart Datasheet
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A2", novalue)).FormulaR1C1 = "CCTV";
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A3", novalue)).FormulaR1C1 = "Security";
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A4", novalue)).FormulaR1C1 = "Lighting";
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A5", novalue)).FormulaR1C1 = "Airport Employee";
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A6", novalue)).FormulaR1C1 = "Blast Walls";
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A7", novalue)).FormulaR1C1 = "Glass Mitigation";
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A8", novalue)).FormulaR1C1 = "Random Security";
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A9", novalue)).FormulaR1C1 = "Barriers";
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A10", novalue)).FormulaR1C1 = "Perimeter";
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A11", novalue)).FormulaR1C1 = "Signage";
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A12", novalue)).FormulaR1C1 = "General Public";
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("B2", novalue)).FormulaR1C1 = det.getNumCCTV();
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("B3", novalue)).FormulaR1C1 = det.getNumSec();
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("B4", novalue)).FormulaR1C1 = det.getNumLight();
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("B5", novalue)).FormulaR1C1 = det.getNumEmp();
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("B6", novalue)).FormulaR1C1 = det.getNumWall();
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("B7", novalue)).FormulaR1C1 = det.getNumGlass();
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("B8", novalue)).FormulaR1C1 = det.getNumRand();
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("B9", novalue)).FormulaR1C1 = det.getNumBarr();
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("B10", novalue)).FormulaR1C1 = det.getNumPerm();
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("B11", novalue)).FormulaR1C1 = det.getNumSign();
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("B12", novalue)).FormulaR1C1 = det.getNumPublic();

                // Add the Figure Caption for the Graph
                WdCaptionLabelID test = WdCaptionLabelID.wdCaptionFigure;
                object Title3 = " - " + trgName + " Deterrence Measures";
                objShape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                objShape.Range.InsertCaption(test, ref Title3, System.Reflection.Missing.Value, WdCaptionPosition.wdCaptionPositionBelow, System.Reflection.Missing.Value);

                FigureNum++;

                // Set the border color, style, and width of the Image
                objShape.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                //objShape.Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth150pt;//Commented out by QZ on June 13, 2014 because the this line always caused "value out of range" exception.
                objShape.Range.Borders.OutsideColor = WdColor.wdColorBlack;
            }
            catch (Exception graphEx) //Added the try-catch block by QZ on June 13, 2014 
            {
                Console.WriteLine("Error in Progrma.addDeterGraph(): " + graphEx.Message); 
            }
            wb.Close();
        }


        static void addScenarioGraph(Airport airp, Range rng)
        {
            object novalue = System.Reflection.Missing.Value;

            Chart xlchart = rng.InlineShapes[1].Chart;
            xlchart.ChartData.Activate();
            Microsoft.Office.Interop.Excel.Workbook wb = (Microsoft.Office.Interop.Excel.Workbook)xlchart.ChartData.Workbook;
            //wb.Application.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized;

            //wb.Windows[1].Visible = false;
            //wb.Application.Visible = false;

            int countAst = airp.getAssetList().Count;

            Microsoft.Office.Interop.Excel.Worksheet dataSheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            Microsoft.Office.Interop.Excel.Range tblRng = dataSheet.get_Range("A1", "C" + (countAst + 1).ToString());

            Microsoft.Office.Interop.Excel.ListObject tbl = dataSheet.ListObjects[1];
            tbl.Resize(tblRng);
            Microsoft.Office.Interop.Word.Axis xAxis =
                              (Microsoft.Office.Interop.Word.Axis)xlchart.Axes(Microsoft.Office.Interop.Word.XlAxisType.xlCategory,
                                                                               Microsoft.Office.Interop.Word.XlAxisGroup.xlPrimary);

            xAxis.HasTitle = true;
            xAxis.AxisTitle.Text = "Asset Name";

            xAxis.HasTitle = false;
            // Populate the Chart Datasheet


            int row = 1;
            foreach (Asset ast in airp.getAssetList())
            {
                row++;
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A" + row, novalue)).Value2 = ast.getAssetName();
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("B" + row, novalue)).Value2 = ast.getScenarios();
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("C" + row, novalue)).Value2 =
                                             Math.Round(ast.getScenariosDetectedPct() / 100.0, 2);

            }


            xlchart.Refresh();
            wb.Close();
            rng.InlineShapes[1].Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            rng.InlineShapes[1].Borders.OutsideLineWidth = WdLineWidth.wdLineWidth225pt;
            rng.InlineShapes[1].Borders.OutsideColor = WdColor.wdColorBlack;

        }


        static void addSituationalGraphOld(Airport airp, Range rng)
        {
            object novalue = System.Reflection.Missing.Value;

            Chart xlchart = rng.InlineShapes[1].Chart;
            xlchart.ChartData.Activate();
            Microsoft.Office.Interop.Excel.Workbook wb = (Microsoft.Office.Interop.Excel.Workbook)xlchart.ChartData.Workbook;
            wb.Application.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized;


            //             wb.Application.Visible = false;

            int countAst = airp.getAssetList().Count;

            Microsoft.Office.Interop.Excel.Worksheet dataSheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            Microsoft.Office.Interop.Excel.Range tblRng = dataSheet.get_Range("A1", "C" + (countAst + 1).ToString());

            Microsoft.Office.Interop.Excel.ListObject tbl = dataSheet.ListObjects[1];
            tbl.Resize(tblRng);
            Microsoft.Office.Interop.Word.Axis xAxis =
                              (Microsoft.Office.Interop.Word.Axis)xlchart.Axes(Microsoft.Office.Interop.Word.XlAxisType.xlCategory,
                                                                               Microsoft.Office.Interop.Word.XlAxisGroup.xlPrimary);

            xAxis.HasTitle = true;
            xAxis.AxisTitle.Text = "Asset Name";

            xAxis.HasTitle = false;

            int row = 1;
            foreach (Asset ast in airp.getAssetList())
            {
                row++;
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A" + row, novalue)).Value2 = ast.getAssetName();
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("B" + row, novalue)).Value2 = ast.getSituations();
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("C" + row, novalue)).Value2 =
                                             Math.Round(ast.getSituationsDetectedPct() / 100.0, 2);

            }

            xlchart.Refresh();
            wb.Close();
            rng.InlineShapes[1].Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            rng.InlineShapes[1].Borders.OutsideLineWidth = WdLineWidth.wdLineWidth225pt;
            rng.InlineShapes[1].Borders.OutsideColor = WdColor.wdColorBlack;

        }

        static void addSituationalGraph(Airport airp, Range rng)
        {
            int maxT = 3;
            int currT = 0;
            int numT = 0;
            int countT = airp.getTerminalNames().Count;
            
            if (countT < maxT) currT = countT;
            while ((countT - numT) > 0)
            {
                currT = maxT;
                if ((countT - numT) < maxT)
                {
                    currT = countT - numT;
                }

                numT += currT;

                object novalue = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.InlineShape objShape = rng.InlineShapes[1];
                Chart xlchart = rng.InlineShapes[1].Chart;

                //Chart xlchart = rng.InlineShapes.AddChart(xlchart2.ChartType, Type.Missing).Chart;

                xlchart.ChartData.Activate();
                Microsoft.Office.Interop.Excel.Workbook wb = (Microsoft.Office.Interop.Excel.Workbook)xlchart.ChartData.Workbook;
                wb.Application.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized;

                Microsoft.Office.Interop.Excel.Worksheet dataSheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

                Microsoft.Office.Interop.Excel.Range tblRng = dataSheet.get_Range("A1", "E" + (currT + 1).ToString());

                Microsoft.Office.Interop.Excel.ListObject tbl = dataSheet.ListObjects[1];
                tbl.Resize(tblRng);

                //Added the following xAxis & yAxis formatting lines by QZ on July 28, 2014
                Microsoft.Office.Interop.Word.Axis xAxis = (Microsoft.Office.Interop.Word.Axis)xlchart.Axes(Microsoft.Office.Interop.Word.XlAxisType.xlCategory,
                                                   Microsoft.Office.Interop.Word.XlAxisGroup.xlPrimary);
                Microsoft.Office.Interop.Word.Axis yAxis = (Microsoft.Office.Interop.Word.Axis)xlchart.Axes(Microsoft.Office.Interop.Word.XlAxisType.xlValue,
                                                   Microsoft.Office.Interop.Word.XlAxisGroup.xlPrimary);
                //xAxis.HasTitle = true;
                //xAxis.AxisTitle.Text = "Terminal Name";
                yAxis.MajorUnit = 2;
                yAxis.MinorUnit = 1;
                yAxis.HasMajorGridlines = true;
                yAxis.HasMinorGridlines = false;
                xlchart.HasTitle = true;
                xlchart.ChartTitle.Text = airp.getAirportName();

                int row = 1;
                //foreach (string tname in airp.getTerminalNames())
                for (int i = 0; i < currT; i++)
                {
                    string tname = airp.getTerminalNames()[i].ToString();
                    row++;
                    ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A" + row, novalue)).Value2 = tname;

                }

                row = 1;
                object situCht = null;//Added by QZ for null checking before the .ToString() is called, June 12, 2014
                for (int i = 0; i < currT; i++)
                {
                    row++;
                    for (int j = 0; j < 4; j++)
                    {
                        situCht = airp.getSituationalChartData()[i, j];
                        if (situCht != null)
                            ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range(Convert.ToChar(66 + j).ToString() + row, novalue)).Value2 = situCht.ToString();
                    }
                }

                xlchart.Refresh();
                wb.Close();

                WdCaptionLabelID test = WdCaptionLabelID.wdCaptionFigure;
                object Title3 = " Situational Awareness by Terminal ";

                objShape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                objShape.Range.InsertCaption(test, ref Title3, System.Reflection.Missing.Value, WdCaptionPosition.wdCaptionPositionBelow, System.Reflection.Missing.Value);

                FigureNum++;

                // If there are more than 3 terminals, there will be multiple charts one for 3 terminals
                if ((countT - numT) > 0)
                {
                    rng.Copy();
                    rng.SetRange(rng.End, rng.End);

                    rng.InsertParagraphAfter();

                    rng.Paste();
                }
            }        
        }

        static void addTargetSituationalGraph(string AssetName, List<Target> tgtList, Range rng)
        {
            object novalue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.InlineShape objShape = rng.InlineShapes[1];

            Chart xlchart = rng.InlineShapes[1].Chart;

            xlchart.ChartData.Activate();
            Microsoft.Office.Interop.Excel.Workbook wb = (Microsoft.Office.Interop.Excel.Workbook)xlchart.ChartData.Workbook;
            wb.Application.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized;

            int count = tgtList.Count;

            Microsoft.Office.Interop.Excel.Worksheet dataSheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            Microsoft.Office.Interop.Excel.Range tblRng = dataSheet.get_Range("A1", "E" + (count + 1).ToString());

            Microsoft.Office.Interop.Excel.ListObject tbl = dataSheet.ListObjects[1];

            tbl.Resize(tblRng);

            //Added the following xAxis & yAxis formatting lines by QZ on July 28, 2014
            Microsoft.Office.Interop.Word.Axis xAxis = (Microsoft.Office.Interop.Word.Axis)xlchart.Axes(Microsoft.Office.Interop.Word.XlAxisType.xlCategory,
                                               Microsoft.Office.Interop.Word.XlAxisGroup.xlPrimary);
            Microsoft.Office.Interop.Word.Axis yAxis = (Microsoft.Office.Interop.Word.Axis)xlchart.Axes(Microsoft.Office.Interop.Word.XlAxisType.xlValue,
                                               Microsoft.Office.Interop.Word.XlAxisGroup.xlPrimary);
            //xAxis.HasTitle = true;
            //xAxis.AxisTitle.Text = "Terminal Name";
            yAxis.MajorUnit = 1;
            yAxis.MinorUnit = 1;
            yAxis.HasMajorGridlines = true;
            yAxis.HasMinorGridlines = false;

            xlchart.HasTitle = true;
            xlchart.ChartTitle.Text = AssetName;

            int row = 1;
            foreach (Target tgt in tgtList)
            {
                row++;
                ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range("A" + row, novalue)).Value2 = tgt.getTarget_Name();

                for (int j = 0; j < 4; j++)
                {
                    ((Microsoft.Office.Interop.Excel.Range)dataSheet.Cells.get_Range(Convert.ToChar(66 + j).ToString() + row, novalue)).Value2 =
                            tgt.getSituationalChartData()[j].ToString();
                }

            }


            xlchart.Refresh();
            wb.Close();

            WdCaptionLabelID test = WdCaptionLabelID.wdCaptionFigure;

            object Title3 = " Situational Awareness Summary: " + AssetName;

            objShape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            objShape.Range.InsertCaption(test, ref Title3, System.Reflection.Missing.Value, WdCaptionPosition.wdCaptionPositionBelow, System.Reflection.Missing.Value);

            FigureNum++;
            //rng.InlineShapes[1].Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            //rng.InlineShapes[1].Borders.OutsideLineWidth = WdLineWidth.wdLineWidth225pt;
            //rng.InlineShapes[1].Borders.OutsideColor = WdColor.wdColorBlack;
        }

        static void populateAppendix(string appendix, Airport airp)
        {

            object normaltype = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleNormal;

            Range rng;

            if (appendix != "A") appendix = "B";

            rng = objWordDoc.Bookmarks.get_Item("Appendix" + appendix + "_Top5_Assets").Range;

            string astList = "";
            string astName = "";
            List<Asset> al = appendix == "A" ? airp.getIEDAssetList() : airp.getVBIEDAssetList();

            foreach (Asset ast in al)
            {
                astName = ast.getAssetName();

                if (astList.Length == 0)
                {
                    astList += astName;
                }
                else
                {
                    astList += Environment.NewLine + astName;

                }
            }
            astList += Environment.NewLine;

            rng.Text = astList;
            rng.HighlightColorIndex = WdColorIndex.wdNoHighlight;
            rng.ListFormat.ApplyNumberDefault();
            rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; // wdAlignParagraphCenter;

            populate_target_area_group(appendix, al);

            return;

            Range rngApxA = objWordDoc.Bookmarks.get_Item(appendix + "_TargetBlock").Range;
            Range rngCpy = rngApxA;

            //System.Windows.Forms.Clipboard.Clear();

            int tgtNum = 0;
            Console.WriteLine("***********");

            foreach (Asset ast in al)
            {
                string assetName = ast.getAssetName();

                Console.WriteLine(" >> " + ast.getAssetName());

                foreach (Target tgt in ast.getTargetList())
                {
                    tgtNum++;

                    Console.WriteLine(tgtNum + "  >>> " + tgt.getTarget_Name());

                    rngApxA.Copy();
                    rngCpy = rngApxA;

                    //populateTargetInfo(appendix, tgtNum, assetName, tgt, rngApxA);

                    rngCpy.SetRange(rngApxA.End, rngApxA.End);
                    rngCpy.Paste();
                    rngApxA = rngCpy;
                    System.Windows.Forms.Clipboard.Clear();
                    //break;
                }
                //break;
            }
            rngApxA.Text = "\n";
        }


        /* ***********************************************
         * Prints out all the parts of the Airport Object
         * << This is a DEBUGGING Tool >>
         * 
         *  -Everything is placed in the string targetList
         *   except for the Airport Name, Abbrev, and Date
         * -Then the Asset Name List
         * 
         * For each Asset in the Asset List
         * - The Asset name, options, target name list
         * - for each Target in the Asset's Target List
         *   * Target Name, Justification, Observations
         *   * List of DeterM
         *   * List of ThreatS
         * 
         * ***********************************************/
        static string printAirport(Airport airp)
        {

            //string assNames = "";
            //ArrayList al = airp.getAssetNames();
            //foreach (string i in al)
            //    assNames += i  + Environment.NewLine;

            List<Asset> assList = airp.getAssetList();
            string targetList = "";
            foreach (Asset ast in assList)
            {
                targetList += "ASSET NAME: " + ast.getAssetName() + Environment.NewLine;

                targetList += "     OPTIONS" + Environment.NewLine;
                ArrayList op = ast.getOptions();
                foreach (string tn in op)
                {
                    targetList += "     " + tn + Environment.NewLine;
                }

                targetList += "     TARGETS" + Environment.NewLine;
                List<Target> tl = ast.getTargetList();
                foreach (Target tn in tl)
                {
                    targetList += Environment.NewLine;
                    targetList += "             TARGET NAME: " + tn.getTarget_Name() + Environment.NewLine;
                    targetList += "             Target Justify: " + tn.getJustify() + Environment.NewLine;
                    targetList += "             Target Observations: " + Environment.NewLine;
                    ArrayList obv = tn.getObservation();
                    foreach (string ob in obv)
                    {
                        targetList += "                       " + ob + Environment.NewLine;
                    }
                    targetList += "             Target DeterM: " + Environment.NewLine;
                    DeterM detm = tn.getDeter();
                    targetList += "                       NumTeam: " + detm.getNumTeam() + Environment.NewLine +
                                  "                             1: " + detm.getNumCCTV() + "   2: " + detm.getNumSec() + "   3: " + detm.getNumLight() + Environment.NewLine +
                                  "                             4: " + detm.getNumEmp() + "   5: " + detm.getNumWall() + "   6: " + detm.getNumGlass() + Environment.NewLine +
                                  "                             7: " + detm.getNumRand() + "   8: " + detm.getNumBarr() + "   9: " + detm.getNumPerm() + Environment.NewLine +
                                  "                            10: " + detm.getNumSign() + "  11: " + detm.getNumPublic() + Environment.NewLine;

                    targetList += "             Target ThreatS: " + Environment.NewLine;
                    List<ThreatS> thr = tn.getThreat();
                    foreach (ThreatS tr in thr)
                    {
                        targetList += "                       ThreatId: " + tr.getThreat_Id() + " DetectId: " + tr.getDetect_Id() + " MinuteId: " + tr.getMinute_Id() +
                                      " ThreatDate: " + tr.getThreatDate() + Environment.NewLine;
                    }
                }

            }

            return "Name: " + airp.getAirportName() + Environment.NewLine +
                   "Abbv: " + airp.getAirportAbbv() + Environment.NewLine +
                   "Date: " + airp.getAirportDate().ToString("MM/dd/yyyy") + Environment.NewLine +
                   targetList;
        }




    }
}
