using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
//using System.Runtime.InteropServices;
//using Microsoft.Office.Interop.Excel;
//using CommandLine;
//using CommandLine.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Extensions;

namespace ExcelManip {

    public class Feature : IEquatable<Feature> {
        public string FeatureName { get; set; }

        public int NumberOfTests { get; set; }

        public override string ToString() {
            //return "ID: " + NumberOfTests + "   Name: " + FeatureName;
            return FeatureName + " (" + NumberOfTests.ToString() + ")";
        }
        public override bool Equals(object obj) {
            if (obj == null) return false;
            Feature objAsFeature = obj as Feature;
            if (objAsFeature == null) return false;
            else return Equals(objAsFeature);
        }
        public override int GetHashCode() {
            return NumberOfTests;
        }
        public bool Equals(Feature other) {
            if (other == null) return false;
            return (this.NumberOfTests.Equals(other.NumberOfTests));
        }
    }

    class Program {

        /*
         * Main(args[])
         * Author: ageorge
         * 
         * in: (INPUT_FILENAME[|OUTPUT_FILENAME])
         * - creates a doc (OUTPUT_FILENAME) with an expanded shell of ALM test cases from INPUT_FILENAME
         */
        static void Main(string[] args) {
            string INPUT_FILENAME = string.Empty;
            INPUT_FILENAME = "ius8-summary.xlsx"; // breakpoints/troubleshooting
            string OUTPUT_FILENAME = string.Empty;
            string strOF = string.Empty;
            string PRODUCT = string.Empty;
            string OWNER = string.Empty;
            int argumentLength;

            WorksheetPart wsFeatureList;
            OrderedDictionary odAreas = new OrderedDictionary();

            /*
             * process arguments
             * INPUT_FILENAME|INPUT_FILENAME OUTPUT_FILENAME
             */
            argumentLength = args.Length;
            if (string.IsNullOrEmpty(INPUT_FILENAME) && argumentLength == 0) { // usage
                return;
            }

            if (argumentLength == 1 || argumentLength == 2) {
                INPUT_FILENAME = args[0];
                try {
                    OUTPUT_FILENAME = args[1];
                }
                catch (Exception ex) {
                    // doesn't exist or bad, stays string.Empty?
                }

                if (!INPUT_FILENAME.EndsWith(".xlsx") || (!string.IsNullOrEmpty(OUTPUT_FILENAME) && !OUTPUT_FILENAME.EndsWith(".xlsx"))) {
                    Console.Out.WriteLine("This only works with Excel 2007+ (.xlsx) spreadsheets.");
                    return;
                }
                else if (argumentLength == 1) {
                    OUTPUT_FILENAME = INPUT_FILENAME.Replace(".xlsx", "-alm-import.xlsx");
                }    
            }
            else if (!string.IsNullOrEmpty(INPUT_FILENAME) && argumentLength == 0) { // debug?
                OUTPUT_FILENAME = INPUT_FILENAME.Replace(".xlsx", "-alm-import.xlsx");
            }


            /*
             * read test case summary doc
             * reads INPUT_FILENAME
             * - base on input filename? output filename configurable?
             */
            using (SpreadsheetDocument docInput = SpreadsheetDocument.Open(INPUT_FILENAME, false)) {
                wsFeatureList = SpreadsheetReader.GetWorksheetPartByName(docInput, "Feature List");
                Row[] rows = wsFeatureList.Worksheet.Descendants<Row>().ToArray();
                string currArea = string.Empty;
                string prevArea = string.Empty;
                bool newArea = false;

                List<Feature> listFeatures = new List<Feature>();
                
                //Console.Out.WriteLine(rows.Length.ToString());

                /*
                * iterate through worksheet
                */
                for (int r = 4; r < rows.Length - 1; r++) {
                    Cell[] cells = rows[r].Elements<Cell>().ToArray();
                    string featureName = string.Empty;
                    string functionalArea = string.Empty;
                    string strNumManTCs = string.Empty;
                    string[] arrFeature;
                    int numManTCs; // = -1;

                    /*
                     * get features and number of cases
                     */
                    featureName = GetCellValue(cells[0]);
                    strNumManTCs = GetCellValue(cells[1]);

                    if (!int.TryParse(strNumManTCs, out numManTCs)) numManTCs = 1; // we assume at least 1 TC per feature
                    arrFeature = featureName.Split(new char[] { '\\' });
                    functionalArea = arrFeature[0];
                    currArea = functionalArea;

                    if (r == 4) prevArea = currArea; // set it for the first run
                    if (currArea != prevArea) newArea = true;

                    featureName = featureName.TrimStart(functionalArea.ToCharArray());
                    featureName = featureName.TrimStart(new char[] { '\\' });

                    /*
                     * add to previous area
                     */
                    if (newArea) {
                        // add features list to previous area
                        odAreas.Add(prevArea, new List<Feature>(listFeatures));

                        // clear features list
                        listFeatures.Clear();

                        // reset newArea/prevArea
                        newArea = false;
                        prevArea = currArea;
                    }

                    /*
                     * add feature to list
                     */
                    listFeatures.Add(new Feature() { FeatureName = featureName, NumberOfTests = numManTCs } );
                } // end rows iteration
            } // end reading doc




            /*
            * build ALM import workbook
            * - base on input filename? output filename configurable?
            */
            //using (SpreadsheetDocument docOutput = SpreadsheetDocument.Create(OUTPUT_FILENAME, SpreadsheetDocumentType.Workbook)) {
            SpreadsheetDocument docOutput = Excel.CreateWorkbook(OUTPUT_FILENAME);
            Worksheet worksheet;
            //SharedStringTablePart tablepart;
            //WorkbookStylesPart stylespart;

            //docOutput.AddWorkbookPart();
            //docOutput.WorkbookPart.Workbook = new Workbook();

            //tablepart = docOutput.WorkbookPart.AddNewPart<SharedStringTablePart>();
            //tablepart.SharedStringTable = new SharedStringTable();

            // iterate through all areas
            IDictionaryEnumerator enumAreas = odAreas.GetEnumerator();
            uint outputRow = 1;
            int numFeatures = -1;
            int numTests = -1;

            while (enumAreas.MoveNext()) {
                List<Feature> outputFeatures = (List<Feature>)enumAreas.Value;
                string area = enumAreas.Key.ToString();

                // create worksheet for current area
                Excel.AddWorksheet(docOutput, area);
                worksheet = docOutput.WorkbookPart.WorksheetParts.First().Worksheet;

                // iterate through all features
                numFeatures = outputFeatures.Count;
                foreach (Feature aFeature in outputFeatures) {
                    Console.Out.WriteLine(area + " " + aFeature);

                    // get number of tests
                    numTests = aFeature.NumberOfTests;
                    // ensure we have rest of TC data (product/status/owner/etc)
                        
                    // add rows (current feature, duplicated over number of tests)
                    for (int i = 0; i < numTests; i++) {
                        Excel.SetCellValue(docOutput, worksheet, 1, outputRow, aFeature.FeatureName, true);
                        outputRow++;
                    }
                }

                worksheet.Save();
            } // end iterating through areas, data should be in doc

            docOutput.Close();
            System.Diagnostics.Process.Start(OUTPUT_FILENAME);

            // styling

            // save doc

            // exit
                //return;
            //} // end creating doc
        } // end Main(args)

        /*
         * GetCellValue(Cell cell)
         * Author: saarp (http://stackoverflow.com/a/13202816/1454048)
         * 
         * in: Cell object
         * out: string value of Cell contents
         */
        public static string GetCellValue(Cell cell) {
            if (cell == null)
                return null;
            if (cell.DataType == null)
                return cell.InnerText;

            string value = cell.InnerText;
            switch (cell.DataType.Value) {
                case CellValues.SharedString:
                    // For shared strings, look up the value in the shared strings table.
                    // Get worksheet from cell
                    OpenXmlElement parent = cell.Parent;
                    while (parent.Parent != null && parent.Parent != parent
                            && string.Compare(parent.LocalName, "worksheet", true) != 0) {
                        parent = parent.Parent;
                    }
                    if (string.Compare(parent.LocalName, "worksheet", true) != 0) {
                        throw new Exception("Unable to find parent worksheet.");
                    }

                    Worksheet ws = parent as Worksheet;
                    SpreadsheetDocument ssDoc = ws.WorksheetPart.OpenXmlPackage as SpreadsheetDocument;
                    SharedStringTablePart sstPart = ssDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    // lookup value in shared string table
                    if (sstPart != null && sstPart.SharedStringTable != null) {
                        value = sstPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                    }
                    break;

                //this case within a case is copied from msdn. 
                case CellValues.Boolean:
                    switch (value) {
                        case "0":
                            value = "FALSE";
                            break;
                        default:
                            value = "TRUE";
                            break;
                    }
                    break;
            }

            return value;
        }



    }
}
