using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelTranslator
{
    class ExcelTranslator
    {
        private List<string> readingIssue = new List<string>();
        private List<Tuple<string, int>> marker;
        private Dictionary<string, double[]>[] testPersonCollection = new Dictionary<string, double[]>[4];
        private List<string> identifiers = new List<string>();

        private _Application exelApp;
        private Workbook wb;
        private Worksheet ws;
        private string[] headers = new string[] { "Systolic Pressure", "Diastolic Pressure", "Mean Pressure", "Heart rate", "Stroke Volume", "Left Ventricular Ejection Time", "Pulse Interval", "Maximum Slope", "Cardiac Output", "Total Peripheral Resistance Medical Unit", "Total Peripheral Resistance CGS" };
        private string[] headerExpansion = new string[] { " (Baseline Beginn->Ende)", " (Kaltwasser Beginn->Ende)", " (Kaltwasser->Schmerzschwelle)", " (Schmerzschwelle->Kaltwasser)" };

        /// <summary>
        /// Creates a list of all xls at the input Path. Also starts the excel Application and initializes the testPersionCollection.
        /// </summary>
        /// <param name="inputPath">The path where the excel files are</param>
        /// <returns>List of all xls files in folder</returns>
        public string[] getFileList(string inputPath)
        {
            exelApp = new Application();
            for (int i = 0; i < testPersonCollection.Length; i++)
            {
                testPersonCollection[i] = new Dictionary<string, double[]>();
            }
            return Directory.GetFiles(inputPath, "*.xls");
        }

        /// <summary>
        /// Reads the excelfile. Seatches for markers, builds marker pairs, calculates mean values and documents issues.
        /// </summary>
        /// <param name="filepath">the path of the file that is read at the moment</param>
        public void readExcel(string filepath)
        {
            wb = exelApp.Workbooks.Open(filepath);
            ws = wb.Worksheets[1];
            int rowHight = ws.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            object[,] excelArray = ws.Cells.get_Range("A1", "M" + rowHight).Value2;
            
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(ws);

            wb.Close();
            Marshal.ReleaseComObject(wb);


            // reading the identifier
            string id = excelArray[4, 1].ToString();
            if (id == null)
            {
                readingIssue.Add(filepath + " missing identifier in document");
                string[] splitData = filepath.Split('\\');
                splitData = splitData[splitData.Length - 1].Split('_');
                id = splitData[0];
            }
            else if (identifiers.Contains(id))
            {
                readingIssue.Add(filepath + " wrong ID in document or document name " + id);
                string[] splitData = filepath.Split('\\');
                splitData = splitData[splitData.Length - 1].Split('_');
                id = splitData[0];
            }

            // finding the markers in file
            double meanValue = 0;
            marker = new List<Tuple<string, int>>();
            
            for (int i = 10; i <= rowHight; i++)
            {
                if ((string)excelArray[i, 13] != null)
                {
                    string value = (string)excelArray[i, 13];
                    marker.Add(new Tuple<string, int>(value, i));
                }
            }

            if (!identifiers.Contains(id))
            {
                identifiers.Add(id);

                //finding marker pairs
                if (marker.Count > 1)
                {
                    bool blbl = false;
                    int blblst = 0;
                    int blblend = 0;
                    bool cwcw = false;
                    int cwcwst = 0;
                    int cwcwend = 0;
                    bool cwpt = false;
                    int cwptst = 0;
                    int cwptend = 0;
                    bool ptcw = false;
                    int ptcwst = 0;
                    int ptcwend = 0;
                    for (int i = 0; i < marker.Count; i++)
                    {
                        switch (marker.ElementAt(i).Item1)
                        {
                            case "Baseline Beginn/ Ende":
                                if (!blbl)
                                {
                                    blblst = marker.ElementAt(i).Item2;
                                    for (int j = i + 1; j < marker.Count; j++)
                                    {
                                        if (marker.ElementAt(j).Item1.Equals("Baseline Beginn/ Ende"))
                                        {
                                            blblend = marker.ElementAt(j).Item2;
                                            blblst = marker.ElementAt(i).Item2;
                                            blbl = true;
                                        }
                                    }
                                    if (!blbl)
                                    {
                                        readingIssue.Add(id + " missing a second \"Baseline Beginn / Ende\" marker");
                                    }
                                }
                                break;
                            case "Kaltwasser Beginn/ Ende":
                                if (!(ptcw && cwpt))
                                {
                                    for (int j = i + 1; j < marker.Count; j++)
                                    {
                                        if (marker.ElementAt(j).Item1.Equals("Kaltwasser Beginn/ Ende"))
                                        {

                                            if (!cwcw)
                                            {
                                                cwcwst = marker.ElementAt(i).Item2;
                                                cwcwend = marker.ElementAt(j).Item2;
                                                cwcw = true;
                                            }
                                            else
                                            {
                                                readingIssue.Add(id + " an additonal \"Kaltwasser Beginn/ Ende\" marker appeared.");
                                            }
                                            break;
                                        }
                                        else if (marker.ElementAt(j).Item1.Equals("Schmerzschwelle"))
                                        {
                                            if (!cwpt)
                                            {
                                                cwptst = marker.ElementAt(i).Item2;
                                                cwptend = marker.ElementAt(j).Item2;
                                                cwpt = true;
                                            }
                                            else
                                            {
                                                readingIssue.Add(id + " an additonal \"Kaltwasser Beginn/ Ende\" marker appeared.");
                                            }
                                            break;
                                        }

                                    }
                                }
                                else
                                {
                                    if (!cwcw)
                                    {
                                        cwcwst = cwptst;
                                        cwcwend = ptcwend;
                                        cwcw = true;
                                    }
                                    else
                                    {
                                        readingIssue.Add(id + " an additonal \"Kaltwasser Beginn/ Ende\" marker appeared.");
                                    }
                                }
                                break;
                            case "Schmerzschwelle":
                                if (!ptcw)
                                {
                                    ptcwst = marker.ElementAt(i).Item2;

                                    for (int j = i + 1; j < marker.Count; j++)
                                    {
                                        if (marker.ElementAt(j).Item1.Equals("Kaltwasser Beginn/ Ende"))
                                        {
                                            ptcwend = marker.ElementAt(i).Item2;
                                            ptcw = true;
                                        }
                                    }
                                    if (!ptcw)
                                    {
                                        readingIssue.Add(id + " missing a \"Kaltwasser Beginn/ Ende\" marker at the end");
                                    }
                                }
                                else
                                {
                                    readingIssue.Add(id + " an additonal \"Schmerzschwelle\" marker appeared.");
                                }
                                break;
                            default:
                                readingIssue.Add(id + " unknown marker \"" + marker.ElementAt(i).Item1 + "\"");
                                break;
                        }
                    }

                    //building mean values
                    for (int n = 0; n < testPersonCollection.Length; n++)
                    {
                        if (!testPersonCollection[n].ContainsKey(id))
                        {
                            testPersonCollection[n].Add(id, new double[11]);

                            if ((((blbl && n == 0) || (cwcw && n == 1)) || (cwpt && n == 2)) || (cwpt && n == 3))
                            {
                                for (int i = 2; i < 13; i++)
                                {
                                    for (int j = blblst; j <= blblend; j++)
                                    {
                                        meanValue += (double)excelArray[j, i];
                                    }
                                    meanValue = meanValue / (blblend - blblst);
                                    testPersonCollection[n][id][i - 2] = meanValue;
                                    meanValue = 0;
                                }
                            }
                            else
                            {
                                readingIssue.Add(id + " marker pair " + headerExpansion[n] + " not found");
                                for (int i = 2; i < 13; i++)
                                {
                                    testPersonCollection[n][id][i - 2] = -1;
                                }
                            }
                        }
                    }
                }
                else
                {
                    readingIssue.Add(id + " no marker found!");
                    for (int n = 0; n < 4; n++)
                    {
                        if (!testPersonCollection[n].ContainsKey(id))
                        {
                            testPersonCollection[n].Add(id, new double[11]);

                            for (int i = 2; i < 13; i++)
                            {
                                testPersonCollection[n][id][i - 2] = -1;
                            }

                        }
                    }
                }
            }
            else
            {
                readingIssue.Add(id + " ID duplicate found on " + filepath);
            }

        }

        /// <summary>
        /// Creates a file with Errors and issues that occured while reding the file
        /// </summary>
        /// <param name="path">the path the issue report will be created at</param>
        public void createIssueReport(string path)
        {
            string filename = path + @"\IssueReport.txt";
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }
            using (FileStream fs = File.Create(filename))
            {
                int ErrorCount = 1;
                foreach (string s in readingIssue)
                {
                    Byte[] text = new UTF8Encoding(true).GetBytes(ErrorCount + ": " + s + "\n");
                    fs.Write(text, 0, text.Length);
                    ErrorCount++;
                }
            }
        }

        /// <summary>
        /// Creates the transformed excel file on the path
        /// </summary>
        /// <param name="path">the path where the excelfile will be created at</param>
        public void createExcel(string path)
        {
            string filename = path + "\\Test_Person_Collection.xls";
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }

            Application cApp = new Application();
            cApp.Visible = false;
            cApp.DisplayAlerts = false;
            Workbook cWb;
            Worksheet cWs;


            cWb = cApp.Workbooks.Add(Type.Missing);
            cWs = (Worksheet)cWb.ActiveSheet;
            cWs.Name = "Collected Data";

            object[,] excelArray = new object[identifiers.Count + 1, 45];

            excelArray[0, 0] = "Identifier";
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 11; j++)
                {
                    excelArray[0, i * 11 + j + 1] = headers[j] + headerExpansion[i];
                }
            }

            int RowCounter = 1;
            foreach (string key in identifiers)
            {
                excelArray[RowCounter, 0] = key;
                for (int i = 0; i < 4; i++)
                {
                    if (testPersonCollection[i].ContainsKey(key))
                    {
                        for (int j = 0; j < 11; j++)
                        {
                            excelArray[RowCounter, (i * 11 + j + 1)] = testPersonCollection[i][key][j];
                        }
                    }
                }
                RowCounter++;
            }
            int dividend = 45;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            cWs.get_Range("A1", columnName + (identifiers.Count + 1)).Value2 = excelArray;

            cWb.SaveAs(path + "\\Test_Person_Data_Collection.xls", XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(cWs);

            cWb.Close();
            Marshal.ReleaseComObject(cWb);

            cApp.Quit();
            Marshal.ReleaseComObject(cApp);
        }
        
        /// <summary>
        /// Necessary clean up
        /// </summary>
        public void clean()
        {
            readingIssue.Clear();
            for (int i = 0; i < testPersonCollection.Length; i++)
            {
                testPersonCollection[i].Clear();
            }
            identifiers.Clear();
            if (exelApp != null)
            {
                exelApp.Quit();
                Marshal.ReleaseComObject(exelApp);
                exelApp = null;
            }
        }
    }
}

