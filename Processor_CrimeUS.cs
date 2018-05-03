using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;


namespace BG.Processors
{
    class Processor_CrimeUS : ProcessorBase
    {
        public Processor_CrimeUS() : base("Crime")
        {}

        private string fileId = "Table-2";

        public override ParsingReturnStructure RunScraper()
        {
            var fileName = DownloadFile(fileId);    //File is downloaded
            var pt = CreateParsedTable(fileName);   //File is formatted from excel workbook into parsed table
            var prsList = ReadData(pt, fileId); //Data is read from parsed table and formatted into prs
            DumpParsingReturnStructure(prsList, fileId);
            var multiPRS = new ParsingReturnStructure();
            foreach (var prs in prsList)    //Once prslist has been dumped each prs in the list is appended to the multiPRS 
            {
                multiPRS.Append(prs);
            }
            multiPRS.NeumKey = _nk;
            return multiPRS;
        }

        private string DownloadFile(string fileId)
        {
            
            try
            {
                var uri = new Uri($"https://ucr.fbi.gov/crime-in-the-u.s/2016/crime-in-the-u.s.-2016/tables/table-2/table-2.xls/output.xls");   //Page in which data is being pulled is inspected for relevant URI string, placed into a variable
                var localFile = $"{DataFolder}{DateTime.Now.ToString("yyyyMMdd - HHmmss.")}{fileId}.xls";   //Formats the Excel file for storage
                using (WebClient webClient = new WebClient())
                {
                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;    //Make sure the correct security protocol is used when downloading file
                    webClient.DownloadFile(uri, localFile); //File is downloaded 
                }
                return localFile;
            }
            catch(Exception ex)
            {
                Debug.Assert(false, ex.Message);    //Throws error in case of incorrect URI or security protocol
                return null;
            }
        }

        private ParsedTable CreateParsedTable(string ExcelFile) //File is downloaded, put into excel workbook format, workbook put into a parsed Table a table defined by another programmer
        {
            var fstream = new FileStream(ExcelFile, FileMode.Open);
            var workbook = new Aspose.Cells.Workbook(fstream);
            var sheet = workbook.Worksheets[0];
            var pt = new ParsedTable(sheet);
            fstream.Close();
            return pt;
        }
        private class Regions //Regions class for making region objects that specficy the regions name and its starting row
        {
            public string RegionName { get; set; }
            public int StartingRow { get; set; }
            public Regions(string regionName, int startingRow)
            {
                RegionName = regionName;
                StartingRow = startingRow;
            }
        }
        private class Crimes    //Crimes class for making crime objects that specficy the crime name and its starting row
        {
            public string CrimeName { get; set; }
            public int StartingCol { get; set; }
            public Crimes(string crimeName, int startingCol)
            {
                CrimeName = crimeName;
                StartingCol = startingCol;
            }
        }
        private class Footnotes //Footnoes class for making footnote objects that specficy the footnote id and the footnote
        {
            public int FootnoteId { get; set; }
            public string Value { get; set; }
            public Footnotes(int footnoteId, string value)
            {
                FootnoteId = footnoteId;
                Value = value;
            }
        }
        private string FormatString(string input)   //Formats the string removing any next line characters and non-letters
        {
            string output = input;
            output = Regex.Replace(output, "\n", string.Empty);
            output = Regex.Replace(output, "[^a-zA-Z]", string.Empty);
            
            return output;
        }

        private List<int> FormatFootnoteNumbers(string input)
        {
            List<int> footnoteNum = new List<int>();
            string[] numbers = Regex.Split(input, @"\D+");
            foreach (string value in numbers)
            {
                if (!string.IsNullOrEmpty(value))
                {
                    int i = int.Parse(value);
                    footnoteNum.Add(i);
                }
            }
            return footnoteNum;
        }
        private List<ParsingReturnStructure> ReadData(ParsedTable pt, string fileId)
        {
            PeriodTypes2017 frequency = PeriodTypes2017.Annual;
            List<Regions> regions = new List<Regions>();
            List<Crimes> crimes = new List<Crimes>();
            List<Footnotes> footnotes = new List<Footnotes>();
            List<string> metaData = new List<string>();
           
            List<ParsingReturnStructure> prsList = new List<ParsingReturnStructure>();  //Since there is a table made per region, instead of retruning one parsing return structure we are returning a list of them
            bool underArea = false;
            int categoryRow = 0;
            for (int row = 0; row <= 213; row++)    //Loop through the all the rows on sheet
            {
                int num;
                if (row <= 202 && pt.Vals[row, 0] != null && underArea == true) //If excel cell is not null and the range of rows occurs before footnotes
                {
                    regions.Add(new Regions(pt.Vals[row, 0].ToString(), row));
                }
                if (pt.Vals[row, 0] != null && pt.Vals[row, 0].ToString() == "Area") //If row is under the Area topic, underArea is true and indicates the starting row where regions start
                {
                    underArea = true;
                    categoryRow = row;
                }
                
                if (pt.Vals[row, 0] != null && Int32.TryParse(pt.Vals[row, 0].ToString().Substring(0,1), out num))  //Puts footnote id value and footnote into footnote object
                {
                    footnotes.Add(new Footnotes(num, pt.Vals[row, 0].ToString().Substring(1, pt.Vals[row, 0].ToString().Length-1)));
                }
                if (pt.Vals[row, 0] != null && pt.Vals[row, 0].ToString().Substring(0, 4) == "NOTE")
                {
                    metaData.Add(pt.Vals[row, 0].ToString());
                }

            }
            for (int col = 2; col < 22; col++)  //iterates through all of crime columns and if it is not an empty cell adds to the crime object
            {
                if (pt.Vals[categoryRow, col] != null)
                {
                    crimes.Add(new Crimes(pt.Vals[categoryRow, col].ToString(), col));
                }
            } 
            
            foreach (var region in regions) //Iterates through all of the regions, with a nested loop going through all of the crimes for each region
            {
                var dataPointList = new DataPointList();    //The Datapoint List stores a list of datapoints
                ParsingReturnStructure prs = new ParsingReturnStructure();  //A parsing return structure's datapoint value adds datapoint list
                List<int> regionFootnotes = new List<int>();
                if (region.RegionName.Any(c => char.IsDigit(c)))    //Determines if region name contains a footnote, if so adds footnote id number
                {
                    regionFootnotes = FormatFootnoteNumbers(region.RegionName);
                }
                
                foreach (var crime in crimes)
                {
                    List<int> crimeFootnotes = new List<int>();
                    if (crime.CrimeName.Any(c => char.IsDigit(c)))  //Determines if crime name contains a footnote, if so adds footnote id number
                    {
                        crimeFootnotes = FormatFootnoteNumbers(crime.CrimeName);
                    }
                    var neumonic = SeriesToNeumonic(FormatString(region.RegionName) + " - " + FormatString(crime.CrimeName));   //Formats the neumonic name for the datapoint appends the region name to the crime name
                    Guid seriesId = _nk.GetValue(neumonic);
                    _nk.AddSeriesName(seriesId, region.RegionName);
                  /* foreach (var regionNote in regionFootnotes)    //this commented out section will add the footnotes into the into the prslist once other code that manipulates footnotes is changed
                    {
                        foreach (var note in footnotes)
                        {
                            if (regionNote == note.FootnoteId)
                            {
                                _nk.AddFootnote(note.Value, seriesId);
                            }
                        }
                    }
                    foreach (var crimeNote in crimeFootnotes)
                    {
                        foreach (var note in footnotes)
                        {
                            if (crimeNote == note.FootnoteId)
                            {
                                _nk.AddFootnote(note.Value, seriesId);
                            }
                        }
                    }*/
                    int numLoops = 1;
                    if (crime.CrimeName.Contains("Population")) //Population category doesnt contain extra column pertaining to the rate of population committing a crime by 100k
                    {
                        numLoops = 0;
                    }
                    for (int col = crime.StartingCol; col <= crime.StartingCol + numLoops; col++)
                    {
                        string tempDate = "";

                        for (int row = region.StartingRow; row <= region.StartingRow + 1; row++)
                        {
                            if (row == region.StartingRow)  //Dates only have year values must be modified to into a DateTime format
                            {
                                tempDate = "12/31/2015";
                            }
                            else if (row == region.StartingRow + 1)
                            {
                                tempDate = "12/31/2016";
                            }
                           
                            var dp = new DataPoint2017(DateTime.Parse(tempDate), frequency, decimal.Parse(pt.Vals[row, col]));  //Datapoint value takes to values a Datetime value and decimal number value
                            dp.Neum = neumonic; //Data point takes a neumonic value as well as the parent series id, same as neumonic value unless datapoint has a parent
                            dp.ParentSeriesId = seriesId;
                            dataPointList.AddPoint(dp); //Data point is added to the datapoint list
                        }
                        if (numLoops == 1)  //if on second colummn of crime there is a given rate of crime per a 100k people
                        {
                            neumonic = SeriesToNeumonic(neumonic + "(RatePer100k)");
                            seriesId = _nk.GetValue(neumonic);
                            _nk.AddSeriesName(seriesId, region.RegionName);
                        }
                    }

                    BGTableInfo tableHere = new BGTableInfo();  //BGTableInfo contains relevant information for adding prs into the database at a later date
                    tableHere.TableName = region.RegionName;
                    prs.TableInfos.Add(tableHere);
                    BGTableLineInformation tableLineInfo = new BGTableLineInformation();
                    tableLineInfo.linelabel = region.RegionName;
                    tableLineInfo.tablelineindents = 0;
                    tableLineInfo.tablelinenum = 1;
                    tableLineInfo.objectID = seriesId;
                    tableHere.Add(tableLineInfo);
                }
                
                prs.DataPoints.AddList(dataPointList);  //prs adds the data point list that has accumulated all of the data points for the per one region
                prsList.Add(prs);   // prs for one region is added prsList, prsList is what is returned 
            }
            return prsList;
        }

        private string SeriesToNeumonic(string seriesName)
        {
            //shorten a series name into an understandable neumonic
            //TODO - make this better
            return seriesName.Replace(" ", "");
        }

        private bool DumpParsingReturnStructure(List<ParsingReturnStructure> prsList, string fileId)
        {
            //dump the prs just so we can check it

            const string Delimiter = ",";
            var stream = new StreamWriter($"{DataFolder}dump.{fileId}.csv");

            foreach (var prs in prsList)
            {
                stream.WriteLine("Date,Neum,Parent,Value");
                for (int i = 0; i < prs.DataPoints.Values.Count; i++)
                {
                    var dp = prs.DataPoints.Values[i];
                    stream.Write(dp.EndDate);
                    stream.Write(Delimiter);
                    stream.Write(dp.Neum);
                    stream.Write(Delimiter);
                    stream.Write(dp.ParentSeriesId);
                    stream.Write(Delimiter);
                    stream.Write(dp.Val);
                    stream.Write("\r\n");
                }
            }
            stream.Flush();
            stream.Close();

            return true;
        }

    }
}
