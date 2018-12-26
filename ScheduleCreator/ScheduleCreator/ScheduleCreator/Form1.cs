using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace ScheduleCreator
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// Automatically assigns values and calls functions as soon as the Form starts up.
        /// Meant to streamline the debugging process
        /// </summary>
        /// <param name="active"></param>
        /// <param name="autoExport"></param>
        /// <param name="forceOutputRawData"></param>
        private void Debuggery(bool active= false, bool autoExport = false, bool forceOutputRawData = false)
        {
            //If Debuggery's enabled
            if (active)
            {
                //Define excel destination
                textBox2.Text = "C:\\Users\\Mike\\Desktop";

                try
                {
                    //Read data and put it on the screen
                    StreamReader quickRead = new StreamReader("C:\\Users\\Mike\\Desktop\\Data.txt");
                    textBox1.Text = quickRead.ReadToEnd();
                    quickRead.Close();

                    //Parse the data that was put on the screen
                    ParseAndDisplayTable();

                    //Export automatically if specified
                    if(autoExport)
                    {
                        GenerateExcelFile(TableData, (forceOutputRawData) ? true : checkBox1.Checked);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message+" : "+ex.StackTrace);
                }
            }
        }

        /// <summary>
        /// Values for every weekday
        /// </summary>
        private enum DayValue {Monday = 0, Tuesday = 2400, Wednesday = 4800, Thursday = 7200, Friday = 9600};

        /// <summary>
        /// The sum of all the credits owed for each class
        /// </summary>
        private int creditSum = 0;

        /// <summary>
        /// Day string values to be accessed by a loop
        /// </summary>
        private string[] DayStrings = new string[5] {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday" };

        /// <summary>
        /// Raw data entries
        /// </summary>
        private List<Entry> entries;

        /// <summary>
        /// Array element to store the data
        /// </summary>
        private class Entry
        {
            //Monday = 0, T = 2400, W = 4800, T = 7200, F = 9600
            public int day1weight = 0;
            public int day2weight = 0;

            //Is a duplicate entry representing the second day
            public bool day2Entry;

            //Represents the staring military time 
            public int timeValue = 0;
            
            //0 = ClassName 
            //1 = Professor 
            //2 = Email 
            //3 = StartEndDate 
            //4 = Day1  
            //5 = Day2  
            //6 = Time  
            //7 = Location  
            //8 = Credits 
            public string[] entryData = new string[9];

            public void CopyData(ref Entry EntryToCopyTo)
            {
                EntryToCopyTo.entryData = this.entryData;
                EntryToCopyTo.timeValue = this.timeValue;
                EntryToCopyTo.day1weight = this.day1weight;
                EntryToCopyTo.day2weight = this.day2weight;
            }
        }

        /// <summary>
        /// Represents a range to find a given substring.
        /// Should be put into an array with a starting marker followed by not a starting marker
        /// </summary>
        private class ParseMarker
        {
            //Target char to look for
            public char target;

            //Extra offset for the front of the substring
            public int StartCharOffset = 0;

            //Offset the index of the char we find either to the left or right (Represents the end of the substring we're looking for)
            public int EndingCharOffset = 0;

            //If it's a starting marker assign this char index to left index
            public bool startingMarker = false;

            //If it's an end marker then substring the value according to the last ParseMarker that was a startingMarker
            public bool endMarker = false;

            //If we don't find it and we don't have to parse the string, skip it
            public bool optional = false;
        }

        /// <summary>
        /// Table for storing schedule data when it's been successfully parsed.
        /// </summary>
        private DataTable TableData;

        /// <summary>
        /// Colors representing Monday (0) through Friday (4)
        /// </summary>
        private Color[] dayColors = 
        {
            Color.FromArgb(1,208,224,227),
            Color.FromArgb(1,207,226,243),
            Color.FromArgb(1,217,210,233),
            Color.FromArgb(1,234,209,220),
            Color.FromArgb(1,230,184,175)
        };

        /// <summary>
        /// ParseMarkers meant for parsing out schedule data.
        /// (Assuming that it's in the format that I have specified)
        /// If YOUR format is different you can create your own using another set of Parse markers to define your data format.
        /// If you defined your own set make sure to update 'comboBox1' so you can select it and use it.
        /// </summary>
        private ParseMarker[] StandardParseMarkers = new ParseMarker[] {
            new ParseMarker{target=')', EndingCharOffset = 2, startingMarker = true}, //0 Start marker for class name
            new ParseMarker{target='\n', EndingCharOffset = -1, startingMarker = true, endMarker = true}, //1 Start marker for Professor's name / end marker for class name
            new ParseMarker{target='\n', EndingCharOffset = -1, startingMarker = true, endMarker = true}, //2 Start marker for Email  / end marker for Professor's name
            new ParseMarker{target='\n', EndingCharOffset = -1, startingMarker = true, endMarker = true}, //3 Start marker for Start/End date  / end marker for Email
            new ParseMarker{target=' ', endMarker = true}, //4 End marker for Start/End date
            new ParseMarker{target=' ', EndingCharOffset = 1, startingMarker = true}, //5 Start marker for Day1
            new ParseMarker{target=',', endMarker = true, optional = true}, //6 End marker for Day1 optional
            new ParseMarker{target=' ', EndingCharOffset = 1, startingMarker = true, optional = true}, //7 start for Day2 optional 
            new ParseMarker{target=' ', startingMarker = true, endMarker = true}, //8 End marker for Day2 (Day1 if the previous wasn't taken) / start marker for time
            new ParseMarker{target=',', StartCharOffset = 1, EndingCharOffset = -1, startingMarker = true, endMarker = true}, //9 End marker for time 
            new ParseMarker{target='\n', StartCharOffset = 1, EndingCharOffset = -2, startingMarker = true, endMarker = true}, //10 End marker for location 
            new ParseMarker{target='\n', StartCharOffset = -1, startingMarker = true ,endMarker = true} //11 End marker for credits
        };

        /// <summary>
        /// When the form is initialized
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            SetupUI();
            Debuggery();
        }

        /// <summary>
        /// Disables a few things and sets up objects to make things look right
        /// </summary>
        private void SetupUI()
        {
            dataGridView1.Visible = false;
            button3.Enabled = false;
            button4.Visible = false;
            comboBox1.SelectedIndex = 0;
        }

        /// <summary>
        /// Generate & Display Table Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button1_Click(object sender, EventArgs e)
        {
            ParseAndDisplayTable();
        }

        /// <summary>
        /// Browse Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            folderBrowser.ShowDialog();
            textBox2.Text = folderBrowser.SelectedPath;
        }

        /// <summary>
        /// Export To Excel Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button3_Click(object sender, EventArgs e)
        {
            if (textBox2.Text.Length >= 4)
            {
                GenerateExcelFile(TableData, checkBox1.Checked);
            }
        }

        /// <summary>
        /// Clear Table Data Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button4_Click(object sender, EventArgs e)
        {
            TableData = null;
            dataGridView1.DataSource = null;
            dataGridView1.Visible = false;

            label1.Text = "Raw Data:";

            button1.Enabled = true;
            button2.Enabled = false;
            textBox2.Enabled = false;
            button3.Enabled = false;
            button4.Visible = false;
            button5.Enabled = true;
            comboBox1.Enabled = true;
        }

        /// <summary>
        /// Load Data by File Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.ShowDialog();

            try
            {
                StreamReader quickRead = new StreamReader(openFileDialog.FileName);

                textBox1.Text = quickRead.ReadToEnd();

                quickRead.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message+" : "+ex.StackTrace);
            }
        }

        /// <summary>
        /// Recursive function.
        /// Parses a given string, returns true if successful, false if not.
        /// Takes the following references: 
        /// datastring, left and right index, 'Entry' object to populate the data, 'Entry' data index, Parse Markers and parseMarker index
        /// </summary>
        /// <param name="dataString"></param>
        /// <param name="leftIndex"></param>
        /// <param name="rightIndex"></param>
        /// <param name="theNewEntryToFill"></param>
        /// <param name="EntryDataIndex"></param>
        /// <param name="parsingMarkers"></param>
        /// <param name="markerIndex"></param>
        /// <returns></returns>
        private bool SingleCharDataParser(ref string dataString, ref int leftIndex, ref int rightIndex, ref Entry theNewEntryToFill, ref int EntryDataIndex, ref ParseMarker[] parsingMarkers, ref int markerIndex)
        {
            while (rightIndex != dataString.Length && theNewEntryToFill.entryData[8] == null)
            {
                //If we found the target char
                if (dataString[rightIndex] == parsingMarkers[markerIndex].target)
                {
                    //If this is a start marker then the next marker should be the end marker
                    if (parsingMarkers[markerIndex].startingMarker)
                    {
                        if (parsingMarkers[markerIndex].endMarker)
                        {

                            //Substring it
                            theNewEntryToFill.entryData[EntryDataIndex++] = dataString.Substring(leftIndex + parsingMarkers[markerIndex].StartCharOffset, rightIndex - leftIndex + parsingMarkers[markerIndex].EndingCharOffset);

                            //Offset right index
                            rightIndex += Math.Abs(parsingMarkers[markerIndex++].EndingCharOffset);
                            leftIndex = rightIndex++;

                            SingleCharDataParser(ref dataString, ref leftIndex, ref rightIndex, ref theNewEntryToFill, ref EntryDataIndex, ref parsingMarkers, ref markerIndex);
                        }
                        else
                        {
                            //Shift the left index to where the starting marker was found (Taking into account the offset) 
                            rightIndex += Math.Abs(parsingMarkers[markerIndex++].EndingCharOffset);
                            leftIndex = rightIndex++;

                            SingleCharDataParser(ref dataString, ref leftIndex, ref rightIndex, ref theNewEntryToFill, ref EntryDataIndex, ref parsingMarkers, ref markerIndex);
                        }
                    }
                    //If it's not a starting marker (It's an end marker)
                    else if (parsingMarkers[markerIndex].endMarker)
                    {

                        //Substring it
                        theNewEntryToFill.entryData[EntryDataIndex++] = dataString.Substring(leftIndex + parsingMarkers[markerIndex].StartCharOffset, rightIndex - leftIndex + parsingMarkers[markerIndex].EndingCharOffset);

                        //Offset right index
                        rightIndex += Math.Abs(parsingMarkers[markerIndex++].EndingCharOffset);
                        leftIndex = rightIndex++;

                        SingleCharDataParser(ref dataString, ref leftIndex, ref rightIndex, ref theNewEntryToFill, ref EntryDataIndex, ref parsingMarkers, ref markerIndex);
                    }
                }
                //If we didn't find the target index with the given marker
                else
                {
                    //If we find a char that's apart of the next target and the one we're on is optional, skip it
                    if (markerIndex + 2 <= parsingMarkers.Length - 1 && parsingMarkers[markerIndex].optional && dataString[rightIndex] == parsingMarkers[markerIndex + 2].target)
                    {
                        //leftIndex = rightIndex++;
                        markerIndex += 2;
                        EntryDataIndex++;

                        SingleCharDataParser(ref dataString, ref leftIndex, ref rightIndex, ref theNewEntryToFill, ref EntryDataIndex, ref parsingMarkers, ref markerIndex);
                    }
                    else
                    {
                        //Increment right index till we find a target char
                        rightIndex++;
                    }
                }
            }

            //If we made it to the end of the string or we haven't gone through every marker, then we failed
            if (parsingMarkers.Length != markerIndex || rightIndex == dataString.Length)
            {
                return false;
            }

            return true;
        }
        
        /// <summary>
        /// Parses the data from a Textbox and displays it on a 'dataGridView'
        /// </summary>
        private void ParseAndDisplayTable()
        {
            //Parse the data
            entries = ParseText(textBox1.Text, comboBox1.SelectedIndex);

            //If any data was found and added
            if (entries.Count != 0)
            {
                //Enable gridView and display the data table
                dataGridView1.Visible = true;
                TableData = GenerateTable(entries);
                dataGridView1.DataSource = TableData;

                label1.Text = "Data Table:";

                //Enables buttons allows the user to either export to Excel or clear the table
                button1.Enabled = false;
                button2.Enabled = true;
                textBox2.Enabled = true;
                button3.Enabled = true;
                button4.Visible = true;
                button5.Enabled = false;
                comboBox1.Enabled = false;
            }
            else
            {
                MessageBox.Show("No Data was found!");
            }
        }

        /// <summary>
        /// Parses the given string, taking into account what type of parser to use.
        /// Returns a list of entrys assembled in order of their weights (Day value + time value)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        private List<Entry> ParseText(string data, int type)
        {
            List<Entry> entries = new List<Entry>();

            try
            {
                //ComboBox1 items in order:
                //Each element represents a parsing method
                // 0   - Standard
                // 0+1 - etc...

                //Type defines what kind of ParseMarkers are used to get the data
                switch(type)
                {
                    //Standard parser (My Parser)
                    case 0:

                        int leftIndex = 0, rightIndex = 0;

                        creditSum = 0;

                        while (rightIndex != data.Length)
                        {
                            Entry tableElement = new Entry();

                            int EntryIndex = 0, ParseMarkerIndex = 0;

                            //If we can parse out a data element, store in a table ordered by their weights (dayValue+timeValue)
                            if (SingleCharDataParser(ref data, ref leftIndex, ref rightIndex, ref tableElement, ref EntryIndex, ref StandardParseMarkers, ref ParseMarkerIndex))
                            {
                                //TODO: FIX THIS BUG!  (Parser will put what's supposed to be a day1 value into day2) 
                                //This condition corrects this instance
                                if (tableElement.entryData[4] == null && tableElement.entryData[5] != null)
                                {
                                    tableElement.entryData[4] = tableElement.entryData[5];
                                    tableElement.entryData[5] = null;
                                }

                                //Determine the weight given the day and time
                                tableElement.timeValue = DecodeTimeValue(tableElement.entryData[6]);
                                tableElement.day1weight = DecodeDayValue(tableElement.entryData[4]) + tableElement.timeValue;

                                //If a second day was found, find the weight of that too
                                if (tableElement.entryData[5] != null)
                                {
                                    tableElement.day2weight = DecodeDayValue(tableElement.entryData[5]) + tableElement.timeValue;
                                }

                                creditSum += Convert.ToInt32(tableElement.entryData[8].Substring(0,1));

                                //If this list was originally empty
                                if (entries.Count == 0)
                                {
                                    //If there was no second day present, just add it
                                    if (tableElement.day2weight == 0)
                                    {
                                        entries.Add(tableElement);
                                    }
                                    else
                                    {
                                        entries.Add(tableElement);

                                        //Create a new element for the second day
                                        Entry day2DuplicateElement = new Entry();
                                        tableElement.CopyData(ref day2DuplicateElement);

                                        day2DuplicateElement.day2Entry = true;

                                        entries.Add(day2DuplicateElement);
                                    }
                                }
                                else
                                {
                                    //If there was no second day present, add the element
                                    if (tableElement.day2weight == 0)
                                    {
                                        InsertIntoDataTable(ref tableElement, ref entries);
                                    }
                                    //else we're adding two elements
                                    else
                                    {
                                        InsertIntoDataTable(ref tableElement, ref entries);

                                        //Create a new element for the second day
                                        Entry day2DuplicateElement = new Entry();
                                        tableElement.CopyData(ref day2DuplicateElement);

                                        day2DuplicateElement.day2Entry = true;

                                        InsertIntoDataTable(ref day2DuplicateElement, ref entries);
                                    }
                                }
                            }
                        }
                       
                        break;

                    default:
                        throw new Exception("Parse Marker Type '"+type+"' doesn't exist.");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message+" : "+ex.StackTrace);
            }

            return entries;
        }

        /// <summary>
        /// A messy function that inserts into a table depending on if it's day1 or day2 weight is less than or equal to the corresponding day1 or day2 weight.
        /// Organizing through this for loop currently doesn't output the best result, so Linq is used after.
        /// </summary>
        /// <param name="elementToAdd"></param>
        /// <param name="EntryList"></param>
        private void InsertIntoDataTable(ref Entry elementToAdd, ref List<Entry> EntryList)
        {
            for(int i = 0, dayMarker = 0; i <= EntryList.Count; i++)
            {
                if((i == EntryList.Count) || dayMarker == 0 && ((elementToAdd.day2Entry && EntryList[i].day2Entry && EntryList[i].day2weight >= elementToAdd.day2weight) || (!elementToAdd.day2Entry && !EntryList[i].day2Entry && EntryList[i].day1weight >= elementToAdd.day1weight)))
                {
                    //Use this marker as a starting point to organize by time
                        dayMarker = i;
                    if ((i == EntryList.Count))
                    {
                        {

                        EntryList.Add(elementToAdd);
                        }

                        break;
                    }
                    else
                    {
                        EntryList.Insert(i, elementToAdd);



                        break;
                    }
                }
            }

            EntryList = EntryList.OrderByDescending(e => (!e.day2Entry) ? e.day1weight : e.day2weight).ToList();
            EntryList.Reverse();

        }

        /// <summary>
        /// Returns a value that corresponds to the day string that was passed in.
        /// Returns 0 and outputs a message if the day value is invalid
        /// </summary>
        /// <param name="day"></param>
        /// <returns></returns>
        private int DecodeDayValue(string day)
        {
            switch(day)
            {
                case "Monday":
                    return (int) DayValue.Monday;
                case "Tuesday":
                    return (int)DayValue.Tuesday;
                case "Wednesday":
                    return (int)DayValue.Wednesday;
                case "Thursday":
                    return (int)DayValue.Thursday;
                case "Friday":
                    return (int)DayValue.Friday;
                default:
                    throw new Exception("Day couldn't be decoded! \n Day String = "+day);
            }
        }

        /// <summary>
        /// Returns a value that corresponds to the time string that was passed in.
        /// Example input = "09:30AM - 10:45AM"
        /// </summary>
        /// <param name="time"></param>
        /// <returns></returns>
        private int DecodeTimeValue(string time)
        {
            //Get the starting hour to judge the weight
            string parsedTime = time.Substring(0, 7);
            bool isPM = (parsedTime[5] == 'P') ? true : false;

            int value = 0;

            //Convert the hour and minute value to an int
            value = Convert.ToInt32(parsedTime.Substring(0, 2) + parsedTime.Substring(3, 2));
            
            //Multiply it by two because we're going off a military time (sorta, if you ignore the exact minute value, which I don't think matters)
            value = (isPM) ? value+1200 : value;

            return value;
        }

        /// <summary>
        /// Takes an entrie of data and puts the to a table in ascending order (n to n+1)
        /// </summary>
        /// <param name="entries"></param>
        /// <returns></returns>
        private DataTable GenerateTable(List<Entry> entries)
        {
            DataTable table = new DataTable();

            //Add columns
            table.Columns.Add("Weekdays", typeof(string));
            table.Columns.Add("Class Name", typeof(string));
            table.Columns.Add("Time", typeof(string));
            table.Columns.Add("Location", typeof(string));
            table.Columns.Add("Start/End Date", typeof(string));
            table.Columns.Add("Professor", typeof(string));
            table.Columns.Add("Credits", typeof(string));

            //Add all all the data of each entry into their own row
            for (int i = 0; i < entries.Count; i++)
            {
                table.Rows.Add(((entries[i].day2Entry) ? entries[i].day2weight : entries[i].day1weight), entries[i].entryData[0], entries[i].entryData[6], entries[i].entryData[7], entries[i].entryData[3], entries[i].entryData[1], entries[i].entryData[8]);
            }

            return table;
        }

        /// <summary>
        /// Takes a day value from 0 to 9600+ and returns an index value representing that week day
        /// </summary>
        /// <param name="dayValue"></param>
        /// <returns></returns>
        private int DetermineDayValue(int dayValue)
        {
            if(dayValue >= (int) DayValue.Monday && dayValue < (int) DayValue.Tuesday)
            {
                return 0;
            }
            else if (dayValue >= (int)DayValue.Tuesday && dayValue < (int)DayValue.Wednesday)
            {
                return 1;
            }
            else if (dayValue >= (int)DayValue.Wednesday && dayValue < (int)DayValue.Thursday)
            {
                return 2;
            }
            else if (dayValue >= (int)DayValue.Thursday && dayValue < (int)DayValue.Friday)
            {
                return 3;
            }
            else if (dayValue >= (int)DayValue.Friday)
            {
                return 4;
            }
            else
            {
                throw new Exception("Dayvalue "+dayValue+" couldn't be determined!");
            }
        }

        /// <summary>
        /// Returns a width value that corresponds with the column number that's put in
        /// </summary>
        /// <param name="XTableCell"></param>
        /// <returns></returns>
        private double GetColumnWidth(int XTableCell)
        {
            switch(XTableCell)
            {
                //Weekday width
                case 1:
                    return 12;

                //Class name width
                case 2:
                    return 33.43;

                //Time width
                case 3:
                    return 19.42;

                //Location width
                case 4:
                    return 38.57;

                //Start/End Date width
                case 5:
                    return 26.28;

                //Professor width
                case 6:
                    return 43.85;

                //Credits width
                case 7:
                    return 7.15;

                default:
                    return 5;
            }
        }

        /// <summary>
        /// Takes the tableData that was created and creates it again in Excel with proper formatting
        /// </summary>
        /// <param name="tableData"></param>
        private void GenerateExcelFile(DataTable tableData, bool OutputRaw=false)
        {
            Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };

            Workbook excelWorkbook = Excel.Workbooks.Add(Type.Missing);
            Worksheet excelWorksheet = (Worksheet)excelWorkbook.ActiveSheet;

            //Find out the name using the data:
            string dateFound = entries[entries.Count - 1].entryData[3].Substring(0, entries[entries.Count - 1].entryData[3].Length / 2);
            int year = Convert.ToInt32(dateFound.Substring(6));
            string season = (Convert.ToInt32(dateFound.Substring(0, 2)) >= 9) ? "Fall" : "Spring";
            string fileNameAndPath = "";

            //Name the excel sheet
            excelWorksheet.Name = "College Schedule " + year.ToString() + " " + season.ToString();

            try
            {

                if (!OutputRaw)
                {
                    Excel.ActiveWindow.DisplayGridlines = false;

                    //Get data for the name of the file assuming that 'entries' has been populated
                    string[] tableHeader = new string[7] { "Weekdays", "Class Name", "Time", "Location", "Start/End Date", "Professor", "Credits" };

                    //Define all cells of size 10:
                    excelWorksheet.Cells.Font.Size = 10;
                    excelWorksheet.Cells.Font.Name = "Arial";

                    //Format the header
                    var HeaderRange = excelWorksheet.Range[excelWorksheet.Cells[1, 1], excelWorksheet.Cells[1, tableData.Columns.Count]];
                    HeaderRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                    HeaderRange.Borders.Weight = 4;

                    //Format the row height of the entire table and add a thin line border
                    var EntireTableRange = excelWorksheet.Range[excelWorksheet.Cells[2, 2], excelWorksheet.Cells[tableData.Rows.Count + 1, tableData.Columns.Count]];
                    EntireTableRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                    EntireTableRange.Borders.Weight = XlBorderWeight.xlThin;

                    EntireTableRange = excelWorksheet.Range[excelWorksheet.Cells[1, 1], excelWorksheet.Cells[tableData.Rows.Count + 1, tableData.Columns.Count]];
                    EntireTableRange.EntireRow.RowHeight = 41.25;

                    //Loop through each entry
                    //Left+right index should be used to create a range for that given day
                    for (int tableYvalue = 1, UpperYIndex = 0, BottomYIndex = 0, lastDayValueFound = -1, foundDayValue = 0; tableYvalue <= 12; tableYvalue++)
                    {
                        //Input each piece of data into each column cell
                        for (int tableXvalue = 1; tableXvalue <= 7; tableXvalue++)
                        {

                            //If we're in the header
                            if (tableYvalue == 1)
                            {
                                //Columns:
                                //Weekdays, Class Name, Time, Location, Start/End Date, Professor, Credits
                                excelWorksheet.Cells[tableYvalue, tableXvalue] = tableHeader[tableXvalue - 1];
                                excelWorksheet.Cells[tableYvalue, tableXvalue].Font.Bold = true;
                                excelWorksheet.Cells[tableYvalue, tableXvalue].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                excelWorksheet.Cells[tableYvalue, tableXvalue].VerticalAlignment = XlVAlign.xlVAlignCenter;
                            }
                            //If we're somewhere in the table
                            //Add actual data to the cell depending on tableXvalue
                            else
                            {
                                //Determine the day
                                foundDayValue = DetermineDayValue((entries[tableYvalue - 2].day2Entry) ? entries[tableYvalue - 2].day2weight : entries[tableYvalue - 2].day1weight);

                                //Set the font of the cells to 12
                                excelWorksheet.Cells[tableYvalue, tableXvalue].Font.Size = 12;

                                //Format each cell depending on the information that's being put in
                                switch (tableXvalue)
                                {
                                    //Class name
                                    case 2:
                                        excelWorksheet.Cells[tableYvalue, tableXvalue] = entries[tableYvalue - 2].entryData[0];

                                        //Assign the color of the cell to that day
                                        excelWorksheet.Cells[tableYvalue, tableXvalue].interior.Color = dayColors[foundDayValue];
                                        break;

                                    //Time
                                    case 3:
                                        //Format the time in a specific way (Ex: "9:30 AM - 10:45 AM" instead of "09:30AM - 10:45AM") 
                                        string time = entries[tableYvalue - 2].entryData[6];

                                        string startTime = time.Substring(0,7);
                                        int startHour = Convert.ToInt16(startTime.Substring(0, 2));

                                        string endTime = time.Substring(time.IndexOf('-') + 2, 7);
                                        int endHour = Convert.ToInt16(endTime.Substring(0, 2));

                                        string formattedTime = startHour.ToString() + startTime.Substring(2, 3)+' ';
                                        formattedTime += (startTime[5] == 'A') ? "AM" : "PM";
                                        formattedTime += " - " + endHour.ToString() + endTime.Substring(2, 3) + ' ';
                                        formattedTime += (endTime[5] == 'A') ? "AM" : "PM";

                                        excelWorksheet.Cells[tableYvalue, tableXvalue] = formattedTime;
                                        excelWorksheet.Cells[tableYvalue, tableXvalue].Font.Bold = true;

                                        //Assign the color of the cell to that day
                                        excelWorksheet.Cells[tableYvalue, tableXvalue].interior.Color = dayColors[foundDayValue];
                                        break;

                                    //Location
                                    case 4:
                                        string location = entries[tableYvalue - 2].entryData[7];

                                        string room = location.Substring(location.IndexOf(',') + 1);
                                        string building = location.Substring(0,location.IndexOf(','));

                                        string formattedLocation = building + room;

                                        excelWorksheet.Cells[tableYvalue, tableXvalue] = formattedLocation;

                                        //Assign the color of the cell to that day
                                        excelWorksheet.Cells[tableYvalue, tableXvalue].interior.Color = dayColors[foundDayValue];
                                        break;

                                    //Start/End date
                                    //Format the data in a specific way (Ex: "9/6/2018  -  1/23/2019" instead of "09/06/2018 - 01/23/2019")
                                    case 5:
                                        string startDate = entries[tableYvalue - 2].entryData[3].Substring(0, 10);
                                        string endDate = entries[tableYvalue - 2].entryData[3].Substring(11, 10);

                                        int startMonth = Convert.ToInt32(startDate.Substring(3, 2));
                                        int startDay = Convert.ToInt32(startDate.Substring(0, 2));

                                        int endMonth = Convert.ToInt32(endDate.Substring(3, 2));
                                        int endDay = Convert.ToInt32(endDate.Substring(0, 2));

                                        string formattedDate = "";

                                        //If the user prefers it to be Day/Month/Year instead of Month/Day/Year
                                        if (checkBox3.Checked)
                                        {
                                            formattedDate = startMonth.ToString() + '/' + startDay + '/' + year;
                                            formattedDate += "  -  " + endMonth + '/' + endDay + '/' + year;
                                        }
                                        else
                                        {
                                            formattedDate = startDay.ToString() + '/' + startMonth + '/' + year;
                                            formattedDate += "  -  " + endDay + '/' + endMonth + '/' + year;
                                        }

                                        excelWorksheet.Cells[tableYvalue, tableXvalue] = formattedDate;

                                        //Assign the color of the cell to that day
                                        excelWorksheet.Cells[tableYvalue, tableXvalue].interior.Color = dayColors[foundDayValue];
                                        break;

                                    //Professors name and email
                                    case 6:
                                        excelWorksheet.Cells[tableYvalue, tableXvalue].Font.Size = 12;

                                        //professorsName + " ("+email+")"
                                        string profcell = entries[tableYvalue - 2].entryData[1] + " (" + entries[tableYvalue - 2].entryData[2] + ")";
                                        int startEmailIndex = profcell.IndexOf('(');

                                        excelWorksheet.Cells[tableYvalue, tableXvalue] = profcell;

                                        //email 10 pnt font and bold
                                        Characters email = excelWorksheet.Cells[tableYvalue, tableXvalue].Characters(startEmailIndex + 2, entries[tableYvalue - 2].entryData[2].Length);
                                        email.Font.Bold = true;
                                        email.Font.Size = 10;

                                        //Assign the color of the cell to that day
                                        excelWorksheet.Cells[tableYvalue, tableXvalue].interior.Color = dayColors[foundDayValue];
                                        break;

                                    //Credits
                                    case 7:
                                        excelWorksheet.Cells[tableYvalue, tableXvalue] = entries[tableYvalue - 2].entryData[8];
                                        break;
                                }

                                //Center it 
                                excelWorksheet.Cells[tableYvalue, tableXvalue].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                excelWorksheet.Cells[tableYvalue, tableXvalue].VerticalAlignment = XlVAlign.xlVAlignCenter;

                                //If we just started remember the day we were on
                                if (lastDayValueFound == -1)
                                {
                                    UpperYIndex = tableYvalue;
                                    lastDayValueFound = foundDayValue;
                                }
                                else
                                {
                                    //If the day we're on isn't the same day or we hit the last entry
                                    if (lastDayValueFound != foundDayValue || tableYvalue == entries.Count + 1)
                                    {
                                        //Shift up one since it's not friday
                                        if (lastDayValueFound != 4)
                                        {
                                            BottomYIndex = tableYvalue - 1;
                                        }
                                        else
                                        {
                                            BottomYIndex = tableYvalue;
                                        }

                                        //Define range for lastDay using BottomYIndex and UpperYIndex and do formating for day range
                                        var DayColumn = excelWorksheet.Range[excelWorksheet.Cells[UpperYIndex, 1], excelWorksheet.Cells[BottomYIndex, 1]];
                                        DayColumn.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThick, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);

                                        //Format the bottom of this day cell 
                                        DayColumn = excelWorksheet.Range[excelWorksheet.Cells[BottomYIndex, 2], excelWorksheet.Cells[BottomYIndex, 7]];
                                        DayColumn.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;

                                        //Find the 'center cell'
                                        //If there's only one cell for this day
                                        if (BottomYIndex - UpperYIndex == 0)
                                        {
                                            excelWorksheet.Cells[UpperYIndex, 1].Font.Bold = true;
                                            excelWorksheet.Cells[UpperYIndex, 1].Font.Size = 10;
                                            excelWorksheet.Cells[UpperYIndex, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                            excelWorksheet.Cells[UpperYIndex, 1] = DayStrings[lastDayValueFound];
                                        }
                                        //If there's only two cells for this day
                                        else if (BottomYIndex - UpperYIndex == 1)
                                        {
                                            excelWorksheet.Cells[UpperYIndex, 1].Font.Bold = true;
                                            excelWorksheet.Cells[UpperYIndex, 1].Font.Size = 10;
                                            excelWorksheet.Cells[UpperYIndex, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                            excelWorksheet.Cells[UpperYIndex, 1].VerticalAlignment = XlVAlign.xlVAlignBottom;
                                            excelWorksheet.Cells[UpperYIndex, 1] = DayStrings[lastDayValueFound];
                                        }
                                        //If there's more than 2 cells
                                        else
                                        {
                                            //Find the distance from the center cell to find the center cell index
                                            int cellOffset = (BottomYIndex - UpperYIndex) / 2;
                                            cellOffset += UpperYIndex;

                                            excelWorksheet.Cells[cellOffset, 1].Font.Bold = true;
                                            excelWorksheet.Cells[cellOffset, 1].Font.Size = 10;
                                            excelWorksheet.Cells[cellOffset, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                            excelWorksheet.Cells[cellOffset, 1].VerticalAlignment = XlVAlign.xlVAlignBottom;
                                            excelWorksheet.Cells[cellOffset, 1] = DayStrings[lastDayValueFound];
                                        }

                                        //TODO: Remember this range for this day so we can go back and format it

                                        lastDayValueFound = foundDayValue;
                                        UpperYIndex = tableYvalue;
                                    }
                                    else
                                    {
                                        BottomYIndex++;
                                    }
                                }
                            }
                        }
                    }

                    //Change the width of each column to appropriately fit the data
                    //Also format the entire column's border
                    for (int i = 1; i <= 7; i++)
                    {
                        excelWorksheet.Columns[i].ColumnWidth = GetColumnWidth(i);
                        excelWorksheet.Range[excelWorksheet.Cells[2, i], excelWorksheet.Cells[entries.Count+1, i]].BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThick, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                    }

                    //If we want to also include the sum of the credits we're going for
                    if (checkBox2.Checked)
                    {
                        var creditCell = excelWorksheet.Cells[entries.Count + 3, 7];

                        creditCell.Font.Bold = true;
                        creditCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        creditCell.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        excelWorksheet.Cells[entries.Count + 3, 7] = creditSum.ToString();

                        var creditTotalLabel = excelWorksheet.Cells[entries.Count + 3, 6];

                        creditTotalLabel.Font.Bold = true;
                        creditTotalLabel.HorizontalAlignment = XlHAlign.xlHAlignRight;
                        creditTotalLabel.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        excelWorksheet.Cells[entries.Count + 3, 6] = "Credit Total:";
                    }
                }
                else
                {
                    //Loop through all entries
                    for (int y = 1; y <= entries.Count; y++)
                    {
                        //Loop through all the data in that entry
                        for (int x = 1; x <= entries[0].entryData.Length; x++)
                        {
                            int weight = (entries[y - 1].day2Entry) ? entries[y - 1].day2weight : entries[y - 1].day1weight;

                            //Store the data into each cell
                            excelWorksheet.Cells[y, x] = entries[y - 1].entryData[x - 1]+' '+weight;
                        }
                    }

                    //Make sure the data fits
                    excelWorksheet.UsedRange.Columns.AutoFit();
                    excelWorksheet.UsedRange.Rows.AutoFit();

                    excelWorksheet.Name += "Raw";
                }

                //Define our path and file name
                fileNameAndPath = textBox2.Text + '/' + excelWorksheet.Name;


            }
            catch (ApplicationException ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
            catch (IndexOutOfRangeException ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " : " + ex.StackTrace);
            }
            finally
            {
                //Save it and close it
                excelWorkbook.SaveAs(fileNameAndPath);
                excelWorkbook.Close();
                Excel.Quit();

                Marshal.ReleaseComObject(excelWorksheet);
                Marshal.ReleaseComObject(excelWorkbook);
                Marshal.ReleaseComObject(Excel);

                //Execute the file
                System.Diagnostics.Process.Start(fileNameAndPath + ".xlsx");
            }
        }
    }
}
