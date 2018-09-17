using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace ModResConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //StreamReader objInput = null;
        string contents;
        string[,] fileArray;
        int countLine;
        List<GridData> dataGrids;
        List<GridSP> dataSP;
        int lengthArray;
        int[] coordinateArray = new int[10];
        Settings win1;
        OpenFileDialog fileDialog;
        string[] arrayYaxis;
        string[] arrayXaxis;
        List<string> TempList;

        public MainWindow()
        {
            InitializeComponent();
            this.Title = "ModResConverter";
            //Application.Current.MainWindow.WindowState = WindowState.Maximized;
            //export.IsEnabled = false;
            comboX.IsEnabled = false;
            comboY.IsEnabled = false;
            comboSpace.IsEnabled = false;
        }

        private void btn1_Click(object sender, RoutedEventArgs e)
        {
            // do something
            comboX.Items.Clear();
            comboY.Items.Clear();
            comboSpace.Items.Clear();
            path1.Items.Clear();
            dataGrids = null;
            dataSP = null;
            OpenDialog();

        }

        private void OpenDialog()
        {
            
            lengthArray = 0;
            fileDialog = new OpenFileDialog();
            if (Properties.Settings.Default.SP_Setting)
            {
                fileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            }
            else
            {
                fileDialog.Filter = "All files (*.*)|*.*";
                fileDialog.Multiselect = true;
            }
            

            if (fileDialog.ShowDialog() == true)
            {

                //MessageBox.Show("excel file");
                if (Properties.Settings.Default.SP_Setting)
                {
                    openSPExcelFile();
                }
                else
                {
                    openFile();
                }           
            }
        }

        private void openSPExcelFile()
        {
            dataSP = new List<GridSP>();
            string fileName = fileDialog.FileName;
            int maxStation = 0;
            path1.Items.Add(fileDialog.FileName);

            try
            {
                using (var excelWorkbook = new XLWorkbook(fileName))
                {
                    var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();
                    int i = 0;
                    
                    foreach (var dataRow in nonEmptyDataRows)
                    {

                        if (i > 0)
                        {
                            ////Console.WriteLine(dataRow.Cell(2).GetValue<string>());
                            //string serial_ = dataRow.Cell(1).GetValue<string>();
                            //string date_ = dataRow.Cell(2).GetValue<string>();
                            //string line_ = dataRow.Cell(3).GetValue<string>();
                            string station_ = dataRow.Cell(4).GetValue<string>();
                            //string north_ = dataRow.Cell(5).GetValue<string>();
                            //string east_ = dataRow.Cell(6).GetValue<string>();
                            //string stime_ = dataRow.Cell(7).GetValue<string>();
                            //string mtime_ = dataRow.Cell(8).GetValue<string>();
                            //string reading1_ = dataRow.Cell(9).GetValue<string>();
                            //string reading2_ = dataRow.Cell(10).GetValue<string>();
                            //string reading3_ = dataRow.Cell(11).GetValue<string>();
                            //string reading4_ = dataRow.Cell(12).GetValue<string>();
                            //float average_num = (dataRow.Cell(9).GetValue<float>() + dataRow.Cell(10).GetValue<float>() + dataRow.Cell(11).GetValue<float>() + dataRow.Cell(12).GetValue<float>()) / 4;
                            //string average_ = string.Format("{0:N3}", average_num);
                            //string elevation_ = dataRow.Cell(14).GetValue<string>();
                            //string x_ = dataRow.Cell(15).GetValue<string>();
                            //string y_ = dataRow.Cell(16).GetValue<string>();
                            //string remarks_ = dataRow.Cell(17).GetValue<string>();

                            //dataSP.Add(new GridSP()
                            //{

                            //    serial = serial_,
                            //    date = date_,
                            //    line = line_,
                            //    station = station_,
                            //    north = north_,
                            //    east = east_,
                            //    second_time = stime_,
                            //    minute_time = mtime_,
                            //    reading_1 = reading1_,
                            //    reading_2 = reading2_,
                            //    reading_3 = reading3_,
                            //    reading_4 = reading4_,
                            //    average = average_,
                            //    elevation = elevation_,
                            //    x = x_,
                            //    y = y_,
                            //    remarks = remarks_


                            //});

                            //find highest value for comboSpace
                            int valueOut = 0;
                            if (int.TryParse(station_, out valueOut))
                            {

                                if (Convert.ToInt32(station_) > maxStation)
                                {
                                    //Console.WriteLine(station_);
                                    maxStation = Convert.ToInt32(station_);
                                }

                            }

                        }
                        i = 1;


                    }
                }
                for (int a = 1; a <= maxStation; a++)
                {
                    comboSpace.Items.Add(a);
                }
                //dataGrid1.ItemsSource = dataSP;
                comboSpace.IsEnabled = true;
            }
            catch
            {
                MessageBox.Show("Please close excel file");

            }
        }

        private void openFile()
        {
            //looping for length of each file
            int a = 1;
            arrayYaxis = new string[1000];
            arrayXaxis = new string[1000];
            TempList = new List<string>();
            fileArray = new string[1000, 4];
            int countLine = 0;
            foreach (string filelength in fileDialog.FileNames)
            {
                string fileType = Path.GetExtension(filelength);
                //Console.WriteLine(fileType);

                if (fileType == ".dat")
                {
                    countLine = datFileOperation(a, filelength, countLine);
                }
                else if (fileType == ".xls" || fileType == ".xlsx" || fileType == ".xlsm")
                {
                    countLine =  excelFileOperation(a, filelength, countLine);
                }
                else
                {
                    MessageBox.Show("Undefined file type. Please reupload only .dat and excel files");
                }
                a++;
                path1.Items.Add(filelength);
            }

            //Sort Value ComboBox
            TempList.Sort();
            foreach (string ListValue in TempList)
            {
                comboX.Items.Add(ListValue);
            }

            comboX.IsEnabled = true;
            comboY.IsEnabled = true;


        }

        private void comboX_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            String selectedValue = (String)comboX.SelectedValue;
            
            dataGrids = new List<GridData>();
            int z = 1;
            string lines = "Line 1";
            Console.WriteLine(coordinateArray[2] + "/" + lengthArray);
            for (int k = 1; k < lengthArray ; k++)
            {

                if(coordinateArray[z] == k)
                {
                    z++;
                    lines = "Line " + z;
                }
                //Console.WriteLine(k +"/"+ fileArray[k, 3]);
                if (fileArray[k, 1] == selectedValue)
                {
                    dataGrids.Add(new GridData()
                    {
                        line = lines,
                        //line = "Line " + k,
                        X = fileArray[k, 1],
                        Y = fileArray[k, 2],
                        Z = fileArray[k, 3]
                    });
                }
            }

            dataGrid1.ItemsSource = dataGrids;
            //export.IsEnabled = true;
        }

        private void comboY_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            String selectedValue = (String)comboY.SelectedValue;
            dataGrids = new List<GridData>();
            int z = 1;
            string lines = "Line 1";
            for (int k = 1; k < lengthArray; k++)
            {

                if (coordinateArray[z] == k)
                {
                    z++;
                    lines = "Line " + z;
                }
                //Console.WriteLine(k +"/"+ fileArray[k, 3]);
                if (fileArray[k, 2] == selectedValue)
                {

                    dataGrids.Add(new GridData()
                    {
                        line = lines,
                        X = fileArray[k, 1],
                        Y = fileArray[k, 2],
                        Z = fileArray[k, 3]
                    });
                }
            }

            dataGrid1.ItemsSource = dataGrids;
            //export.IsEnabled = true;

        }

        private void comboSpace_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            dataSP.Clear();
            int selectedValue = (int)comboSpace.SelectedValue;
            string select = selectedValue.ToString();
            string fileName = fileDialog.FileName;
            try
            {
                using (var excelWorkbook = new XLWorkbook(fileName))
                {
                    var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();
                    int  n = 1;

                    foreach (var dataRow in nonEmptyDataRows)
                    {
                        string station = dataRow.Cell(4).GetValue<string>();
                        int valueOut = 0;
                        if (int.TryParse(station, out valueOut))
                        {
                            //Console.WriteLine(Convert.ToInt32(station));
                            if (Convert.ToInt32(station) != 0)
                            {
                                

                                if (Convert.ToInt32(station) == (selectedValue * n) )
                                {
                                
                                    //Console.WriteLine(dataRow.Cell(2).GetValue<string>());
                                    string serial_ = dataRow.Cell(1).GetValue<string>();
                                    string date_ = dataRow.Cell(2).GetValue<string>();
                                    string line_ = dataRow.Cell(3).GetValue<string>();
                                    string station_ = dataRow.Cell(4).GetValue<string>();
                                    string north_ = dataRow.Cell(5).GetValue<string>();
                                    string east_ = dataRow.Cell(6).GetValue<string>();
                                    string stime_ = dataRow.Cell(7).GetValue<string>();
                                    string mtime_ = dataRow.Cell(8).GetValue<string>();
                                    string reading1_ = dataRow.Cell(9).GetValue<string>();
                                    string reading2_ = dataRow.Cell(10).GetValue<string>();
                                    string reading3_ = dataRow.Cell(11).GetValue<string>();
                                    string reading4_ = dataRow.Cell(12).GetValue<string>();
                                    float average_num = (dataRow.Cell(9).GetValue<float>() + dataRow.Cell(10).GetValue<float>() + dataRow.Cell(11).GetValue<float>() + dataRow.Cell(12).GetValue<float>()) / 4;
                                    string average_ = string.Format("{0:N3}", average_num);
                                    string elevation_ = dataRow.Cell(14).GetValue<string>();
                                    string x_ = dataRow.Cell(15).GetValue<string>();
                                    string y_ = dataRow.Cell(16).GetValue<string>();
                                    string remarks_ = dataRow.Cell(17).GetValue<string>();

                                    dataSP.Add(new GridSP()
                                    {

                                        serial = serial_,
                                        date = date_,
                                        line = line_,
                                        station = station_,
                                        north = north_,
                                        east = east_,
                                        second_time = stime_,
                                        minute_time = mtime_,
                                        reading_1 = reading1_,
                                        reading_2 = reading2_,
                                        reading_3 = reading3_,
                                        reading_4 = reading4_,
                                        average = average_,
                                        elevation = elevation_,
                                        x = x_,
                                        y = y_,
                                        remarks = remarks_


                                    });
                                    n++;
                                }
                                
                            }
                            else
                            {
                                string serial_ = dataRow.Cell(1).GetValue<string>();
                                string date_ = dataRow.Cell(2).GetValue<string>();
                                string line_ = dataRow.Cell(3).GetValue<string>();
                                string station_ = dataRow.Cell(4).GetValue<string>();
                                string north_ = dataRow.Cell(5).GetValue<string>();
                                string east_ = dataRow.Cell(6).GetValue<string>();
                                string stime_ = dataRow.Cell(7).GetValue<string>();
                                string mtime_ = dataRow.Cell(8).GetValue<string>();
                                string reading1_ = dataRow.Cell(9).GetValue<string>();
                                string reading2_ = dataRow.Cell(10).GetValue<string>();
                                string reading3_ = dataRow.Cell(11).GetValue<string>();
                                string reading4_ = dataRow.Cell(12).GetValue<string>();
                                float average_num = (dataRow.Cell(9).GetValue<float>() + dataRow.Cell(10).GetValue<float>() + dataRow.Cell(11).GetValue<float>() + dataRow.Cell(12).GetValue<float>()) / 4;
                                string average_ = string.Format("{0:N3}", average_num);
                                string elevation_ = dataRow.Cell(14).GetValue<string>();
                                string x_ = dataRow.Cell(15).GetValue<string>();
                                string y_ = dataRow.Cell(16).GetValue<string>();
                                string remarks_ = dataRow.Cell(17).GetValue<string>();

                                dataSP.Add(new GridSP()
                                {

                                    serial = serial_,
                                    date = date_,
                                    line = line_,
                                    station = station_,
                                    north = north_,
                                    east = east_,
                                    second_time = stime_,
                                    minute_time = mtime_,
                                    reading_1 = reading1_,
                                    reading_2 = reading2_,
                                    reading_3 = reading3_,
                                    reading_4 = reading4_,
                                    average = average_,
                                    elevation = elevation_,
                                    x = x_,
                                    y = y_,
                                    remarks = remarks_


                                });
                                n++;

                                n = 1; //reset for line
                            }
                            //i++;
                        }
                    }
                }
                
                dataGrid1.ItemsSource = dataSP;
                dataGrid1.Items.Refresh();

            }
            catch
            {
                MessageBox.Show("Error");
            }
        }

        private void export_Click(object sender, RoutedEventArgs e)
        {
            if(dataGrids != null)
            {
                var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Data");
                        int row = 2;
                        worksheet.Cell("A1").Value = "Line";
                        worksheet.Cell("B1").Value = "X";
                        worksheet.Cell("C1").Value = "Y";
                        worksheet.Cell("D1").Value = "Z";
                        foreach (GridData GridData in dataGrids)
                        {
                            worksheet.Cell("A" + row.ToString()).Value = GridData.line.ToString();
                            worksheet.Cell("B" + row.ToString()).Value = GridData.X.ToString();
                            worksheet.Cell("C" + row.ToString()).Value = GridData.Y.ToString();
                            worksheet.Cell("D" + row.ToString()).Value = GridData.Z.ToString();
                            row++;

                        }
                        workbook.SaveAs(saveFileDialog.FileName);

                        MessageBox.Show("Successfully Export");
                    }
                }
            }
            else if(dataSP != null)
            {
                var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Data");
                        int row = 2;
                        int count = 1;
                        char countChar = 'A';

                        if (Properties.Settings.Default.Serial_number)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "S/N";
                            countChar++;
                        }
                        if (Properties.Settings.Default.Date)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "Date";
                            countChar++;
                        }
                        if (Properties.Settings.Default.Lines)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "Lines";
                            countChar++;
                        }
                        if (Properties.Settings.Default.Station)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "Station(m)";
                            countChar++;
                        }
                        if (Properties.Settings.Default.Northing)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "Northing";
                            countChar++;
                        }
                        if (Properties.Settings.Default.Easting)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "Easting";
                            countChar++;
                        }
                        if (Properties.Settings.Default.Second)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "Time(s)";
                            countChar++;
                        }
                        if (Properties.Settings.Default.Minute)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "Time(m)";
                            countChar++;
                        }
                        if (Properties.Settings.Default.Reading_1)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "Reading1";
                            countChar++;
                        }
                        if (Properties.Settings.Default.Reading_2)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "Reading2";
                            countChar++;
                        }
                        if (Properties.Settings.Default.Reading_3)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "Reading3";
                            countChar++;
                        }
                        if (Properties.Settings.Default.Reading_4)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "Reading4";
                            countChar++;
                        }
                        if (Properties.Settings.Default.Average)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "Average";
                            countChar++;
                        }
                        if (Properties.Settings.Default.Elevation)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "Elevation";
                            countChar++;

                        }
                        if (Properties.Settings.Default.X)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "x";
                            countChar++;

                        }
                        if (Properties.Settings.Default.Y)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "y";
                            countChar++;

                        }
                        if (Properties.Settings.Default.Remarks)
                        {
                            String cell = String.Concat(countChar, count);
                            worksheet.Cell(cell).Value = "Remarks";
                            countChar++;

                        }

                        foreach (GridSP GridData in dataSP)
                        {
                            countChar = 'A';
                            if (Properties.Settings.Default.Serial_number)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.serial.ToString();
                                countChar++;
                            }
                            if (Properties.Settings.Default.Date)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.date.ToString();
                                countChar++;
                            }
                            if (Properties.Settings.Default.Lines)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.line.ToString();
                                countChar++;
                            }
                            if (Properties.Settings.Default.Station)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.station.ToString();
                                countChar++;
                            }
                            if (Properties.Settings.Default.Northing)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.north.ToString();
                                countChar++;
                            }
                            if (Properties.Settings.Default.Easting)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.east.ToString();
                                countChar++;
                            }
                            if (Properties.Settings.Default.Second)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.second_time.ToString();
                                countChar++;
                            }
                            if (Properties.Settings.Default.Minute)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.minute_time.ToString();
                                countChar++;
                            }
                            if (Properties.Settings.Default.Reading_1)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.reading_1.ToString();
                                countChar++;
                            }
                            if (Properties.Settings.Default.Reading_2)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.reading_2.ToString();
                                countChar++;
                            }
                            if (Properties.Settings.Default.Reading_3)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.reading_3.ToString();
                                countChar++;
                            }
                            if (Properties.Settings.Default.Reading_4)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.reading_4.ToString();
                                countChar++;
                            }
                            if (Properties.Settings.Default.Average)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.average.ToString();
                                countChar++;
                            }
                            if (Properties.Settings.Default.Elevation)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.elevation.ToString();
                                countChar++;

                            }
                            if (Properties.Settings.Default.X)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.x.ToString();
                                countChar++;

                            }
                            if (Properties.Settings.Default.Y)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.y.ToString();
                                countChar++;

                            }
                            if (Properties.Settings.Default.Remarks)
                            {
                                String cell = String.Concat(countChar, row);
                                worksheet.Cell(cell).Value = GridData.remarks.ToString();
                                countChar++;

                            }
                            row++;

                        }
                        workbook.SaveAs(saveFileDialog.FileName);

                        MessageBox.Show("Successfully Export");
                    }
                }
            }
            else
            {
                MessageBox.Show("No Data. Please upload data first!");
            }
            
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            
            if (win1 != null)
            {
                if (win1.WindowState == WindowState.Minimized)
                {
                    win1.WindowState = WindowState.Normal;
                }
                else
                {
                    win1 = new Settings();
                    win1.Show();
                }
            }
            else
            {
                win1 = new Settings();
                win1.Show();
            }
            
        }

        private void DataGrid_OnAutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (!Properties.Settings.Default.Serial_number)
            {
                if (e.PropertyName == "serial")
                {
                    e.Column = null;
                }
            }
            if (!Properties.Settings.Default.Date)
            {
                if (e.PropertyName == "date")
                {
                    e.Column = null;
                }
            }
            if (!Properties.Settings.Default.Lines)
            {
                if (e.PropertyName == "line")
                {
                    e.Column = null;
                }
            }
            if (!Properties.Settings.Default.Station)
            {
                if (e.PropertyName == "station")
                {
                    e.Column = null;
                }
            }
            if (!Properties.Settings.Default.Northing)
            {
                if (e.PropertyName == "north")
                {
                    e.Column = null;
                }
            }
            if (!Properties.Settings.Default.Easting)
            {
                if (e.PropertyName == "east")
                {
                    e.Column = null;
                }
            }
            if (!Properties.Settings.Default.Second)
            {
                if (e.PropertyName == "second_time")
                {
                    e.Column = null;
                }
            }
            if (!Properties.Settings.Default.Minute)
            {
                if (e.PropertyName == "minute_time")
                {
                    e.Column = null;
                }
            }
            if (!Properties.Settings.Default.Reading_1)
            {
                if (e.PropertyName == "reading_1")
                {
                    e.Column = null;
                }
            }
            if (!Properties.Settings.Default.Reading_2)
            {
                if (e.PropertyName == "reading_2")
                {
                    e.Column = null;
                }
            }
            if (!Properties.Settings.Default.Reading_3)
            {
                if (e.PropertyName == "reading_3")
                {
                    e.Column = null;
                }
            }
            if (!Properties.Settings.Default.Reading_4)
            {
                if (e.PropertyName == "reading_4")
                {
                    e.Column = null;
                }
            }
            if (!Properties.Settings.Default.Average)
            {
                if (e.PropertyName == "average")
                {
                    e.Column = null;
                }
            }
            if (!Properties.Settings.Default.Elevation)
            {
                if (e.PropertyName == "elevation")
                {
                    e.Column = null;
                }

            }
            if (!Properties.Settings.Default.X)
            {
                if (e.PropertyName == "x")
                {
                    e.Column = null;
                }

            }
            if (!Properties.Settings.Default.Y)
            {
                if (e.PropertyName == "y")
                {
                    e.Column = null;
                }

            }
            if (!Properties.Settings.Default.Remarks)
            {
                if (e.PropertyName == "remarks")
                {
                    e.Column = null;
                }

            }
            
        }

        private int datFileOperation(int a, string filelength, int countline)
        {
            //Console.WriteLine(a);
            try
            {
                StreamReader objInputs = new StreamReader(filelength, System.Text.Encoding.Default);
                contents = objInputs.ReadToEnd().Trim();
                string[] splits = System.Text.RegularExpressions.Regex.Split(contents, "\\r+", RegexOptions.None);
                lengthArray = lengthArray + splits.Length;
                coordinateArray[a] = lengthArray;
                a++;
                //Console.WriteLine(lengthArray);

                int skip = 0, i = countline;
                foreach (string s in splits)
                {

                    //Console.WriteLine(s);
                    if (skip != 0)
                    {
                        int j = 0;
                        string[] space = System.Text.RegularExpressions.Regex.Split(s, "\\s+", RegexOptions.None);
                        foreach (string p in space)
                        {
                            //Console.WriteLine(i + "/" + p);
                            string p_replace = p.Replace("\"", "");
                            if (j == 1)
                            {
                                if (arrayXaxis.Contains(p) == false && p_replace != "X-location,Z-location,Resistivity")
                                {
                                    //comboX.Items.Add(p);
                                    TempList.Add(p);

                                }
                                arrayXaxis[i] = p;
                                fileArray[i, j] = p;
                                j++;
                            }
                            else if (j == 2)
                            {
                                if (arrayYaxis.Contains(p) == false)
                                {
                                    comboY.Items.Add(p);
                                }
                                arrayYaxis[i] = p;
                                fileArray[i, j] = p;
                                j++;
                            }
                            else if (j == 3)
                            {
                                //Console.WriteLine(p);
                                fileArray[i, j] = p;
                                j = 0;
                            }
                            else
                            {
                                j++;
                            }
                        }
                    }
                    skip = 1;
                    i++;

                }
                countline = i;

            }
            catch (IOException)
            {
                MessageBox.Show("Please currently use by another proceess.");
            }

            return countline;
        }

        private int excelFileOperation( int a, string filelength, int countline)
        {

            using (var excelWorkbook = new XLWorkbook(filelength))
            {
                var ws = excelWorkbook.Worksheet(1);
                var nonEmptyDataRows = ws.RowsUsed().Count();
                int row = nonEmptyDataRows + countline;
                int m;
                Console.WriteLine( a +" / " + row);
                coordinateArray[a] = row;
                lengthArray = row;

                for (int n = 2; n < nonEmptyDataRows; n++)
                {
                    m = countline + n;
                    //Console.WriteLine(m);
                    //getdata
                    String x = ws.Cell(n, 1).GetString();
                    String y = ws.Cell(n, 2).GetString();
                    String z = ws.Cell(n, 3).GetString();
                    //Console.WriteLine(z);


                    //add to x,y,z axis
                    if (arrayXaxis.Contains(x) == false)
                    {
                        comboX.Items.Add(x);
                        //Console.WriteLine(x);
                    }
                    arrayXaxis[m] = x;
                    fileArray[m, 1] = x;

                    if (arrayYaxis.Contains(y) == false)
                    {
                        comboY.Items.Add(y);
                    }
                    arrayYaxis[m] = y;
                    fileArray[m, 2] = y;
                    fileArray[m, 3] = z;
                    //Boolean cellDouble = (Boolean)cellBoolean.Value;
                    //Console.WriteLine(x);
                }

                countline = row;
            }
            return countline;
        }

    }
}
