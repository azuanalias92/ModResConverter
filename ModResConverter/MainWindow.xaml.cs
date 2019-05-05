﻿using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;
using Microsoft.Win32;
using CoordinateSharp;
using System.Windows.Media;

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
        int countLine = 0;
        List<GridData> dataGrids;
        List<GridSP> dataSP;
        int lengthArray;
        int[] coordinateArray = new int[10];
        Settings win1;
        OpenFileDialog fileDialog;
        string[] arrayYaxis;
        string[] arrayXaxis;
        List<string> TempList;
        int maxStation = 0;

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
                    comboX.IsEnabled = false;
                    comboY.IsEnabled = false;
                    openSPExcelFile();
                }
                else
                {
                    comboSpace.IsEnabled = false;
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

            if (!comboX.Items.Contains("ALL"))
            {
                comboX.Items.Add("ALL");
            }
            if (!comboY.Items.Contains("ALL"))
            {
                comboY.Items.Add("ALL");
            }

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
                            
                            string station_ = dataRow.Cell(4).GetValue<string>();
                            string x_ = dataRow.Cell(15).GetValue<string>();
                            string y_ = dataRow.Cell(16).GetValue<string>();

                            //insert value into x and y combo
                            if (!comboX.Items.Contains(x_))
                            {
                                comboX.Items.Add(x_);
                            }
                            if (!comboY.Items.Contains(y_))
                            {
                                comboY.Items.Add(y_);
                            }
                            

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
                comboSpace.IsEnabled = true;
                comboX.IsEnabled = true;
                comboY.IsEnabled = true;
            }
            catch
            {
                MessageBox.Show("Please close excel file");
                comboX.Items.Clear();
                comboY.Items.Clear();
                comboSpace.Items.Clear();
                path1.Items.Clear();
                dataGrids = null;
                dataSP = null;
                dataGrid1.ItemsSource = null;

            }
        }

        private void openFile()
        {
            //looping for length of each file
            int a = 1;
            TempList = new List<string>();
            
            int totalLine = 0;
            int countLine = 0;
            bool fileLock = false;
            foreach (string filelength in fileDialog.FileNames)
            {
                FileInfo fi1 = new FileInfo(filelength);
                path1.Items.Add(filelength);

                if (IsFileLocked(fi1))
                {
                    MessageBox.Show("Close this file " + filelength);
                    comboX.Items.Clear();
                    comboY.Items.Clear();
                    comboSpace.Items.Clear();
                    path1.Items.Clear();
                    dataGrids = null;
                    dataSP = null;
                    dataGrid1.ItemsSource = null;
                    fileLock = true;
                }
                else
                {
                    string fileType = Path.GetExtension(filelength);

                    if (!fileLock)
                    {
                        if (fileType == ".dat")
                        {
                            totalLine += File.ReadLines(filelength).Count();
                            //Console.WriteLine("total Line:" + totalLine);
                        }
                        else if (fileType == ".xls" || fileType == ".xlsx" || fileType == ".xlsm")
                        {
                            using (var excelWorkbook = new XLWorkbook(filelength))
                            {
                                var ws = excelWorkbook.Worksheet(1);
                                var nonEmptyDataRows = ws.RowsUsed().Count();

                                totalLine += nonEmptyDataRows;
                                //Console.WriteLine("total Line:" + totalLine);
                            }

                        }
                        else
                        {
                            MessageBox.Show("Undefined file type. Please reupload only .dat and excel files");
                        }
                    }
                    
                }
            }

            totalLine = totalLine + 1000;
            //Console.WriteLine("total Line:" + totalLine);
            fileArray = new string[totalLine, 3];
            arrayYaxis = new string[totalLine];
            arrayXaxis = new string[totalLine];

            foreach (string filelength in fileDialog.FileNames)
            {
                string fileType = Path.GetExtension(filelength);

                if (!fileLock)
                {
                    if (fileType == ".dat")
                    {
                        countLine = datFileOperation(a, filelength, countLine, totalLine);
                    }
                    else if (fileType == ".xls" || fileType == ".xlsx" || fileType == ".xlsm")
                    {
                        // Console.WriteLine(countLine);
                        countLine = excelFileOperation(a, filelength, countLine);
                    }
                    else
                    {
                        MessageBox.Show("Undefined file type. Please reupload only .dat and excel files");
                    }
                }
                a++;
                
            }

            //Sort Value ComboBox
            TempList.Sort();
            foreach (string ListValue in TempList)
            {
                comboX.Items.Add(ListValue);
            }


            //add value to space
            Console.WriteLine("maxStation" + maxStation);
            for (int s = 1; s <= maxStation; s++)
            {
                comboSpace.Items.Add(s);
            }

            comboX.IsEnabled = true;
            comboY.IsEnabled = true;
            comboSpace.IsEnabled = true;
        }

        private void comboX_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            String selectedValueX = (String)comboX.SelectedValue;
            String selectedValueY = (String)comboY.SelectedValue;
            int selectedValueSpace;
            try
            {
                selectedValueSpace = (int)comboSpace.SelectedValue;
            }
            catch
            {
                selectedValueSpace = 1;
            }
            if (Properties.Settings.Default.SP_Setting)
            {
                comboXSP(selectedValueX, selectedValueY, selectedValueSpace);
            }
            else
            {
                comboXRes(selectedValueX, selectedValueY);
            }
        }

        private void comboXSP(string selectedValueX, string selectedValueY, int selectedValueSpace)
        {
            dataSP.Clear();
            string fileName = fileDialog.FileName;
            try
            {
                using (var excelWorkbook = new XLWorkbook(fileName))
                {
                    var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();
                    int counter = 0;
                    int n = 1;

                    foreach (var dataRow in nonEmptyDataRows)
                    {
                        string x = dataRow.Cell(15).GetValue<string>();
                        string y = dataRow.Cell(16).GetValue<string>();

                        if (counter > 0)
                        {
                           if (x == selectedValueX || selectedValueX == "ALL")
                            {
                                if (y == selectedValueY || selectedValueY == "ALL" || selectedValueY == null)
                                {
                                    string station = dataRow.Cell(4).GetValue<string>();
                                    int valueOut = 0;
                                    if (int.TryParse(station, out valueOut))
                                    {
                                        //Console.WriteLine(Convert.ToInt32(station));
                                        if (Convert.ToInt32(station) != 0)
                                        {
                                            if (Convert.ToInt32(station) == (selectedValueSpace * n))
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
                                                //convert to utm
                                                Coordinate c = new Coordinate(dataRow.Cell(5).GetValue<double>(), dataRow.Cell(6).GetValue<double>(), new DateTime(2018, 6, 5, 10, 10, 0));
                                                string utm = c.UTM.ToString();

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
                                                    remarks = remarks_,
                                                    UTM = utm


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
                                            //convert to utm
                                            Coordinate c = new Coordinate(dataRow.Cell(5).GetValue<double>(), dataRow.Cell(6).GetValue<double>(), new DateTime(2018, 6, 5, 10, 10, 0));
                                            string utm = c.UTM.ToString();

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
                                                remarks = remarks_,
                                                UTM = utm


                                            });
                                            n++;

                                            n = 1; //reset for line
                                        }

                                    }
                                }

                            }
                        }
                        counter++;
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

        private void comboXRes(string selectedValueX, string selectedValueY)
        {
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
                if (fileArray[k, 0] == selectedValueX || selectedValueX == "ALL")
                {
                    if (fileArray[k, 1] == selectedValueY || selectedValueY == "ALL" || selectedValueY == null)
                    {
                        Double x_double;
                        Double.TryParse(fileArray[k, 0], out x_double);
                        Double y_double;
                        Double.TryParse(fileArray[k, 1], out y_double);
                        dataGrids.Add(new GridData()
                        {
                            line = lines,
                            X = fileArray[k, 0],
                            Y = fileArray[k, 1],
                            Z = fileArray[k, 2]

                        });
                    }
                }
            }
            dataGrid1.ItemsSource = dataGrids;
        }

        private void comboY_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            String selectedValueY = (String)comboY.SelectedValue;
            String selectedValueX = (String)comboX.SelectedValue;
            int selectedValueSpace;
            try
            {
                selectedValueSpace = (int)comboSpace.SelectedValue;
            }
            catch
            {
                selectedValueSpace = 1;
            }

            if (Properties.Settings.Default.SP_Setting)
            {
                comboYSP(selectedValueX, selectedValueY, selectedValueSpace);
            }
            else
            {
                comboYRes(selectedValueX, selectedValueY);
            }
        }

        private void comboYSP(string selectedValueX, string selectedValueY, int selectedValueSpace)
        {
            dataSP.Clear();
            string fileName = fileDialog.FileName;
            try
            {
                using (var excelWorkbook = new XLWorkbook(fileName))
                {
                    var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();
                    int counter = 0;
                    int n = 1;
                    foreach (var dataRow in nonEmptyDataRows)
                    {
                        string x = dataRow.Cell(15).GetValue<string>();
                        string y = dataRow.Cell(16).GetValue<string>();

                        if (counter > 0)
                        {
                            if (y == selectedValueY  || selectedValueY == "ALL")
                            {
                                if (x == selectedValueX || selectedValueX == "ALL" || selectedValueX == null)
                                {

                                    string station = dataRow.Cell(4).GetValue<string>();
                                    int valueOut = 0;
                                    if (int.TryParse(station, out valueOut))
                                    {
                                        //Console.WriteLine(Convert.ToInt32(station));
                                        if (Convert.ToInt32(station) != 0)
                                        {
                                            if (Convert.ToInt32(station) == (selectedValueSpace * n))
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
                                                //convert to utm
                                                Coordinate c = new Coordinate(dataRow.Cell(5).GetValue<double>(), dataRow.Cell(6).GetValue<double>(), new DateTime(2018, 6, 5, 10, 10, 0));
                                                string utm = c.UTM.ToString();

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
                                                    remarks = remarks_,
                                                    UTM = utm


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
                                            //convert to utm
                                            Coordinate c = new Coordinate(dataRow.Cell(5).GetValue<double>(), dataRow.Cell(6).GetValue<double>(), new DateTime(2018, 6, 5, 10, 10, 0));
                                            string utm = c.UTM.ToString();

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
                                                remarks = remarks_,
                                                UTM = utm


                                            });
                                            n++;

                                            n = 1; //reset for line
                                        }

                                    }
                                }

                            }
                        }
                        counter++;
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

        private void comboYRes(string selectedValueX, string selectedValueY)
        {
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
                if (fileArray[k, 1] == selectedValueY || selectedValueY == "ALL")
                {
                    if (fileArray[k, 0] == selectedValueX || selectedValueX == "ALL" || selectedValueX == null)
                    {
                        dataGrids.Add(new GridData()
                        {
                            line = lines,
                            X = fileArray[k, 0],
                            Y = fileArray[k, 1],
                            Z = fileArray[k, 2]
                        });
                    }

                }
            }

            dataGrid1.ItemsSource = dataGrids;
        }

        private void comboSpace_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            String selectedValueY = (String)comboY.SelectedValue;
            String selectedValueX = (String)comboX.SelectedValue;
            int selectedValueSpace;
            try
            {
                selectedValueSpace = (int)comboSpace.SelectedValue;
            }
            catch
            {
                selectedValueSpace = 1;
            }

            if (Properties.Settings.Default.SP_Setting)
            {
                comboSpaceSP(selectedValueX, selectedValueY, selectedValueSpace);
            }
            else
            {
                comboSpaceRes(selectedValueSpace);
            }
        }

        private void comboSpaceSP(String selectedValueX, String selectedValueY, int selectedValueSpace)
        {
            dataSP.Clear();
            string select = selectedValueSpace.ToString();
            string fileName = fileDialog.FileName;
            try
            {
                using (var excelWorkbook = new XLWorkbook(fileName))
                {
                    var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();
                    int n = 1;

                    foreach (var dataRow in nonEmptyDataRows)
                    {
                        string station = dataRow.Cell(4).GetValue<string>();
                        string x = dataRow.Cell(15).GetValue<string>();
                        string y = dataRow.Cell(16).GetValue<string>();
                        int valueOut = 0;

                        if (y == selectedValueY || selectedValueY == "ALL" || selectedValueY == null)
                        {
                            if (x == selectedValueX || selectedValueX == "ALL" || selectedValueX == null)
                            {
                                if (int.TryParse(station, out valueOut))
                                {
                                    //Console.WriteLine(Convert.ToInt32(station));
                                    if (Convert.ToInt32(station) != 0)
                                    {
                                        if (Convert.ToInt32(station) == (selectedValueSpace * n))
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
                                            //convert to utm
                                            Coordinate c = new Coordinate(dataRow.Cell(5).GetValue<double>(), dataRow.Cell(6).GetValue<double>(), new DateTime(2018, 6, 5, 10, 10, 0));
                                            string utm = c.UTM.ToString();

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
                                                remarks = remarks_,
                                                UTM = utm


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
                                        //convert to utm
                                        Coordinate c = new Coordinate(dataRow.Cell(5).GetValue<double>(), dataRow.Cell(6).GetValue<double>(), new DateTime(2018, 6, 5, 10, 10, 0));
                                        string utm = c.UTM.ToString();

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
                                            remarks = remarks_,
                                            UTM = utm


                                        });
                                        n++;

                                        n = 1; //reset for line
                                    }
                                    //i++;
                                }
                            }
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

        private void comboSpaceRes(int selectedValue)
        {
            int selectedValueSpace   = (int)comboSpace.SelectedValue;
            String selectedValueY       = (String)comboY.SelectedValue;
            String selectedValueX       = (String)comboX.SelectedValue;

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
                for (int a = maxStation; a > 0; a-= selectedValueSpace)
                {
                    if(Convert.ToInt32(Convert.ToDouble(fileArray[k, 2])) == a)
                    {
                        if (fileArray[k, 1] == selectedValueY || selectedValueY == "ALL" || selectedValueY == null)
                        {
                            if (fileArray[k, 0] == selectedValueX || selectedValueX == "ALL" || selectedValueX == null)
                            {
                                dataGrids.Add(new GridData()
                                {
                                    line = lines,
                                    X = fileArray[k, 0],
                                    Y = fileArray[k, 1],
                                    Z = fileArray[k, 2]
                                });
                            }

                        }
                    }
                    
                }

                
            }

            dataGrid1.ItemsSource = dataGrids;
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

        private int datFileOperation(int a, string filelength, int countline, int totalLine)
        {
            //Console.WriteLine("a:" +  a);
            //Console.WriteLine("filelength:" + filelength);
            //Console.WriteLine("countline:" + countline);
            //fileArray = new string[totalLine, 3];

            try
            {
                StreamReader objInputs = new StreamReader(filelength, System.Text.Encoding.Default);
                contents = objInputs.ReadToEnd().Trim();
                string[] splits = System.Text.RegularExpressions.Regex.Split(contents, "\\r+", RegexOptions.None);
                lengthArray = lengthArray + splits.Length;
                coordinateArray[a] = lengthArray;
                //a++;
                //Console.WriteLine(splits);

                //default value
                if (!comboX.Items.Contains("ALL"))
                {
                    comboX.Items.Add("ALL");
                }
                if (!comboY.Items.Contains("ALL"))
                {
                    comboY.Items.Add("ALL");
                }

                int skip = 0, i = countline;

                foreach (string s in splits)
                {

                    //Console.WriteLine(s);
                    if (skip != 0)
                    {
                        int j = -1;
                        string[] space = System.Text.RegularExpressions.Regex.Split(s, "\\s+", RegexOptions.None);
                        foreach (string p in space)
                        {

                            string p_replace = p.Replace("\"", "");
                            if (j == 0)
                            {
                                if (arrayXaxis.Contains(p) == false && p_replace != "X-location,Z-location,Resistivity")
                                {
                                    comboX.Items.Add(p);
                                    TempList.Add(p);

                                }
                                //Console.WriteLine(i + "/" + j + "/" + p);
                                arrayXaxis[i] = p;
                                fileArray[i, j] = p;
                                //Console.WriteLine(fileArray[i, j]);

                                j++;
                            }
                            else if (j == 1)
                            {
                                if (arrayYaxis.Contains(p) == false)
                                {
                                    comboY.Items.Add(p);
                                }
                                arrayYaxis[i] = p;
                                fileArray[i, j] = p;
                                j++;
                            }
                            else if (j == 2)
                            {
                                Console.WriteLine(p);
                                //find highest value for comboSpace
                                if (Convert.ToDouble(p) > maxStation)
                                {
                                    //Console.WriteLine(station_);
                                    maxStation = Convert.ToInt32(Convert.ToDouble(p));
                                    Console.WriteLine("max" + maxStation);
                                }

                                
                                fileArray[i, j] = p;
                                j = -1;
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
                countline = i - 1;

            }
            catch (IOException)
            {
                MessageBox.Show("Please currently use by another proceess.");
            }

            return countline;
        }

        private int excelFileOperation( int a, string filelength, int countline)
        {
            //Console.WriteLine(filelength);
            using (var excelWorkbook = new XLWorkbook(filelength))
            {
                var ws = excelWorkbook.Worksheet(1);
                var nonEmptyDataRows = ws.RowsUsed().Count();
                int row = nonEmptyDataRows + countline;
                int m;
                //Console.WriteLine( a +" / " + row);
                coordinateArray[a] = row;
                lengthArray = row;

                //default value
                if (!comboX.Items.Contains("ALL"))
                {
                    comboX.Items.Add("ALL");
                }
                if (!comboY.Items.Contains("ALL"))
                {
                    comboY.Items.Add("ALL");
                }

                int counter = 0;
                for (int n = 2; n < nonEmptyDataRows; n++)
                {
                    m = countline + counter;
                    counter++;
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
                        //////Console.WriteLine(x);
                    }
                    arrayXaxis[m] = x;
                    fileArray[m, 0] = x;
                    //Console.WriteLine(n);

                    if (arrayYaxis.Contains(y) == false)
                    {
                        comboY.Items.Add(y);
                    }

                    //find max space value
                    if (Convert.ToDouble(z) > maxStation)
                    {
                        //Console.WriteLine(station_);
                        maxStation = Convert.ToInt32(Convert.ToDouble(z));
                        Console.WriteLine("max" + maxStation);
                    }
                    arrayYaxis[m] = y;
                    fileArray[m, 1] = y;
                    fileArray[m, 2] = z;
                    //Boolean cellDouble = (Boolean)cellBoolean.Value;
                    //Console.WriteLine(x);
                }

                countline = row;
                
            }
            return countline;
        }

        private void clearBtn_Click(object sender, RoutedEventArgs e)
        {
            comboX.Items.Clear();
            comboY.Items.Clear();
            comboSpace.Items.Clear();
            path1.Items.Clear();
            dataGrids = null;
            dataSP = null;
            dataGrid1.ItemsSource = null;
            //dataGrids.Clear();
        }

        protected virtual bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }

        private void Toggle_btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void sp_Click(object sender, RoutedEventArgs e)
        {
            if (Properties.Settings.Default.SP_Setting)
            {
                Properties.Settings.Default.SP_Setting = false;
                toggle_btn.Background = Brushes.PaleVioletRed;
                toggle_btn.Content = "OFF";

            }
            else
            {
                Properties.Settings.Default.SP_Setting = true;
                toggle_btn.Background = Brushes.LightGreen;
                toggle_btn.Content = "ON";
            }
        }
    }
}
