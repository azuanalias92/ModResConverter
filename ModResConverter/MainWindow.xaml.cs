using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
//using System.Windows.Forms;
using System.Linq;
using System.Data;
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

        public MainWindow()
        {
            InitializeComponent();
            this.Title = "ModResConverter";
            //Application.Current.MainWindow.WindowState = WindowState.Maximized;
            export.IsEnabled = false;
            comboX.IsEnabled = false;
            comboY.IsEnabled = false;
        }

        private void btn1_Click(object sender, RoutedEventArgs e)
        {
            // do something
            comboX.Items.Clear();
            comboY.Items.Clear();
            OpenDialog();
        }

        private void OpenDialog()
        {
            path1.Items.Clear();
            lengthArray = 0;
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.DefaultExt = ".dat"; // Required file extension 
            fileDialog.Filter = "DAT Files |*.dat; | Excel Files|*.xls;*.xlsx;*.xlsm"; // Optional file extensions
            //fileDialog.Filter = ;
            fileDialog.Multiselect = true;

            if (fileDialog.ShowDialog() == true)
            {
                //Console.WriteLine(Path.GetExtension(fileDialog.FileName));
                String fileType = Path.GetExtension(fileDialog.FileName);
                if (fileType == ".dat")
                {
                    openDATFile(fileDialog);
                }
                else
                {
                    //MessageBox.Show("excel file");
                    openExcelFile(fileDialog);
                }
                

            }
        }

        private void openExcelFile(OpenFileDialog fileDialog)
        {
            dataSP = new List<GridSP>();
            string fileName = fileDialog.FileName;
            using (var excelWorkbook = new XLWorkbook(fileName))
            {
                var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();
                int i = 0;
                foreach (var dataRow in nonEmptyDataRows)
                {
                    if(i> 0)
                    {
                        //Console.WriteLine(dataRow.Cell(2).GetValue<string>());
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
                        string elevation_ = dataRow.Cell(14).GetValue<string>();
                        string x_ = dataRow.Cell(15).GetValue<string>();
                        string y_ = dataRow.Cell(16).GetValue<string>();
                        string remarks_ = dataRow.Cell(17).GetValue<string>();

                        dataSP.Add(new GridSP()
                        {
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
                            //average = ,
                            elevation = elevation_,
                            x = x_,
                            y = y_,
                            remarks = remarks_


                        });
                        //Console.WriteLine(date_);
                    }
                    i = 1;
                }
            }
            dataGrid1.ItemsSource = dataSP;


        }

        private void openDATFile(OpenFileDialog fileDialog)
        {
            //looping for length of each file
            int z = 1;
            foreach (string filelength in fileDialog.FileNames)
            {
                try
                {
                    StreamReader objInputs = new StreamReader(filelength, System.Text.Encoding.Default);
                    contents = objInputs.ReadToEnd().Trim();
                    string[] splits = System.Text.RegularExpressions.Regex.Split(contents, "\\r+", RegexOptions.None);
                    lengthArray = lengthArray + splits.Length;
                    Console.WriteLine(lengthArray);
                    coordinateArray[z] = lengthArray;
                    z++;
                }
                catch (IOException)
                {
                    MessageBox.Show("Please currently use by another proceess.");
                }

            }

            //do process
            int i = 0, j = 0;
            countLine = 1;
            fileArray = new string[lengthArray, 4];
            string[] arrayYaxis = new string[lengthArray];
            string[] arrayXaxis = new string[lengthArray];



            foreach (string filename in fileDialog.FileNames)
            {
                path1.Items.Add(Path.GetFileName(filename));
                //Console.WriteLine("------------------------------------------" + countLine);

                int skip = 0;
                try
                {
                    StreamReader objInput = new StreamReader(filename, System.Text.Encoding.Default);
                    contents = objInput.ReadToEnd().Trim();
                    string[] split = System.Text.RegularExpressions.Regex.Split(contents, "\\r+", RegexOptions.None);

                    foreach (string s in split)
                    {

                        //Console.WriteLine(s);
                        if (skip != 0)
                        {
                            string[] space = System.Text.RegularExpressions.Regex.Split(s, "\\s+", RegexOptions.None);
                            foreach (string p in space)
                            {
                                //Console.WriteLine(i + "/" + p);
                                string p_replace = p.Replace("\"", "");
                                if (j == 1)
                                {
                                    if (arrayXaxis.Contains(p) == false && p_replace != "X-location,Z-location,Resistivity")
                                    {
                                        comboX.Items.Add(p);
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

                }
                catch (IOException)
                {
                    MessageBox.Show("Please currently use by another proceess.");
                }
                countLine++;
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
            for (int k = 1; k < lengthArray ; k++)
            {

                if(coordinateArray[z] == k)
                {
                    z++;
                    lines = "Line " + z;
                    Console.WriteLine(coordinateArray[z]);
                    Console.WriteLine(lines);
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
            export.IsEnabled = true;
        }

        private void comboY_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            String selectedValue = (String)comboY.SelectedValue;
            dataGrids = new List<GridData>();

            for (int k = 1; k < lengthArray; k++)
            {
                //Console.WriteLine(k +"/"+ fileArray[k, 3]);
                if (fileArray[k, 2] == selectedValue)
                {

                    dataGrids.Add(new GridData()
                    {
                        line = "Line " + k,
                        X = fileArray[k, 1],
                        Y = fileArray[k, 2],
                        Z = fileArray[k, 3]
                    });
                }
            }

            dataGrid1.ItemsSource = dataGrids;
            export.IsEnabled = true;

        }

        private void export_Click(object sender, RoutedEventArgs e)
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
    }
}
