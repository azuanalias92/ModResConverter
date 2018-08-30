using Microsoft.Win32;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Linq;
using System.Data;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace ModResConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //StreamReader objInput = null;
        string contents, filePath, dataValue, dataPrint;
        string[,] fileArray;
        List<GridData> dataGrids;

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
            Microsoft.Win32.OpenFileDialog fileDialog = new Microsoft.Win32.OpenFileDialog();
            fileDialog.DefaultExt = ".dat"; // Required file extension 
            fileDialog.Filter = "DAT file (.dat)|*.dat"; // Optional file extensions

            if (fileDialog.ShowDialog() == true)
            {
                filePath = fileDialog.FileName;
                int i = 0, j = 0;
                

                try
                {
                    StreamReader objInput = new StreamReader(filePath, System.Text.Encoding.Default);
                    contents = objInput.ReadToEnd().Trim();
                    string[] split = System.Text.RegularExpressions.Regex.Split(contents, "\\r+", RegexOptions.None);
                    fileArray = new string[split.Length, 4];
                    string[] arrayYaxis = new string[split.Length];
                    string[] arrayXaxis = new string[split.Length];
                    //string[] arrayZaxis = new string[split.Length];
                    foreach (string s in split)
                    {

                        //Console.WriteLine(s);
                        if (i > 0)
                        {
                            string[] space = System.Text.RegularExpressions.Regex.Split(s, "\\s+", RegexOptions.None);
                            foreach (string p in space)
                            {
                                //Console.WriteLine(i + "/" + p);
                                if (j == 1)
                                {
                                    if (arrayXaxis.Contains(p) == false)
                                    {
                                        comboX.Items.Add(p);
                                    }
                                    arrayXaxis[i] = p;
                                    fileArray[i,j] = p;
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

                        i++;
                        
                    }

                    //Console.WriteLine(i);

                    comboX.IsEnabled = true;
                    comboY.IsEnabled = true;
                    path1.Text = filePath;
                    
                }
                catch (IOException)
                {
                    System.Windows.Forms.MessageBox.Show("Please currently use by another proceess.");
                }
                
                

            }
        }

        private void comboX_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            String selectedValue = (String)comboX.SelectedValue;
            //Console.WriteLine(selectedValue);
            dataGrids = new List<GridData>();


            dataValue = "X Y Z \r";
            for (int k = 1; k < fileArray.GetLength(0); k++)
            {
                if (fileArray[k, 1] == selectedValue)
                {
                    dataGrids.Add(new GridData()
                    {
                        X = fileArray[k, 1],
                        Y = fileArray[k, 2],
                        Z = fileArray[k, 3]
                    });
                    //dataValue = dataValue + fileArray[k, 1] + " "  + fileArray[k, 2] + " "  + fileArray[k, 3] + "\r";
                    dataPrint = dataPrint + fileArray[k, 3] + Environment.NewLine;
                }
            }

           

            dataGrid1.ItemsSource = dataGrids;
            export.IsEnabled = true;
        }

        private void comboY_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            String selectedValue = (String)comboY.SelectedValue;
            dataValue = null;

            dataGrids = new List<GridData>();

            for (int k = 1; k < fileArray.GetLength(0); k++)
            {
                if (fileArray[k, 2] == selectedValue)
                {
                    dataGrids.Add(new GridData()
                    {
                        X = fileArray[k, 1],
                        Y = fileArray[k, 2],
                        Z = fileArray[k, 3]
                    });
                    dataValue = dataValue + fileArray[k, 3] + "\r";
                    dataPrint = dataPrint + fileArray[k, 3] + Environment.NewLine;
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
                    worksheet.Cell("A1").Value = "X";
                    worksheet.Cell("B1").Value = "Y";
                    worksheet.Cell("C1").Value = "Z";
                    foreach (GridData GridData in dataGrids)
                    {
                        worksheet.Cell("A" + row.ToString()).Value = GridData.X.ToString();
                        worksheet.Cell("B" + row.ToString()).Value = GridData.Y.ToString();
                        worksheet.Cell("C" + row.ToString()).Value = GridData.Z.ToString();
                        row++;

                    }
                    workbook.SaveAs(saveFileDialog.FileName);

                    System.Windows.MessageBox.Show("Successfully Export");
                }
            }
        }
    }
}
