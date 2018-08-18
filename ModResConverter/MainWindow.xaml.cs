using Microsoft.Win32;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;

namespace ModResConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //StreamReader objInput = null;
        string contents, filePath, dataValue, dataPrint;
        int maxValue;



        public MainWindow()
        {
            InitializeComponent();
            this.Title = "ModResConverter";
            //Application.Current.MainWindow.WindowState = WindowState.Maximized;
            export.IsEnabled = false;
            comboX.IsEnabled = false;
        }

        private void btn1_Click(object sender, RoutedEventArgs e)
        {
            // do something
            comboX.Items.Clear();
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

                int i = 0, countRow = 0;

                path1.Text = filePath;

                StreamReader objInput = new StreamReader(filePath, System.Text.Encoding.Default);
                contents = objInput.ReadToEnd().Trim();
                string[] split = System.Text.RegularExpressions.Regex.Split(contents, "\\r+", RegexOptions.None);
                foreach (string s in split)
                {
                    if (i < 1)
                    {
                        string[] space = System.Text.RegularExpressions.Regex.Split(s, "\\s+", RegexOptions.None);
                        foreach (string p in space)
                        {
                            countRow++;
                            comboX.Items.Add(countRow);
                            //combo1.SelectedIndex = 1;
                            maxValue = countRow;
                        }
                    }

                    i++;
                }

                comboX.IsEnabled = true;
            }
        }

        private void combo1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int selectedIndex = comboX.SelectedIndex;

            int i = 0;
            dataValue = null;
            int localMaxValue = maxValue - 1;
            string[] splits = System.Text.RegularExpressions.Regex.Split(contents, "\\s+", RegexOptions.None);
            foreach (string s in splits)
            {

                //Console.WriteLine(selectedIndex + "/" + i + "/" + localMaxValue) ;
                if (i == selectedIndex)
                {

                    dataValue = dataValue + s + "\r";
                    dataPrint = dataPrint + s + Environment.NewLine;
                    i++;

                    //last value
                    if (selectedIndex == localMaxValue)
                    {
                        i = 0;
                        //Console.WriteLine("skip");
                    }
                }
                else if (i == localMaxValue)
                {
                    i = 0;
                    //Console.WriteLine("skip");
                }
                else
                {
                    //default
                    i++;
                }


            }

            box1.Text = dataValue;
            export.IsEnabled = true;
        }

        private void export_Click(object sender, RoutedEventArgs e)
        {

            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = "Document"; // Default file name
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Spreedsheet documents (.xlsx)|*.xlsx";

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {

                // Save document
                string filename = dlg.FileName;
                //Console.WriteLine(filename);

                System.IO.File.WriteAllText(filename, dataPrint);
                System.Windows.MessageBox.Show("Succesfully exported");
            }
        }
    }
}
