using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ModResConverter
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Settings : Window
    {
        public Settings()
        {
            InitializeComponent();
            this.Title = "Settings";


            //mode
            if (Properties.Settings.Default.SP_Setting)
            {
                toggle_btn.Background = Brushes.LightGreen;
                toggle_btn.Content = "ON";
            }
            else
            {
                toggle_btn.Background = Brushes.PaleVioletRed;
                toggle_btn.Content = "OFF";
            }

            //checkbox
            serial_cb.IsChecked     = Properties.Settings.Default.Serial_number;
            date_cb.IsChecked       = Properties.Settings.Default.Date;
            line_cb.IsChecked       = Properties.Settings.Default.Lines;
            station_cb.IsChecked    = Properties.Settings.Default.Station;
            north_cb.IsChecked      = Properties.Settings.Default.Northing;
            east_cb.IsChecked       = Properties.Settings.Default.Easting;
            second_cb.IsChecked     = Properties.Settings.Default.Second;
            minute_cb.IsChecked     = Properties.Settings.Default.Minute;
            reading1_cb.IsChecked = Properties.Settings.Default.Reading_1;
            reading2_cb.IsChecked = Properties.Settings.Default.Reading_2;
            reading3_cb.IsChecked = Properties.Settings.Default.Reading_3;
            reading4_cb.IsChecked = Properties.Settings.Default.Reading_4;
            average_cb.IsChecked = Properties.Settings.Default.Average;
            elevation_cb.IsChecked = Properties.Settings.Default.Elevation;
            x_cb.IsChecked = Properties.Settings.Default.X;
            y_cb.IsChecked = Properties.Settings.Default.Y;
            remark_cb.IsChecked = Properties.Settings.Default.Remarks;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
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

        private void CheckBoxChanged(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Serial_number   = (bool)serial_cb.IsChecked;
        
        }

        private void CheckBoxChanged1(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Date            = (bool)date_cb.IsChecked;
        }

        private void CheckBoxChanged2(object sender, RoutedEventArgs e)
        {

            Properties.Settings.Default.Lines = (bool)line_cb.IsChecked;
        }

        private void CheckBoxChanged3(object sender, RoutedEventArgs e)
        {

            Properties.Settings.Default.Station = (bool)station_cb.IsChecked;
            
        }

        private void CheckBoxChanged4(object sender, RoutedEventArgs e)
        {
          
            Properties.Settings.Default.Northing = (bool)north_cb.IsChecked;
           
        }

        private void CheckBoxChanged5(object sender, RoutedEventArgs e)
        {
            
            Properties.Settings.Default.Easting = (bool)east_cb.IsChecked;
           
        }

        private void CheckBoxChanged6(object sender, RoutedEventArgs e)
        {

            Properties.Settings.Default.Second = (bool)second_cb.IsChecked;
        }
           
        private void CheckBoxChanged7(object sender, RoutedEventArgs e)
        {
            
            Properties.Settings.Default.Minute = (bool)minute_cb.IsChecked;
            
        }

        private void CheckBoxChanged8(object sender, RoutedEventArgs e)
        {

            Properties.Settings.Default.Reading_1 = (bool)reading1_cb.IsChecked;
           
        }

        private void CheckBoxChanged9(object sender, RoutedEventArgs e)
            {

                Properties.Settings.Default.Reading_2 = (bool)reading2_cb.IsChecked;
            }

        private void CheckBoxChanged10(object sender, RoutedEventArgs e)
        {

            Properties.Settings.Default.Reading_3 = (bool)reading3_cb.IsChecked;

        }

        private void CheckBoxChanged11(object sender, RoutedEventArgs e)
        {

            Properties.Settings.Default.Reading_4 = (bool)reading4_cb.IsChecked;

        }

        private void CheckBoxChanged12(object sender, RoutedEventArgs e)
        {
            
            Properties.Settings.Default.Average = (bool)average_cb.IsChecked;

        }

        private void CheckBoxChanged13(object sender, RoutedEventArgs e)
        {
           
            Properties.Settings.Default.Elevation = (bool)elevation_cb.IsChecked;

        }

        private void CheckBoxChanged14(object sender, RoutedEventArgs e)
        {

            Properties.Settings.Default.X = (bool)x_cb.IsChecked;

        }

        private void CheckBoxChanged15(object sender, RoutedEventArgs e)
            {
                Properties.Settings.Default.Serial_number = (bool)serial_cb.IsChecked;

                Properties.Settings.Default.Y = (bool)y_cb.IsChecked;

            }

        private void CheckBoxChanged16(object sender, RoutedEventArgs e)
            {

               
                Properties.Settings.Default.Remarks = (bool)remark_cb.IsChecked;
            }
        }
}
