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

namespace AOKMPO_BD_Sensors.Views
{
    /// <summary>
    /// Логика взаимодействия для ExpiredSensorsDialog.xaml
    /// </summary>
    public partial class ExpiredSensorsDialog : Window
    {
        public ExpiredSensorsDialog(List<Sensor> approachingSensors)
        {
            InitializeComponent();
            SensorsGrid.ItemsSource = approachingSensors;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

    }
}
