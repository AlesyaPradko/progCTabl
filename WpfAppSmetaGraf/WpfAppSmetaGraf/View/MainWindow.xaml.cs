using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;




namespace WpfAppSmetaGraf
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
  
        }

        private void People_Checked(object sender, RoutedEventArgs e)
        {
            RadioButton rb = sender as RadioButton;

            if (rb != null)
            {
                string Name = rb.Name;
                switch (Name)
                {
                    case "People":
                        ComboPeople.Visibility = Visibility.Visible;
                        LabPeople.Visibility = Visibility.Visible;
                        ComboDays.Visibility = Visibility.Collapsed;
                        LabDay.Visibility = Visibility.Collapsed;
                        break;
                    case "Days":
                        ComboDays.Visibility = Visibility.Visible;
                        LabDay.Visibility = Visibility.Visible;
                        ComboPeople.Visibility = Visibility.Collapsed;
                        LabPeople.Visibility = Visibility.Collapsed;
                        break;
                }
            }
        }
    }
}
