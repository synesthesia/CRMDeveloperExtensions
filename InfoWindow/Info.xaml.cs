using System.Diagnostics;
using System.Windows;

namespace InfoWindow
{
    public partial class Info
    {
        public Info()
        {
            InitializeComponent();
        }

        private void Donate_OnClick(object sender, RoutedEventArgs e)
        {
            string url = "https://www.paypal.me/JLattimer";

            Process.Start(new ProcessStartInfo(url));
            e.Handled = true;
        }
    }
}
