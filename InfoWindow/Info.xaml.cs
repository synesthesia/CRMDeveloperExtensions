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
            string url = "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=KGV72FKEY8TJL";

            Process.Start(new ProcessStartInfo(url));
            e.Handled = true;
        }
    }
}
