using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;


namespace excel_create.Controls
{
    public partial class close : ChildWindow
    {
        public close()
        {
            InitializeComponent();
        }

        private void Yesbtn_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            Uri redirect = new Uri("https://teams.aexp.com/sites/excel/SitePages/manage.aspx");
            System.Windows.Browser.HtmlPage.Window.Navigate(redirect, "_parent");
            
        }

        private void Nobtn_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }
    }
}

