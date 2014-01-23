using System;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace excel_create.Controls
{
    public class AccountList
    {
        public String AccountName { get; set; }
        public String DisplayName { get; set; }

        public AccountList(String acName, String dispName)
        {
            this.AccountName = acName;
            this.DisplayName = dispName;

        }

        public override string ToString()
        {
            return DisplayName;
        }
    }
}
