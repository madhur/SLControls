using excel_create.Controls;
using ExcelPart.PeopleService;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace ExcelPart.Controls
{
    public partial class PeopleChooser : UserControl
    {
        #region Event Handler
        ContextMenu cMenu;
        MenuItem mnuItem;


        #endregion

        private const string siteUrl = "https://teams.aexp.com/sites/excel/";
        private const string peopleWsUrl = "/_vti_bin/People.asmx";

        public SelectedAccounts selectedAccounts;
        public bool AllowMultiple { get; set; }
        PPLPicker peoplePicker;
        Dictionary<String, PickerEntry> values;

        public PeopleChooser()
        {

            this.Loaded += PeopleChooser_Loaded;
            InitializeComponent();
            peoplePicker = new PPLPicker();
           
            peoplePicker.SubmitClicked += peoplePicker_SubmitClicked;
            selectedAccounts = new SelectedAccounts();

        }

        void PeopleChooser_Loaded(object sender, RoutedEventArgs e)
        {
            if (AllowMultiple)
            {
                UsersListBox.Visibility = System.Windows.Visibility.Visible;
                UserTextBox.Visibility = System.Windows.Visibility.Collapsed;
                ResolveButton.Visibility = Visibility.Collapsed;

                UsersListBox.DataContext = selectedAccounts;
                UsersListBox.ItemsSource = selectedAccounts;

            }
            else
            {
                UsersListBox.Visibility = System.Windows.Visibility.Collapsed;
                UserTextBox.Visibility = System.Windows.Visibility.Visible;
                ResolveButton.Visibility = Visibility.Visible;
            }

            peoplePicker.AllowMultiple = AllowMultiple;

        }
        
        void peoplePicker_SubmitClicked(object sender, EventArgs e)
        {
            selectedAccounts.Clear();

            foreach (AccountList ac in peoplePicker.selectedAccounts)
            {
                selectedAccounts.Add(new AccountList(ac.AccountName, ac.DisplayName));
            }

            if (!AllowMultiple && selectedAccounts.Count>0)
            {
                UserTextBox.Text = selectedAccounts[0].DisplayName;

            }
            
        }

        private void ResolveButton_Click(object sender, RoutedEventArgs e)
        {

            if (string.IsNullOrEmpty(UserTextBox.Text))
            {
                MessageBox.Show("You must enter a search term.", "Missing Search Term",
                    MessageBoxButton.OK);
                UserTextBox.Focus();
                return;
            }
            try
            {
                this.Cursor = Cursors.Wait;

                PeopleSoapClient ps = new PeopleSoapClient();
                //use the host name property to configure the request against the site in 
                //which the control is hosted
                ps.Endpoint.Address =
               new System.ServiceModel.EndpointAddress(siteUrl+peopleWsUrl);

                //create the handler for when the call completes
                ps.SearchPrincipalsCompleted += new EventHandler<SearchPrincipalsCompletedEventArgs>(ps_SearchPrincipalsCompleted);
                //execute the search
                ps.SearchPrincipalsAsync(UserTextBox.Text, 50, SPPrincipalType.User);
            }
            catch (Exception ex)
            {
                //ERROR LOGGING HERE
                Debug.WriteLine(ex.Message);

                MessageBox.Show("There was a problem executing the search; please try again " +
                     "later.", "Search Error",
                    MessageBoxButton.OK);
                //reset cursor
                this.Cursor = Cursors.Arrow;
            }

        }


        void ps_SearchPrincipalsCompleted(object sender, SearchPrincipalsCompletedEventArgs e)
        {
            try
            {
                if (e.Error != null)
                    MessageBox.Show("An error was returned: " + e.Error.Message, "Search Error",
                       MessageBoxButton.OK);
                else
                {
                    System.Collections.ObjectModel.ObservableCollection<PrincipalInfo>
                        results = e.Result;

                    if (e.Result.Count == 0)
                    {
                        MessageBox.Show("No match was found", "No match was found",
                  MessageBoxButton.OK);

                    }
                    else if (e.Result.Count > 1)
                    {
                        values = new Dictionary<string, PickerEntry>();

                        foreach (PrincipalInfo pi in results)
                        {
                            String decodedAccount = Utils.checkClaimsUser(pi.AccountName);
                            if (!values.ContainsKey(decodedAccount))
                                values.Add(decodedAccount, new PickerEntry(pi.DisplayName, decodedAccount, pi.Email, pi.Department));
                        }

                        if (values.Count == 1)
                        {
                            SetSingleResult(values);
                        }
                        else
                        {
                            MessageBox.Show("There was more than one match found", "People chooser",
MessageBoxButton.OK);
                        }

                    }
                    else if (e.Result.Count == 1)
                    {
                        values = new Dictionary<string, PickerEntry>();

                        foreach (PrincipalInfo pi in results)
                        {
                            String decodedAccount = Utils.checkClaimsUser(pi.AccountName);
                            if (!values.ContainsKey(decodedAccount))
                                values.Add(decodedAccount, new PickerEntry(pi.DisplayName, decodedAccount, pi.Email, pi.Department));
                        }

                        SetSingleResult(values);
                      
                    }
                    //clear the search results listbox
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was an error processing the search results: " + ex.Message,
                   "Search Error", MessageBoxButton.OK);
            }
            finally
            {
                //reset cursor
                this.Cursor = Cursors.Arrow;
            }
        }

        private void SetSingleResult(Dictionary<String, PickerEntry> values)
        {
            PickerEntry pi = values.Values.ToArray<PickerEntry>()[0];

            UserTextBox.Text = pi.DisplayName;

            selectedAccounts.Clear();
            selectedAccounts.Add(new AccountList(pi.AccountName, pi.DisplayName));


        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            peoplePicker.Show();

            // Restore the selected accounts to people picker
            if (selectedAccounts.Count > 0)
            {
                peoplePicker.selectedAccounts.Clear();

                foreach (AccountList ac in selectedAccounts)
                {
                   
                    peoplePicker.selectedAccounts.Add(new AccountList(ac.AccountName, ac.DisplayName));

                }

            }
        }

        private void ShowModalDialog()
        {
            AutoResetEvent waitHandle = new AutoResetEvent(false);
            Dispatcher.BeginInvoke(() =>
            {
                ChildWindow cw = new ChildWindow();
                cw.Content = "Modal Dialog";
                cw.Closed += (s, e) => waitHandle.Set();
                cw.Show();
            });
            waitHandle.WaitOne();
        }

        private void UserTextBox_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (values.Count > 1)
                e.Handled = true;
            else
                e.Handled = false;
        }

        private void UserTextBox_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            // Display a menu on mouse right button up
            cMenu = new ContextMenu();
            
            foreach (PickerEntry pi in values.Values)
            {
                mnuItem = new MenuItem();
                mnuItem.Header = pi.DisplayName;
                mnuItem.Name = pi.AccountName;
                mnuItem.Click += mnuItem_Click;
                cMenu.Items.Add(mnuItem);

            }

            cMenu.IsOpen = true;
        }

        void mnuItem_Click(object sender, RoutedEventArgs e)
        {
            MenuItem mnu = sender as MenuItem;

            values = new Dictionary<string, PickerEntry>();
            values.Add(mnu.Name, new PickerEntry(mnu.Header.ToString(), mnu.Name, string.Empty, string.Empty));

            SetSingleResult(values);
        }
    }
}
