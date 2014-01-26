using excel_create.Controls;
using Microsoft.SharePoint.Client;
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

namespace ExcelPart
{
    public partial class MainPage : UserControl
    {
        Web webSite;
        private List MadhurList;
        User user;

        public MainPage()
        {
            InitializeComponent();
            this.Loaded += MainPage_Loaded;
        }

        void MainPage_Loaded(object sender, RoutedEventArgs e)
        {
            ClientContext ctx = new ClientContext("http://teams.aexp.com/sites/excel/");

            Web web = ctx.Web;

            ctx.Load(web);

            List list = ctx.Web.Lists.GetByTitle("Madhur");
            ctx.Load(list);

            ListItem targetItem = list.GetItemById(1);

            ctx.Load(targetItem);

            ctx.ExecuteQueryAsync((s, ee) =>
            {

                FieldUserValue singleValue = (FieldUserValue)targetItem["Single"];

                User user = ctx.Web.EnsureUser(singleValue.LookupValue);

                ctx.Load(user);

                ctx.ExecuteQueryAsync((ss, eee) =>
                {
                    Dispatcher.BeginInvoke(() =>
                    {
                        SinglePeopleChooser.selectedAccounts.Clear();
                        SinglePeopleChooser.selectedAccounts.Add(new AccountList(user.LoginName, user.Title));

                    }
                        );

                },
          (ss, eee) =>
          {


          });



            },
           (s, ee) =>
           {


           });








        }

        private void second_Closed(object sender, EventArgs e)
        {

        }

        private void successuserget(object sender, ClientRequestSucceededEventArgs args)
        {




        }

        private void failureuserget(object sender, ClientRequestFailedEventArgs args) { }

        private void successget(object sender, ClientRequestSucceededEventArgs args) { }

        private void failureget(object sender, ClientRequestFailedEventArgs args) { }

        private void SubmitButton_Click(object sender, RoutedEventArgs e)
        {

            User Singleuser;
            // UserCollection multipleUsers;

            if (SinglePeopleChooser.selectedAccounts.Count > 0 || MultiplePeopleChooser.selectedAccounts.Count > 0)
            {
                ClientContext context = new ClientContext("http://teams.aexp.com/sites/excel/");
                MadhurList = context.Web.Lists.GetByTitle("Madhur");
                ListItem newItem = MadhurList.AddItem(new ListItemCreationInformation());

                if (SinglePeopleChooser.selectedAccounts.Count > 0)
                {
                    Singleuser = context.Web.EnsureUser(SinglePeopleChooser.selectedAccounts[0].AccountName);
                    newItem["Single"] = Singleuser;

                }
                if (MultiplePeopleChooser.selectedAccounts.Count > 0)
                {
                    List<FieldUserValue> usersList = new List<FieldUserValue>();
                    foreach (AccountList ac in MultiplePeopleChooser.selectedAccounts)
                    {
                        usersList.Add(FieldUserValue.FromUser(ac.AccountName));
                    }

                    newItem["Multiple"] = usersList;
                }

                newItem.Update();
                context.Load(MadhurList, list => list.Title);

                context.ExecuteQueryAsync(QuerySucceed, QueryFailed);


            }




        }

        private void QueryFailed(object sender, ClientRequestFailedEventArgs args)
        {
            // MessageBox.Show("error", "Error", MessageBoxButton.OK);
            throw args.Exception;
        }

        private void QuerySucceed(object sender, ClientRequestSucceededEventArgs args)
        {
            // MessageBox.Show("Submitted", "Error", MessageBoxButton.OK);

        }

        private void MultiplePeopleChooser_Loaded(object sender, RoutedEventArgs e)
        {

        }


    }
}
