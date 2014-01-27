using excel_create.Controls;
using ExcelPart.Controls;
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

        public MainPage()
        {
            InitializeComponent();
            this.Loaded += MainPage_Loaded;
        }

        void MainPage_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void LoadUser(ClientContext ctx, FieldUserValue singleValue, FieldUserValue[] multValue)
        {
            List userList = ctx.Web.SiteUserInfoList;
            ctx.Load(userList);

            ListItemCollection users = userList.GetItems(CamlQuery.CreateAllItemsQuery());

            ctx.Load(users, items => items.Include(
                item => item.Id, item => item["Name"]));



            ctx.ExecuteQueryAsync((ss, eee) =>
            {
                ListItem principal = users.GetById(singleValue.LookupId);

                ctx.Load(principal);



                ctx.ExecuteQueryAsync((sss, eeee) =>
                {
                    string username = principal["Name"] as string;

                    string decodedName = Utils.checkClaimsUser(username);
                    string dispName = principal["Title"] as string;

                    Dispatcher.BeginInvoke(() =>
{
    SinglePeopleChooser.selectedAccounts.Clear();

    SinglePeopleChooser.selectedAccounts.Add(new AccountList(decodedName, dispName));
    SinglePeopleChooser.UserTextBox.Text = dispName;

}
    );

                },
                  (sss, eeee) =>
                  {
                      Console.WriteLine(eeee.Message);

                  });


            },
             (sss, eeee) =>
             {
                 Console.WriteLine(eeee.Message);

             });



            userList = ctx.Web.SiteUserInfoList;
            ctx.Load(userList);

            users = userList.GetItems(CamlQuery.CreateAllItemsQuery());

            ctx.Load(users, items => items.Include(
                item => item.Id, item => item["Name"]));


            ctx.ExecuteQueryAsync((s, ee) =>
            {
                ListItem[] principals = new ListItem[multValue.Length];

                for (int i = 0; i < multValue.Length; i++)
                {
                    principals[i] = users.GetById(multValue[i].LookupId);
                    ctx.Load(principals[i]);
                }

                ctx.ExecuteQueryAsync((ssss, eeeee) =>
                {
                    string username;

                    for (int i = 0; i < multValue.Length; i++)
                    {


                        try
                        {
                            username = principals[i]["Name"] as string;
                        }
                        catch (IndexOutOfRangeException ii)
                        {
                            return;
                        }

                        string decodedName = Utils.checkClaimsUser(username);
                        string dispName = principals[i]["Title"] as string;

                        Dispatcher.BeginInvoke(() =>
                        {


                            MultiplePeopleChooser.selectedAccounts.Add(new AccountList(decodedName, dispName));


                        }
                        );
                    }


                },
               (ssss, eeeee) =>
               {
                   Console.WriteLine(eeeee.Message);

               });





            },

             (ssss, eeeee) =>
             {
                 Console.WriteLine(eeeee.Message);

             });

        }

        private void SubmitButton_Click(object sender, RoutedEventArgs e)
        {

            User Singleuser;

            if (SinglePeopleChooser.selectedAccounts.Count > 0 || MultiplePeopleChooser.selectedAccounts.Count > 0)
            {
                ClientContext context = new ClientContext("https://teams.aexp.com/sites/excel/");
                List MadhurList = context.Web.Lists.GetByTitle("Madhur");
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



                context.ExecuteQueryAsync((s, ee) =>
                {


                    Dispatcher.BeginInvoke(() =>
                    {

                        MessageBox.Show("Success", "Success", MessageBoxButton.OK);
                    }
                        );


                },
         (s, ee) =>
         {
             Console.WriteLine(ee.Message);

         });


            }




        }



        private void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            ClientContext ctx = new ClientContext("https://teams.aexp.com/sites/excel/");

            Web web = ctx.Web;

            ctx.Load(web);

            List list = ctx.Web.Lists.GetByTitle("Madhur");
            ctx.Load(list);

            ListItem targetItem = list.GetItemById(1);

            ctx.Load(targetItem);

            ctx.ExecuteQueryAsync((s, ee) =>
            {

                FieldUserValue singleValue = (FieldUserValue)targetItem["Single"];
                FieldUserValue[] multValue = targetItem["Multiple"] as FieldUserValue[];

                Dispatcher.BeginInvoke(() =>
                {
                    LoadUser(ctx, singleValue, multValue);

                }
                    );


            },
           (s, ee) =>
           {
               Console.WriteLine(ee.Message);

           });


        }


    }
}
