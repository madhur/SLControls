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
		private const string siteUrl = "https://teams.aexp.com/sites/excel/";

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
				ClientContext context = new ClientContext(siteUrl);
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
			ClientContext ctx = new ClientContext(siteUrl);

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

		private void SubmitFilesButton_Click(object sender, RoutedEventArgs e)
		{
			RenameFolder(siteUrl, "Shared Documents", string.Empty, "Madhur", "MadhurNewFolder");
		}


		public void RenameFolder(string siteUrl, string listName, string relativePath, string folderName, string folderNewName)
		{
			using (ClientContext clientContext = new ClientContext(siteUrl))
			{
				Web web = clientContext.Web;
				List list = web.Lists.GetByTitle(listName);

				//  string FolderFullPath = GetFullPath(listName, relativePath, folderName);

				CamlQuery query = new CamlQuery();
				query.ViewXml = "<View Scope=\"RecursiveAll\"> " +
								"<Query>" +
									"<Where>" +
									   // "<And>" +
											"<Eq>" +
												"<FieldRef Name=\"FSObjType\" />" +
												"<Value Type=\"Integer\">1</Value>" +
											 "</Eq>" +
											 /* "<Eq>" +
												"<FieldRef Name=\"Title\"/>" +
												"<Value Type=\"Text\">" + folderName + "</Value>" +
											  "</Eq>" +*/
									   // "</And>" +
									 "</Where>" +
								"</Query>" +
								"</View>";

			   /* if (relativePath.Equals(string.Empty))
				{
					query.FolderServerRelativeUrl = "/lists/" + listName;
				}
				else
				{
					query.FolderServerRelativeUrl = "/lists/" + listName + "/" + relativePath;
				}*/

				//query.FolderServerRelativeUrl = "/"+listName;

				var folders = list.GetItems(query);

				clientContext.Load(list);
				clientContext.Load(list.Fields);
				clientContext.Load(folders, fs => fs.Include(fi => fi["Title"],
					fi => fi["DisplayName"],
					fi => fi["FileLeafRef"]));
			   // clientContext.ExecuteQuery();

				clientContext.ExecuteQueryAsync((s, ee) =>
				{

					if (folders.Count == 1)
					{

						folders[0]["Title"] = folderNewName;
						folders[0]["FileLeafRef"] = folderNewName;
						folders[0].Update();
						clientContext.ExecuteQueryAsync((ss, eee) =>
						{

							Dispatcher.BeginInvoke(() =>
							{

								MessageBox.Show("Success", "Success", MessageBoxButton.OK);
							});
						



						},
		  (ss, eee) =>
		  {
			  Console.WriteLine(eee.Message);

		  });



					}
				   


				},
		  (s, ee) =>
		  {
			  Console.WriteLine(ee.Message);

		  });

			   
			}
		}


	}
}
