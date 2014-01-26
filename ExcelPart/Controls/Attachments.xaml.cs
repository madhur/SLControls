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
using Microsoft.SharePoint.Client;
using System.Threading;
using System.IO;


namespace excel_create.Controls
{
    public partial class Attachments : ChildWindow
    {
        private ClientContext myClContext;
            
        public Attachments()
        {
            InitializeComponent();

            ConnectToSP();
        }

        private void FileListBox_Drop(object sender, DragEventArgs e)
        {

            if (e.Data == null)
                return;

           
                IDataObject dataObject = e.Data as IDataObject;
                FileInfo[] files = dataObject.GetData(DataFormats.FileDrop) as FileInfo[];

               //  InfoList listDetails = row.DataContext as InfoList;
                foreach (FileInfo file in files)
                {
                    UploadFile(file, "Shared Documents");
                }
        }


        private void UploadFile(FileInfo fileToUpload, string libraryTitle)
        {
            var web = myClContext.Web;
            List destinationList = web.Lists.GetByTitle(libraryTitle);

            var fciFileToUpload = new FileCreationInformation();

            Stream streamToUpload = fileToUpload.OpenRead();
            int length = (int)streamToUpload.Length;  // get file length

            fciFileToUpload.Content = new byte[length];

            int count = 0;                        // actual number of bytes read
            int sum = 0;                          // total number of bytes read

            while ((count = streamToUpload.Read(fciFileToUpload.Content, sum, length - sum)) > 0)
                sum += count;  // sum is a buffer offset for next reading
            streamToUpload.Close();

            fciFileToUpload.Url = fileToUpload.Name;

            Microsoft.SharePoint.Client.File clFileToUpload = destinationList.RootFolder.Files.Add(fciFileToUpload);

            myClContext.Load(clFileToUpload);
            myClContext.ExecuteQueryAsync(OnLoadingSucceeded, OnLoadingFailed);
            busyIndicatorElement.IsBusy = true;
        }

        private void ConnectToSP()
        {
            myClContext = new ClientContext("https://teams.aexp.com/sites/excel/");

           /* myClContext.Load(myClContext.Web);
            myClContext.Load(myClContext.Web.Lists);

            myClContext.ExecuteQueryAsync(OnConnectSucceeded, OnConnectFailed);
            busyIndicatorElement.IsBusy = true;*/
        }

        public void CreateFolder(string siteUrl, string listName, string relativePath, string folderName)
        {
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                Web web = clientContext.Web;
                List list = web.Lists.GetByTitle(listName);

                ListItemCreationInformation newItem = new ListItemCreationInformation();
                newItem.UnderlyingObjectType = FileSystemObjectType.Folder;
                newItem.FolderUrl = siteUrl + "/lists/" + listName;
                if (!relativePath.Equals(string.Empty))
                {
                    newItem.FolderUrl += "/" + relativePath;
                }
                newItem.LeafName = folderName;
                ListItem item = list.AddItem(newItem);
                item.Update();
                clientContext.ExecuteQuery();
            }
        }


        public void SearchFolder(string siteUrl, string listName, string relativePath)
        {
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                Web web = clientContext.Web;
                List list = web.Lists.GetByTitle(listName);

                string FolderFullPath = null;

                CamlQuery query = CamlQuery.CreateAllFoldersQuery();

                if (relativePath.Equals(string.Empty))
                {
                    FolderFullPath = "/lists/" + listName;
                }
                else
                {
                    FolderFullPath = "/lists/" + listName + "/" + relativePath;
                }
                if (!string.IsNullOrEmpty(FolderFullPath))
                {
                    query.FolderServerRelativeUrl = FolderFullPath;
                }
                IList<Folder> folderResult = new List<Folder>();

                var listItems = list.GetItems(query);

                clientContext.Load(list);
                clientContext.Load(listItems, litems => litems.Include(
                    li => li["DisplayName"],
                    li => li["Id"]
                    ));

                clientContext.ExecuteQuery();

                foreach (var item in listItems)
                {

                    Console.WriteLine("{0}----------{1}", item.Id, item.DisplayName);
                }
            }
        }

        public void DeleteFolder(string siteUrl, string listName, string relativePath, string folderName)
        {
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                Web web = clientContext.Web;
                List list = web.Lists.GetByTitle(listName);

                CamlQuery query = new CamlQuery();
                query.ViewXml = "<View Scope=\"RecursiveAll\"> " +
                                "<Query>" +
                                    "<Where>" +
                                        "<And>" +
                                            "<Eq>" +
                                                "<FieldRef Name=\"FSObjType\" />" +
                                                "<Value Type=\"Integer\">1</Value>" +
                                             "</Eq>" +
                                              "<Eq>" +
                                                "<FieldRef Name=\"Title\"/>" +
                                                "<Value Type=\"Text\">" + folderName + "</Value>" +
                                              "</Eq>" +
                                        "</And>" +
                                     "</Where>" +
                                "</Query>" +
                                "</View>";

                if (relativePath.Equals(string.Empty))
                {
                    query.FolderServerRelativeUrl = "/lists/" + listName;
                }
                else
                {
                    query.FolderServerRelativeUrl = "/lists/" + listName + "/" + relativePath;
                }

                var folders = list.GetItems(query);

                clientContext.Load(list);
                clientContext.Load(folders);
                clientContext.ExecuteQuery();
                if (folders.Count == 1)
                {
                    folders[0].DeleteObject();
                    clientContext.ExecuteQuery();
                }
            }
        }


        public void RenameFolder(string siteUrl, string listName, string relativePath, string folderName, string folderNewName)
        {
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                Web web = clientContext.Web;
                List list = web.Lists.GetByTitle(listName);

                string FolderFullPath = GetFullPath(listName, relativePath, folderName);

                CamlQuery query = new CamlQuery();
                query.ViewXml = "<View Scope=\"RecursiveAll\"> " +
                                "<Query>" +
                                    "<Where>" +
                                        "<And>" +
                                            "<Eq>" +
                                                "<FieldRef Name=\"FSObjType\" />" +
                                                "<Value Type=\"Integer\">1</Value>" +
                                             "</Eq>" +
                                              "<Eq>" +
                                                "<FieldRef Name=\"Title\"/>" +
                                                "<Value Type=\"Text\">" + folderName + "</Value>" +
                                              "</Eq>" +
                                        "</And>" +
                                     "</Where>" +
                                "</Query>" +
                                "</View>";

                if (relativePath.Equals(string.Empty))
                {
                    query.FolderServerRelativeUrl = "/lists/" + listName;
                }
                else
                {
                    query.FolderServerRelativeUrl = "/lists/" + listName + "/" + relativePath;
                }
                var folders = list.GetItems(query);

                clientContext.Load(list);
                clientContext.Load(list.Fields);
                clientContext.Load(folders, fs => fs.Include(fi => fi["Title"],
                    fi => fi["DisplayName"],
                    fi => fi["FileLeafRef"]));
                clientContext.ExecuteQuery();

                if (folders.Count == 1)
                {

                    folders[0]["Title"] = folderNewName;
                    folders[0]["FileLeafRef"] = folderNewName;
                    folders[0].Update();
                    clientContext.ExecuteQuery();
                }
            }
        }
       

       
    }
}

