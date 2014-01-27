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
using System.Windows.Resources;
using ExcelPart.Controls;


namespace excel_create.Controls
{
    public partial class Attachments : UserControl
    {
        private const string siteUrl="https://teams.aexp.com/sites/excel/";
        private const string libName = "Shared Documents";

        private ClientContext myClContext;
        public SelectedFiles selectedFiles;            


        public Attachments()
        {
            InitializeComponent();
            selectedFiles = new SelectedFiles();

            ConnectToSP();

            FileListBox.DataContext = selectedFiles;
            FileListBox.ItemsSource = selectedFiles;

        }

        private void FileListBox_Drop(object sender, DragEventArgs e)
        {

            if (e.Data == null)
                return;

           
                IDataObject dataObject = e.Data as IDataObject;
                FileInfo[] files = dataObject.GetData(DataFormats.FileDrop) as FileInfo[];


               //  InfoList listDetails = row.DataContext as InfoList;

               /* CreateFolder(siteUrl, libName, string.Empty, "Madhur");

                foreach (FileInfo file in files)
                {
                    UploadFile(file, libName);
                }*/
        }


        private void UploadFile(FileInfo fileToUpload, string libraryTitle, string folderName)
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

            Microsoft.SharePoint.Client.File clFileToUpload = null;
            if (string.IsNullOrEmpty(folderName))
            {
                clFileToUpload = destinationList.RootFolder.Files.Add(fciFileToUpload);

                myClContext.Load(clFileToUpload);

                myClContext.ExecuteQueryAsync((s, ee) =>
                {

                    Dispatcher.BeginInvoke(() =>
                    {
                        selectedFiles.Add(new FileList(fileToUpload.Name, fileToUpload.Name));
                        MessageBox.Show("Success", "Success", MessageBoxButton.OK);
                    }
                        );



                },
                (s, ee) =>
                {
                    Console.WriteLine(ee.Message);

                });

            }
            else
            {
                FolderCollection folderCol=destinationList.RootFolder.Folders;
                //myClContext.Load(folderCol, items => items.Include(fldr => fldr.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase)));
                
                myClContext.Load(folderCol);
                

                myClContext.ExecuteQueryAsync((s, ee) =>
                {

                    for (int i = 0; i < folderCol.Count; ++i)
                    {
                        if (folderCol[i].Name.Equals(folderName, StringComparison.OrdinalIgnoreCase))
                        {
                            clFileToUpload = folderCol[i].Files.Add(fciFileToUpload);

                            myClContext.Load(clFileToUpload);
                            break;
                        }

                    }

                    myClContext.ExecuteQueryAsync((ss, eee) =>
                    {

                        Dispatcher.BeginInvoke(() =>
                        {
                            selectedFiles.Add(new FileList(fileToUpload.Name, fileToUpload.Name));
                            MessageBox.Show("Success", "Success", MessageBoxButton.OK);
                        }
                            );



                    },
              (ss, eee) =>
              {
                  Console.WriteLine(eee.Message);

              });



                },
              (s, ee) =>
              {
                  Console.WriteLine(ee.Message);

              });


              
            }

         
         
        }

        private void ConnectToSP()
        {
            myClContext = new ClientContext("https://teams.aexp.com/sites/excel/");

          
        }

        public void CreateFolder(string siteUrl, string listName, string relativePath, string folderName)
        {
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                Web web = clientContext.Web;
                List list = web.Lists.GetByTitle(listName);

                Folder rootFolder=list.RootFolder;

                clientContext.Load(rootFolder);



                ListItemCreationInformation newItem = new ListItemCreationInformation();
                newItem.UnderlyingObjectType = FileSystemObjectType.Folder;
                //newItem.FolderUrl = siteUrl + listName;
                if (!relativePath.Equals(string.Empty))
                {
                    newItem.FolderUrl += "/" + relativePath;
                }
                newItem.LeafName = folderName;
                ListItem item = list.AddItem(newItem);
                item.Update();

                clientContext.Load(list);

                clientContext.ExecuteQueryAsync((s, ee) =>
                {

                    Folder newFolder=rootFolder.Folders.Add(folderName);


                    Dispatcher.BeginInvoke(() =>
                    {

                        MessageBox.Show("Created", "Created", MessageBoxButton.OK);
                    });

                   /* clientContext.ExecuteQueryAsync((ss, eee) =>
                    {

                        Dispatcher.BeginInvoke(() =>
                        {

                            MessageBox.Show("Created", "Created", MessageBoxButton.OK);
                        }
                    );



                    },
        (ss, eee) =>
        {
            Console.WriteLine(eee.Message);

        });*/

                },
          (s, ee) =>
          {
              Console.WriteLine(ee.Message);

          });
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

        private void FileUpload_Click(object sender, RoutedEventArgs e)
        {
            this.txtProgress.Text = string.Empty;

            OpenFileDialog oFileDialog = new OpenFileDialog();
            oFileDialog.Filter = "All Files|*.*";
            oFileDialog.FilterIndex = 1;
            oFileDialog.Multiselect = true;

            string data = string.Empty;

            if (oFileDialog.ShowDialog() == true)
            {
                CreateFolder(siteUrl, libName, string.Empty, "Madhur");


                foreach (FileInfo file in oFileDialog.Files)
                {
                    UploadFile(file, "Shared Documents", "Madhur");
                }
            }

        }


        

      
       
    }
}

