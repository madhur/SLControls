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
       

       
    }
}

