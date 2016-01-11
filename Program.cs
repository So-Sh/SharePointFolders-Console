using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Json;

namespace Console_SharePoint_Folders
{
    class Program
    {
        static string sharepointURL = "Your SharePoint Online URL";
        static string documentLibraryName = "Documnents";

        static string sharepointUsername = "User Name";
        static string sharepointPassword = "Password";
        static Uri spSite;
        private static SpoAuthUtility _spo;
        static void Main(string[] args)
        {

            try
            {

                //Initialize the SPoAuthUtility
                spSite = new Uri(sharepointURL);
                _spo = SpoAuthUtility.Create(spSite, sharepointUsername, WebUtility.HtmlEncode(sharepointPassword), false);

                if (_spo != null)
                {
                    string foldername = "";

                    string digest = _spo.GetRequestDigest();


                    Console.WriteLine("Create Folder: 1.");
                    Console.WriteLine("Rename Folder: 2.");
                    Console.WriteLine("Find Folder: 3. ");
                    Console.WriteLine("Exit: any key");

                    string choice = Console.ReadLine();

                    if (choice == "1")
                    {
                        Console.WriteLine("Please provide a name for the folder and press enter!");
                        foldername = Console.ReadLine();


                        CreateFolder(foldername);

                    }
                    else if (choice == "2")
                    {
                        Console.WriteLine("Please enter the old folder name and hit enter!");
                        string oldfolder = Console.ReadLine();

                        Console.WriteLine("Please enter the new folder name and hit enter!");
                        string newfolder = Console.ReadLine();

                        RenameFolder(oldfolder, newfolder);

                    }
                    else if (choice == "3")
                    {
                        Console.WriteLine("Please enter the  folder name and hit enter!");
                        string foldertofind = Console.ReadLine();


                        Console.WriteLine(FolderExists(foldertofind).ToString());

                    }

                    else
                    {
                        return;


                    }


                }
                else
                {
                    Console.WriteLine("Could not authenticate user! Press any key to exit the application");

                }

                Console.ReadLine();

            }
            catch (Exception ex)
            {

                Console.WriteLine("Error: " + ex.Message);
            }



        } // End of main function


        /// <summary>
        /// Creates a new folder in the Document Library specified in the
        /// documentLibraryName member
        /// </summary>
        /// <param name="foldername">Name of the folder to be created</param>
        public static void CreateFolder(string foldername)
        {



            Uri spSite = new Uri(sharepointURL);


            string odataQuery = "_api/web/getfolderbyserverrelativeurl('" + documentLibraryName + "/" + foldername + "')/folders";

            byte[] content = ASCIIEncoding.ASCII.GetBytes(@"{ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': '" + documentLibraryName + "/" + foldername + "'}");


            string digest = _spo.GetRequestDigest();

            Uri url = new Uri(String.Format("{0}/{1}", _spo.SiteUrl, odataQuery));
            // Set X-RequestDigest
            var webRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            webRequest.Headers.Add("X-RequestDigest", digest);

            // Send a json odata request to SPO rest services to fetch all list items for the list.
            byte[] result = HttpHelper.SendODataJsonRequest(
              url,
              "POST", // sending data to SP through the rest api usually uses the POST verb 
              content,
              webRequest,
              _spo // pass in the helper object that allows us to make authenticated calls to SPO rest services
              );

            string response = Encoding.UTF8.GetString(result, 0, result.Length);


            Console.WriteLine(response);

        }

        /// <summary>
        /// Renameas an exisiting folder in the Document Library specified in the 
        /// documentLibraryName member
        /// </summary>
        /// <param name="oldName">Name of the existing folder</param>
        /// <param name="newName">New name of the folder to be set</param>
        public static void RenameFolder(string oldName, string newName)
        {

            Uri spSite = new Uri(sharepointURL);

            string odataQuery = "_api/web/GetFolderByServerRelativeUrl('" + documentLibraryName + "/" + oldName + "')/ListItemAllFields";

            byte[] content = ASCIIEncoding.ASCII.GetBytes(@"{ '__metadata':{ 'type': 'SP.Data.AccountItem' }, 'FileLeafRef': '" + newName + "' ,'Title:': '" + newName + "' }");


            string digest = _spo.GetRequestDigest();

            Uri url = new Uri(String.Format("{0}/{1}", _spo.SiteUrl, odataQuery));

            // Set X-RequestDigest
            var webRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            webRequest.Headers.Add("X-RequestDigest", digest);

            //Set additional Headers
            webRequest.Headers.Add("IF-MATCH", "*");
            webRequest.Headers.Add("X-HTTP-Method", "PATCH");


            // Send a json odata request to SPO rest services to fetch all list items for the list.
            byte[] result = HttpHelper.SendODataJsonRequest(
              url,
              "POST", // sending data to SP through the rest api usually uses the POST verb 
              content,
              webRequest,
              _spo // pass in the helper object that allows us to make authenticated calls to SPO rest services
              );

            string response = Encoding.UTF8.GetString(result, 0, result.Length);


            Console.WriteLine(response);

        }


        /// <summary>
        /// Checks if a folder exists in the Document Library specified 
        /// in the documentLibraryName member
        /// </summary>
        /// <param name="folderUrl">The relative url of the folder</param>
        /// <returns></returns>
        static bool FolderExists(string folderUrl)
        {


            string restQuery = "_api/web/GetFolderByServerRelativeUrl('" + documentLibraryName + "/" + folderUrl + "')";

            try
            {

                Uri url = new Uri(String.Format("{0}/{1}", _spo.SiteUrl, restQuery));

                // Send a json odata request to SPO rest services to fetch all list items for the list.
                byte[] result = HttpHelper.SendODataJsonRequest(
                  url,
                  "GET", // reading data from SP through the rest api usually uses the GET verb 
                  null,
                  (HttpWebRequest)HttpWebRequest.Create(url),
                  _spo // pass in the helper object that allows us to make authenticated calls to SPO rest services
                  );

                string response = Encoding.UTF8.GetString(result, 0, result.Length);
                Console.WriteLine(response);

                dynamic parsedvalue = JsonValue.Parse(response);

                Console.WriteLine("Folder URL: " + parsedvalue.d.__metadata.id.ToString());



                if (response.Contains("\"Exists\":true,"))
                    return true;
                else
                    return false;

            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return false;
            }


        }

    }
}
