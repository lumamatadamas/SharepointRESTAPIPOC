using Microsoft.SharePoint.Client;
using System.Security;
using System.Net;
using System;
using System.IO;
using File = System.IO.File;
using System.Text;
using System.Net.Http;
using System.Xml.Linq;

namespace SharepointRESTAPIPOC
{
    class Program
    {
        static void Main(string[] args)
        {
            //urlsite where we want to access organization and projectname
            var urlSite = "https://organization.sharepoint.com/sites/projectname";

            SecureString secureStr = new SecureString();


            foreach (char c in "password".ToCharArray())
            {
                secureStr.AppendChar(c);
            }

            //user credential
            ICredentials credentials = new SharePointOnlineCredentials("user@organization.com", secureStr);

            //DownloadFileViaRestAPI(urlSite, credentials, "Documentos compartidos\\Azure Docs", "Subresource Integrity.docx", "c:\\");


            UploadViaAPI(urlSite, credentials, "Documentos compartidos\\Azure Docs");
        }


        //TODO: parse propertly to retrieve the value for digest
        public static string GetDigest(ICredentials credentials, string url = "")
        {
            using (WebClient client = new WebClient())
            {
                try
                {
                    client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                    client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                    client.Credentials = credentials;

                    string resp = client.UploadString(url, string.Empty);
                    var digestString = XDocument.Parse(resp).Root.Value;
                    return digestString;
                }
                catch (Exception ex)
                {

                    throw;
                }
            }
        }

        public static void UploadViaAPI(string webUrl, ICredentials credentials, string targetFolder)
        {
            var digestFormInfoUrl = string.Format($"{webUrl}/_api/contextinfo", webUrl);

            //here will be the code that get fil's byte[] to send to shared folder in sharepoint
            var file = File.ReadAllBytes("testpdfFile.pdf");
            webUrl = webUrl.EndsWith("/") ? webUrl.Substring(0, webUrl.Length - 1) : webUrl;

            string webRelativeUrl = null;
            if (webUrl.Split('/').Length > 3)
            {
                webRelativeUrl = "/" + webUrl.Split(new char[] { '/' }, 4)[3];
            }
            else
            {
                webRelativeUrl = "";
            }

            string digestValue = GetDigest(credentials, digestFormInfoUrl);
            using (WebClient client = new WebClient())
            {
                try
                {
                    client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                    client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                    //value for X-RequestDigest must be seted to the response of GetDigest function
                    client.Headers.Add("X-RequestDigest", digestValue);

                    client.Credentials = credentials;

                    var uri = webUrl + "/_api/web/GetFolderByServerRelativeUrl('" + targetFolder + "')/Files/add(url='test2.pdf',overwrite=true)";

                    Uri endpoint = new Uri(uri);

                    var response = client.UploadData(endpoint, file);
                }
                catch (Exception ex)
                {

                    throw;
                }
            }


        }

        public static void DownloadFileViaRestAPI(string webUrl, ICredentials credentials, string folder, string fileName, string path)
        {
            webUrl = webUrl.EndsWith("/") ? webUrl.Substring(0, webUrl.Length - 1) : webUrl;

            string webRelativeUrl = null;
            if (webUrl.Split('/').Length > 3)
            {
                webRelativeUrl = "/" + webUrl.Split(new char[] { '/' }, 4)[3];
            }
            else
            {
                webRelativeUrl = "";
            }

            using (WebClient client = new WebClient())
            {
                try
                {
                    client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                    client.Credentials = credentials;
                    var uri = webUrl + "/_api/web/GetFileByServerRelativeUrl('" + webRelativeUrl + "/" + folder + "/" + fileName + "')/$value";

                    Uri endpoint = new Uri(uri);

                    byte[] data = client.DownloadData(endpoint);

                    FileStream outputStream = new FileStream(path + fileName, FileMode.OpenOrCreate | FileMode.Append, FileAccess.Write, FileShare.None);

                    outputStream.Write(data, 0, data.Length);

                    outputStream.Flush(true);

                    outputStream.Close();
                }
                catch (Exception ex)
                {

                    throw;
                }
            }
        }
    }
}



