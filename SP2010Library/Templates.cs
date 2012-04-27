using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Xml;
using Microsoft.SharePoint;

namespace SP2010Library
{
    public class Templates
    {
        public static void AddTemplateList(SPWeb spWeb, String templateName, String templateFilePath, String listName, String listDescription)
        {
            // Adding Configuration List Template
            SPList templateList = spWeb.Site.RootWeb.Lists["List Template Gallery"];
            // Uploading Configuration List Template to the Root web
            bool f = (from SPFile file in templateList.RootFolder.Files let fileName = Path.GetFileName(templateFilePath) where fileName != null && (file.Name.ToLower() == templateName.ToLower() || fileName.ToLower() == file.Name.ToLower()) select file).Any();
            if (f == false)
            {
                if (File.Exists(templateFilePath))
                {
                    byte[] b = File.ReadAllBytes(templateFilePath);
                    templateList.RootFolder.Files.Add(Path.GetFileName(templateFilePath), b);
                    templateList.Update();
                    spWeb.Update();
                }
            }
            // Creating Configuration List to Web
            SPListTemplate listTemplate = spWeb.Site.GetCustomListTemplates(spWeb.Site.RootWeb)[templateName];
            SPList slist;
            try
            {
                slist = spWeb.Lists[listName];
            }
            catch (Exception)
            {
                slist = null;
            }
            if (slist == null)
            {
                spWeb.Lists.Add(listName, listDescription, listTemplate);
                spWeb.Update();
            }
            int index = 0;
            foreach (SPListItem item in templateList.Items)
            {
                if (listName != null)
                    if (item.DisplayName.ToLower() == listName.ToLower())
                    {
                        templateList.Items.Delete(index);
                        templateList.Update();
                        break;
                    }
                index += 1;
            }
        }

        public static void CopyListTemplateAcrossSites(SPSite sourceSite, String templateName, SPSite destSite, bool overwriteIfExists)
        {
            try
            {
                SPList sourcetemplateList = sourceSite.RootWeb.Lists["List Template Gallery"];
                SPList desttemplateList = sourceSite.RootWeb.Lists["List Template Gallery"];
                SPFile spFile = (from SPFile file in sourcetemplateList.RootFolder.Files let fileName = templateName where fileName != null && (file.Name.ToLower() == templateName.ToLower() || fileName.ToLower() == file.Name.ToLower()) select file).First();
                if (spFile != null)
                {
                    desttemplateList.RootFolder.Files.Add(templateName, spFile.OpenBinary(), overwriteIfExists);
                    desttemplateList.Update();
                }
            }
            catch (Exception)
            {
                return;
            }
        }

        public static void CopyAllListTemplateAcrossSites(SPSite sourceSite, SPSite destSite, bool overwriteIfExists)
        {
            try
            {
                SPList sourcetemplateList = sourceSite.RootWeb.Lists["List Template Gallery"];
                SPList desttemplateList = destSite.RootWeb.Lists["List Template Gallery"];
                foreach (SPListItem spListItem in sourcetemplateList.Items)
                {
                    try
                    {
                        SPFile spFile = spListItem.File;
                        if (spFile != null)
                        {
                            desttemplateList.RootFolder.Files.Add(spListItem.File.Name, spFile.OpenBinary(), overwriteIfExists);
                            desttemplateList.Update();
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
            }
            catch (Exception)
            {
                return;
            }
        }

        public static void CopyAllListTemplateAcrossSites(string remoteSiteUrl, SPSite destSite, bool overwriteIfExists)
        {
            try
            {
                var objLists = new SP2010Library.SPListWebSvc.Lists
                                   {
                                       Credentials = CredentialCache.DefaultCredentials,
                                       Url = remoteSiteUrl + "/_vti_bin/lists.asmx",
                                       AllowAutoRedirect = true
                                   };
                var xmlDoc = new XmlDocument();
                XmlNode ndQuery = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");
                XmlNode ndViewFields = xmlDoc.CreateNode(XmlNodeType.Element, "ViewFields", "");
                XmlNode ndQueryOptions = xmlDoc.CreateNode(XmlNodeType.Element, "QueryOptions", "");
                ndQueryOptions.InnerXml = "<IncludeAttachmentUrls>TRUE</IncludeAttachmentUrls>";
                ndViewFields.InnerXml = "";
                ndQuery.InnerXml = "";
                try
                {
                    XmlNode ndListItems = objLists.GetListItems("List Template Gallery", null, ndQuery, ndViewFields, null, ndQueryOptions, null);
                    XmlNodeList oNodes = ndListItems.ChildNodes;
                    foreach (XmlNode node in oNodes)
                    {
                        var objReader = new XmlNodeReader(node);
                        while (objReader.Read())
                        {
                            if (objReader["ows_EncodedAbsUrl"] != null && objReader["ows_LinkFilename"] != null)
                            {
                                string downloadUrl = objReader["ows_EncodedAbsUrl"];
                                //https://noe.univartest.com/_catalogs/lt/TopEnthusiastsAllItems.stp
                                if (downloadUrl != null)
                                {
                                    var request = (HttpWebRequest)WebRequest.Create(downloadUrl);
                                    request.Credentials = CredentialCache.DefaultCredentials;
                                    request.Timeout = 10000;
                                    request.AllowWriteStreamBuffering = false;
                                    var response = (HttpWebResponse)request.GetResponse();
                                    Stream s = response.GetResponseStream();
                                    SPList desttemplateList = destSite.RootWeb.Lists["List Template Gallery"];
                                    try
                                    {
                                        string tempFilename =
                                            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" +
                                            DateTime.Now.Second + DateTime.Now.Minute + DateTime.Now.Hour + DateTime.Now.Day;
                                        var fs = new FileStream(tempFilename, FileMode.Create);
                                        var read = new byte[256];
                                        int count = s.Read(read, 0, read.Length);
                                        while (count > 0)
                                        {
                                            fs.Write(read, 0, count);
                                            count = s.Read(read, 0, read.Length);
                                        }
                                        //Close everything
                                        fs.Close();
                                        var bytes = File.ReadAllBytes(tempFilename); //new byte[response.ContentLength];
                                        //s.Read(bytes, 0, bytes.Length);
                                        desttemplateList.ParentWeb.AllowUnsafeUpdates = true;
                                        desttemplateList.RootFolder.Files.Add(objReader["ows_LinkFilename"], bytes, overwriteIfExists);
                                        desttemplateList.Update();
                                    }
                                    catch (Exception)
                                    {
                                        continue;
                                    }
                                    s.Close();
                                    response.Close();
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    return;
                }
            }
            catch (Exception)
            {
                return;
            }
        }
    }
}