using System;
using System.Collections;
using System.IO;
using Microsoft.SharePoint;

namespace SP2010Library
{
    public class DocumentLibrary
    {
        public static void UploadFilesToPictureLibrary(SPWeb spWeb, String libraryName, String filesFolderPath, bool ifAlreadyExistsOverwrite)
        {
            SPPictureLibrary picLib;
            try
            {
                picLib = (SPPictureLibrary)spWeb.Lists[libraryName];
            }
            catch (Exception)
            {
                picLib = null;
            }
            if (picLib == null)
            {
                spWeb.Lists.Add(libraryName, "", SPListTemplateType.PictureLibrary);
                SPList listPage = spWeb.Lists[libraryName];
                //ListPage.OnQuickLaunch = true;
                spWeb.Update();
                listPage.Update();
                try
                {
                    picLib = (SPPictureLibrary)spWeb.Lists[libraryName];
                }
                catch (Exception)
                {
                    picLib = null;
                }
            }
            else
            {
                if (ifAlreadyExistsOverwrite)
                {
                    try
                    {
                        while (picLib.Items.Count > 0)
                        {
                            picLib.Items.Delete(0);
                        }
                        picLib.Update();
                    }
                    catch (Exception)
                    {
                        picLib.Update();
                    }
                }
                else
                    return;
            }
            var props = new Hashtable
                            {
                                {"Title", "Testing"},
                                {"Date Picture Taken", DateTime.Now},
                                {"Description", "Some fancy description"}
                            };
            if (picLib != null)
            {
                string libraryRelativePath = picLib.RootFolder.ServerRelativeUrl;
                string libraryPath = spWeb.Site.MakeFullUrl(libraryRelativePath);
                foreach (String str in Directory.GetFiles(filesFolderPath))
                {
                    using (var fs = new FileStream(str, FileMode.Open))
                    {
                        string imageName = @"\" + Path.GetFileName(str);
                        SPFile file = spWeb.Files.Add(libraryPath + imageName, fs, props, true);
                        int itemId = file.Item.ID;
                        SPListItem item = picLib.GetItemById(itemId);
                        foreach (DictionaryEntry entry in props)
                        {
                            item[entry.Key.ToString()] = entry.Value;
                        }
                        item.UpdateOverwriteVersion();
                    }
                }
            }
        }

        public static void CheckInAllDocuments(string siteUrl, string documentLibraryName)
        {
            using (var site = new SPSite(siteUrl))
            {
                foreach (SPSite spSite in site.WebApplication.Sites)
                {
                    foreach (SPWeb allWeb in spSite.AllWebs)
                    {
                        try
                        {
                            SPList docList = allWeb.Lists[documentLibraryName];
                            foreach (SPListItem item in docList.Items)
                            {
                                if (item.File != null && item.File.CheckedOutByUser != null)
                                {
                                    ImpersonateAndCheckIn(spSite, allWeb.ID, documentLibraryName, item.UniqueId, item.File.CheckedOutByUser.UserToken);
                                }
                            }
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                    }
                }
            }
        }

        private static void ImpersonateAndCheckIn(SPSite spSite, Guid webId, string documentLibraryName, Guid itemId, SPUserToken token)
        {
            using (var site = new SPSite(spSite.Url, token))
            {
                try
                {
                    SPListItem item = site.AllWebs[webId].Lists[documentLibraryName].Items[itemId];
                    item.File.CheckIn("Automated checkin", SPCheckinType.MajorCheckIn);
                    item.Update();
                }
                catch (Exception)
                {
                    return;
                }
            }
        }

        // pending -- need to complete this by creating folder hierarchy --
        public static void CopyDocumentLibrary(SPWeb sourceWeb, SPWeb destinationWeb, string sourceDocumentLibraryName, string destinationDocumentLibraryName)
        {
            SPList sourceDocuments = sourceWeb.Lists[sourceDocumentLibraryName];
            SPList destinationDocuments = destinationWeb.Lists[destinationDocumentLibraryName];

            if (sourceDocuments.ItemCount > 0)
            {
                foreach (SPListItem currentSourceDocument in sourceDocuments.Items)
                {
                    if (currentSourceDocument.FileSystemObjectType == SPFileSystemObjectType.Folder)
                    {
                        string relativeUrl = currentSourceDocument.File.ServerRelativeUrl;
                        destinationDocuments.Items.Add(destinationDocuments.RootFolder.ServerRelativeUrl, SPFileSystemObjectType.Folder, relativeUrl);
                    }
                    else
                    {
                        CreateFolder(destinationDocuments, currentSourceDocument.File.ParentFolder);
                        byte[] fileBytes = currentSourceDocument.File.OpenBinary();
                        const bool overwriteDestinationFile = true;
                        //string relativeDestinationUrl = currentSourceDocument.File.ParentFolder.Name + "/" + currentSourceDocument.File.Name;
                        string relativeDestinationUrl = currentSourceDocument.File.ParentFolder + "/" + currentSourceDocument.File.Name;
                        SPFile destinationFile = ((SPDocumentLibrary)destinationDocuments).RootFolder.Files.Add(relativeDestinationUrl, fileBytes, overwriteDestinationFile);
                    }
                }
            }
        }

        private static void CreateFolder(SPList list, SPFolder spFolder)
        {
            if (spFolder == null || string.IsNullOrEmpty(spFolder.Url))
                return;
            CreateFolder(list, spFolder.ParentFolder);
            string relativeUrl = spFolder.Url;
            list.Items.Add(list.RootFolder.ServerRelativeUrl, SPFileSystemObjectType.Folder, relativeUrl);
            list.Update();
        }

        static public void Copyfiles(SPWeb spWeb, string sourceDocumentLibraryName, string destinationDocumentLibraryName)
        {
            SPList sourceDocuments = spWeb.Lists[sourceDocumentLibraryName];
            SPList destinationDocuments = spWeb.Lists[destinationDocumentLibraryName];
            if (sourceDocuments.ItemCount > 0)
            {
                foreach (SPListItem currentSourceDocument in sourceDocuments.Items)
                {
                    if (currentSourceDocument.FileSystemObjectType == SPFileSystemObjectType.Folder)
                    {
                        string relativeUrl = currentSourceDocument.File.ServerRelativeUrl;
                        destinationDocuments.Items.Add(destinationDocuments.RootFolder.ServerRelativeUrl, SPFileSystemObjectType.Folder, relativeUrl);
                    }
                    else
                    {
                        CreateFolder(destinationDocuments, currentSourceDocument.File.ParentFolder);
                        byte[] fileBytes = currentSourceDocument.File.OpenBinary();
                        const bool overwriteDestinationFile = true;
                        //string relativeDestinationUrl = currentSourceDocument.File.ParentFolder.Name + "/" + currentSourceDocument.File.Name;
                        string relativeDestinationUrl = currentSourceDocument.File.ParentFolder + "/" + currentSourceDocument.File.Name;
                        SPFile destinationFile = ((SPDocumentLibrary)destinationDocuments).RootFolder.Files.Add(relativeDestinationUrl, fileBytes, overwriteDestinationFile);
                    }
                }
            }
            SPFolder oFolder = spWeb.GetFolder(sourceDocumentLibraryName);
            SPFileCollection collFile = oFolder.Files;

            //Copying files from the Root folder.
            CopyFiles(collFile);

            // Get the sub folder collection
            SPFolderCollection collFolder = oFolder.SubFolders;

            EnumerateFolders(collFolder);
        }

        private static void CopyFiles(SPFileCollection collFile)
        {
            //SPFileCollection collFile = oFolder.Files;

            foreach (SPFile oFile in collFile)
            {
                oFile.CopyTo("destlibUrl" + "/" + oFile.Name, true);
            }
        }

        private static void EnumerateFolders(SPFolderCollection copyFolders)
        {
            foreach (SPFolder subFolder in copyFolders)
            {
                if (subFolder.Name != "Forms")
                {
                    SPFileCollection subFiles = subFolder.Files;

                    CopyFiles(subFiles);
                }

                SPFolderCollection subFolders = subFolder.SubFolders;

                EnumerateFolders(subFolders);
            }
        }
    }
}