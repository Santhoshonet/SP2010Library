using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Navigation;

namespace SP2010Library
{
    public class Deployment
    {
        private static bool CreateSpList(SPWeb web, string listName, string description, bool showInQuickLaunch, SPListTemplateType type)
        {
            bool status = false;
            SPSecurity.RunWithElevatedPrivileges(delegate
                                                     {
                                                         using (var site = new SPSite(web.Url))
                                                         {
                                                             SPWeb webLocal = site.AllWebs[web.ID];
                                                             SPList list;
                                                             try
                                                             {
                                                                 list = webLocal.Lists[listName];
                                                                 status = true;
                                                             }
                                                             catch (Exception)
                                                             {
                                                                 list = null;
                                                             }
                                                             if (list == null)
                                                             {
                                                                 webLocal.Lists.Add(listName, description, type);
                                                                 webLocal.Update();
                                                                 try
                                                                 {
                                                                     list = webLocal.Lists[listName];
                                                                     if (showInQuickLaunch)
                                                                     {
                                                                         list.OnQuickLaunch = true;
                                                                         list.Update();
                                                                     }
                                                                     status = true;
                                                                 }
                                                                 catch (Exception)
                                                                 {
                                                                     status = false;
                                                                 }
                                                             }
                                                         }
                                                     });
            return status;
        }

        public static bool CreateGenericList(SPWeb web, string listName, string description, bool showInQuickLaunch)
        {
            return CreateSpList(web, listName, description, showInQuickLaunch, SPListTemplateType.GenericList);
        }

        public static bool CreateDocumentLibrary(SPWeb web, string documentName, string description, bool showInQuickLaunch)
        {
            return CreateSpList(web, documentName, description, showInQuickLaunch, SPListTemplateType.DocumentLibrary);
        }

        public static bool CreateWebPartPage(SPWeb spWeb, string documentName, string pageName, bool showInQuickLaunch)
        {
            bool status = false;
            SPSecurity.RunWithElevatedPrivileges(delegate
            {
                using (var site = new SPSite(spWeb.Url))
                {
                    SPWeb web = site.AllWebs[spWeb.ID];
                    SPList list;
                    try
                    {
                        list = web.Lists[documentName];
                    }
                    catch (Exception)
                    {
                        list = null;
                    }
                    if (list == null)
                    {
                        if (CreateDocumentLibrary(web, documentName, "", true))
                        {
                            web = site.AllWebs[spWeb.ID];
                            try
                            {
                                list = web.Lists[documentName];
                            }
                            catch (Exception)
                            {
                                list = null;
                            }
                        }
                    }
                    if (list != null)
                    {
                        bool isAlreadyFound;
                        try
                        {
                            isAlreadyFound = list.Items.Cast<SPListItem>().Any(listItem => listItem.DisplayName.ToLower() == pageName.ToLower());
                        }
                        catch (Exception)
                        {
                            isAlreadyFound = false;
                        }
                        if (!isAlreadyFound)
                        {
                            string postInformation = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + "<Method>" +
                                                     "<SetList Scope=\"Request\">"
                                                     + list.ID + "</SetList>" + "<SetVar Name=\"ID\">New</SetVar>" +
                                                     "<SetVar Name=\"Cmd\">NewWebPage</SetVar>"
                                                     + "<SetVar Name=\"Type\">WebPartPage</SetVar>" +
                                                     "<SetVar Name=\"WebPartPageTemplate\">1</SetVar>"
                                                     + "<SetVar Name=\"Title\">" + pageName + "</SetVar>" +
                                                     "<SetVar Name=\"Overwrite\">true</SetVar>" + "</Method>";
                            web.ProcessBatchData(postInformation);
                        }
                        status = true;
                    }
                }
            });
            return status;
        }

        public static bool InsertWebPart(SPWeb spWeb, string documentName, bool showInQuickLaunch, string pageName, string webPartName, string webPartProperyName, string userControlLayoutsPath, string zoneId, int zoneIndex, bool removeAllWebParts)
        {
            bool status = false;
            SPSecurity.RunWithElevatedPrivileges(delegate
                                                     {
                                                         using (var site = new SPSite(spWeb.Url))
                                                         {
                                                             SPWeb web = site.AllWebs[spWeb.ID];
                                                             using (System.Web.UI.WebControls.WebParts.WebPart webPart = WebParts.CreateWebPart(web, webPartName))
                                                             {
                                                                 Object obj = webPart.WebBrowsableObject;
                                                                 foreach (PropertyInfo propertyInfo in obj.GetType().GetProperties())
                                                                 {
                                                                     if (propertyInfo.Name == webPartProperyName)
                                                                     {
                                                                         propertyInfo.SetValue(obj, userControlLayoutsPath, null);
                                                                         break;
                                                                     }
                                                                 }
                                                                 SPList list;
                                                                 try
                                                                 {
                                                                     list = web.Lists[documentName];
                                                                 }
                                                                 catch (Exception)
                                                                 {
                                                                     list = null;
                                                                 }
                                                                 if (list == null)
                                                                 {
                                                                     CreateDocumentLibrary(web, documentName, "", showInQuickLaunch);
                                                                     web = site.AllWebs[spWeb.ID];
                                                                     try
                                                                     {
                                                                         list = web.Lists[documentName];
                                                                     }
                                                                     catch (Exception)
                                                                     {
                                                                         list = null;
                                                                     }
                                                                 }
                                                                 if (list != null)
                                                                 {
                                                                     SPListItem item;
                                                                     try
                                                                     {
                                                                         item = list.Items.Cast<SPListItem>().FirstOrDefault(listItem => listItem.DisplayName.ToLower() == pageName.ToLower());
                                                                     }
                                                                     catch (Exception)
                                                                     {
                                                                         item = null;
                                                                     }
                                                                     if (item == null)
                                                                     {
                                                                         CreateWebPartPage(web, documentName, pageName, showInQuickLaunch);
                                                                         web = site.AllWebs[spWeb.ID];
                                                                         list = web.Lists[documentName];
                                                                         foreach (SPListItem listItem in list.Items)
                                                                         {
                                                                             if (listItem.DisplayName.ToLower() == pageName.ToLower())
                                                                             {
                                                                                 item = listItem;
                                                                                 break;
                                                                             }
                                                                         }
                                                                     }
                                                                     if (item != null)
                                                                     {
                                                                         Microsoft.SharePoint.WebPartPages.SPLimitedWebPartManager manager;
                                                                         try
                                                                         {
                                                                             manager = web.GetLimitedWebPartManager(item.Url, System.Web.UI.WebControls.WebParts.PersonalizationScope.Shared);
                                                                         }
                                                                         catch (Exception)
                                                                         {
                                                                             manager = null;
                                                                         }
                                                                         if (manager != null)
                                                                         {
                                                                             Boolean webPartdonotExist = true;
                                                                             if (removeAllWebParts)
                                                                             {
                                                                                 while (true)
                                                                                 {
                                                                                     Boolean exitFlag = true;
                                                                                     for (int i = 0;
                                                                                          i < manager.WebParts.Count;
                                                                                          i++)
                                                                                     {
                                                                                         if (
                                                                                             manager.WebParts[i].Title ==
                                                                                             webPartName)
                                                                                         {
                                                                                             webPartdonotExist = false;
                                                                                         }
                                                                                         else
                                                                                         {
                                                                                             exitFlag = false;
                                                                                             manager.DeleteWebPart(
                                                                                                 manager.WebParts[i]);
                                                                                             break;
                                                                                         }
                                                                                     }
                                                                                     if (exitFlag)
                                                                                         break;
                                                                                 }
                                                                             }
                                                                             if (webPartdonotExist)
                                                                             {
                                                                                 webPart.ChromeType =
                                                                                     System.Web.UI.WebControls.WebParts.
                                                                                         PartChromeType.None;
                                                                                 manager.AddWebPart(webPart, zoneId,
                                                                                                    zoneIndex);
                                                                                 status = true;
                                                                             }
                                                                             manager.Dispose();
                                                                         }
                                                                     }
                                                                 }
                                                             }
                                                         }
                                                     });
            return status;
        }

        public static bool InsertWebPart(SPWeb spWeb, string listName, bool showInQuickLaunch, string webPartName, string webPartProperyName, string userControlLayoutsPath, string zoneId, int zoneIndex, bool removeAllWebParts)
        {
            bool status = false;
            SPSecurity.RunWithElevatedPrivileges(delegate
            {
                using (var site = new SPSite(spWeb.Url))
                {
                    SPWeb web = site.AllWebs[spWeb.ID];
                    using (System.Web.UI.WebControls.WebParts.WebPart webPart = WebParts.CreateWebPart(web, webPartName))
                    {
                        Object obj = webPart.WebBrowsableObject;
                        foreach (PropertyInfo propertyInfo in obj.GetType().GetProperties())
                        {
                            if (propertyInfo.Name == webPartProperyName)
                            {
                                propertyInfo.SetValue(obj, userControlLayoutsPath, null);
                                break;
                            }
                        }
                        SPList list;
                        try
                        {
                            list = web.Lists[listName];
                        }
                        catch (Exception)
                        {
                            list = null;
                        }
                        if (list == null)
                        {
                            CreateGenericList(web, listName, "", showInQuickLaunch);
                            web = site.AllWebs[spWeb.ID];
                            try
                            {
                                list = web.Lists[listName];
                            }
                            catch (Exception)
                            {
                                list = null;
                            }
                        }
                        if (list != null)
                        {
                            Microsoft.SharePoint.WebPartPages.SPLimitedWebPartManager manager;
                            try
                            {
                                manager = web.GetLimitedWebPartManager("Lists/" + list.Title + "/AllItems.aspx", System.Web.UI.WebControls.WebParts.PersonalizationScope.Shared);
                            }
                            catch (Exception)
                            {
                                manager = null;
                            }
                            if (manager != null)
                            {
                                Boolean webPartdonotExist = true;
                                if (removeAllWebParts)
                                {
                                    while (true)
                                    {
                                        Boolean exitFlag = true;
                                        for (int i = 0;
                                             i < manager.WebParts.Count;
                                             i++)
                                        {
                                            if (
                                                manager.WebParts[i].Title ==
                                                webPartName)
                                            {
                                                webPartdonotExist = false;
                                            }
                                            else
                                            {
                                                exitFlag = false;
                                                manager.DeleteWebPart(
                                                    manager.WebParts[i]);
                                                break;
                                            }
                                        }
                                        if (exitFlag)
                                            break;
                                    }
                                }
                                if (webPartdonotExist)
                                {
                                    webPart.ChromeType =
                                        System.Web.UI.WebControls.WebParts.
                                            PartChromeType.None;
                                    manager.AddWebPart(webPart, zoneId,
                                                       zoneIndex);
                                    status = true;
                                }
                                manager.Dispose();
                            }
                        }
                    }
                }
            });
            return status;
        }

        public static bool InsertWebPart(SPWeb spWeb, string documentLibraryName, string webPartName, string webPartProperyName, string userControlLayoutsPath, string zoneId, int zoneIndex, bool removeAllWebParts)
        {
            bool status = false;
            SPSecurity.RunWithElevatedPrivileges(delegate
            {
                using (var site = new SPSite(spWeb.Url))
                {
                    SPWeb web = site.AllWebs[spWeb.ID];
                    using (System.Web.UI.WebControls.WebParts.WebPart webPart = WebParts.CreateWebPart(web, webPartName))
                    {
                        Object obj = webPart.WebBrowsableObject;
                        foreach (PropertyInfo propertyInfo in obj.GetType().GetProperties())
                        {
                            if (propertyInfo.Name == webPartProperyName)
                            {
                                propertyInfo.SetValue(obj, userControlLayoutsPath, null);
                                break;
                            }
                        }
                        SPList list;
                        try
                        {
                            list = web.Lists[documentLibraryName];
                        }
                        catch (Exception)
                        {
                            list = null;
                        }
                        if (list != null)
                        {
                            Microsoft.SharePoint.WebPartPages.SPLimitedWebPartManager manager;
                            try
                            {
                                manager = web.GetLimitedWebPartManager(list.RootFolder.Name + "/Forms/AllItems.aspx", System.Web.UI.WebControls.WebParts.PersonalizationScope.Shared);
                            }
                            catch (Exception)
                            {
                                manager = null;
                            }
                            if (manager != null)
                            {
                                Boolean webPartdonotExist = true;
                                if (removeAllWebParts)
                                {
                                    while (true)
                                    {
                                        Boolean exitFlag = true;
                                        for (int i = 0;
                                             i < manager.WebParts.Count;
                                             i++)
                                        {
                                            if (
                                                manager.WebParts[i].Title ==
                                                webPartName)
                                            {
                                                webPartdonotExist = false;
                                            }
                                            else
                                            {
                                                exitFlag = false;
                                                manager.DeleteWebPart(
                                                    manager.WebParts[i]);
                                                break;
                                            }
                                        }
                                        if (exitFlag)
                                            break;
                                    }
                                }
                                if (webPartdonotExist)
                                {
                                    webPart.ChromeType =
                                        System.Web.UI.WebControls.WebParts.
                                            PartChromeType.None;
                                    manager.AddWebPart(webPart, zoneId,
                                                       zoneIndex);
                                    status = true;
                                }
                                manager.Dispose();
                            }
                        }
                    }
                }
            });
            return status;
        }

        public static bool InsertWebPartAtUrl(SPWeb spWeb, string pageUrl, string webPartName, string webPartProperyName, string userControlLayoutsPath, string zoneId, int zoneIndex, bool removeAllWebParts)
        {
            bool status = false;
            SPSecurity.RunWithElevatedPrivileges(delegate
            {
                using (var site = new SPSite(spWeb.Url))
                {
                    SPWeb web = site.AllWebs[spWeb.ID];
                    using (System.Web.UI.WebControls.WebParts.WebPart webPart = WebParts.CreateWebPart(web, webPartName))
                    {
                        Object obj = webPart.WebBrowsableObject;
                        foreach (PropertyInfo propertyInfo in obj.GetType().GetProperties())
                        {
                            if (propertyInfo.Name == webPartProperyName)
                            {
                                propertyInfo.SetValue(obj, userControlLayoutsPath, null);
                                break;
                            }
                        }

                        Microsoft.SharePoint.WebPartPages.SPLimitedWebPartManager manager;
                        try
                        {
                            manager = web.GetLimitedWebPartManager(pageUrl,
                                                                   System.Web.UI.WebControls.WebParts.
                                                                       PersonalizationScope.Shared);
                        }
                        catch (Exception)
                        {
                            manager = null;
                        }
                        if (manager != null)
                        {
                            Boolean webPartdonotExist = true;
                            if (removeAllWebParts)
                            {
                                while (true)
                                {
                                    Boolean exitFlag = true;
                                    for (int i = 0;
                                         i < manager.WebParts.Count;
                                         i++)
                                    {
                                        if (
                                            manager.WebParts[i].Title ==
                                            webPartName)
                                        {
                                            webPartdonotExist = false;
                                        }
                                        else
                                        {
                                            exitFlag = false;
                                            manager.DeleteWebPart(
                                                manager.WebParts[i]);
                                            break;
                                        }
                                    }
                                    if (exitFlag)
                                        break;
                                }
                            }
                            if (webPartdonotExist)
                            {
                                webPart.ChromeType =
                                    System.Web.UI.WebControls.WebParts.
                                        PartChromeType.None;
                                manager.AddWebPart(webPart, zoneId,
                                                   zoneIndex);
                                status = true;
                            }
                            manager.Dispose();
                        }
                    }
                }
            });
            return status;
        }

        public static void ExecuteProcess(String filePath, ProcessWindowStyle style, bool waitForExit)
        {
            try
            {
                var startInfo = new ProcessStartInfo(filePath) { UseShellExecute = true, WindowStyle = style };
                var batchExecute = new Process { StartInfo = startInfo };
                batchExecute.Start();
                if (waitForExit)
                    batchExecute.WaitForExit();
            }
            catch (Exception)
            {
                return;
            }
        }

        public static void CopyFilesAndFolders(String sourcePath, String destinationPath, bool overWriteFiles)
        {
            if (!Directory.Exists(destinationPath))
            {
                Directory.CreateDirectory(destinationPath);
            }
            string[] fileSet1 = Directory.GetFiles(sourcePath);
            foreach (string s in fileSet1)
            {
                string fileName = Path.GetFileName(s);
                if (fileName != null)
                {
                    string destFile = Path.Combine(destinationPath, fileName);
                    try
                    {
                        File.Copy(s, destFile, overWriteFiles);
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
            }
            foreach (String directoryPath in Directory.GetDirectories(sourcePath))
            {
                string[] dir = directoryPath.Split('\\');
                string dirName = dir[dir.GetUpperBound(0)];
                CopyFilesAndFolders(directoryPath, destinationPath + "\\" + dirName, overWriteFiles);
            }
        }

        public static List<VirtualDirectories> GetVirtualDirectories()
        {
            var iis = new DirectoryEntry("IIS://localhost/W3SVC");
            return (from DirectoryEntry index in iis.Children where index.SchemaClassName == "IIsWebServer" select Convert.ToInt32(index.Name) into id select new DirectoryEntry("IIS://localhost/W3SVC/" + id + "/Root") into rootVDir select rootVDir.Properties["Path"].Value.ToString() into path let str = path.Split('\\') where str.Length > 0 select new VirtualDirectories { Port = Convert.ToInt32(str[str.GetUpperBound(0)]), Path = path }).ToList();
        }

        public static string GetWssVirtualDirectoryPath(string folderPortOrName)
        {
            var iis = new DirectoryEntry("IIS://localhost/W3SVC");
            foreach (DirectoryEntry index in iis.Children)
            {
                if (index.SchemaClassName == "IIsWebServer")
                {
                    int id = Convert.ToInt32(index.Name);
                    var rootVDir = new DirectoryEntry("IIS://localhost/W3SVC/" + id + "/Root");
                    string path = rootVDir.Properties["Path"].Value.ToString();
                    String[] str = path.Split('\\');
                    if (str.Length > 0 && str[str.GetUpperBound(0)].ToLower() == folderPortOrName.ToLower())
                        return path;
                }
            }
            return string.Empty;
        }

        public static string GetVirtualDirectoryPath(SPUrlZone zone, SPSite site)
        {
            try
            {
                return site.WebApplication.IisSettings[zone].Path.FullName;
            }
            catch (Exception)
            { return string.Empty; }
        }

        public static void RestartIis(string applicationStartupPath)
        {
            if (File.Exists(applicationStartupPath + "\\Install31.bat"))
                File.Delete(applicationStartupPath + "\\Install31.bat");
            try
            {
                var writer = new StreamWriter(applicationStartupPath + "\\Install31.bat");
                writer.WriteLine("@ echo off");
                writer.WriteLine("iisreset");
                writer.Flush();
                writer.Close();
                writer.Dispose();
                ExecuteProcess(applicationStartupPath + "\\Install31.bat", ProcessWindowStyle.Hidden, true);
            }
            catch (Exception)
            { ExecuteProcess(applicationStartupPath + "\\Install31.bat", ProcessWindowStyle.Hidden, true); }
        }

        public static void CreateQuickLaunchMenuLink(SPWeb spWeb, string title, string url, bool isExternal, string parentHeading, QuickLaunchPosition position)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate
                                                     {
                                                         using (var site = new SPSite(spWeb.Url))
                                                         {
                                                             SPWeb web = site.AllWebs[spWeb.ID];
                                                             foreach (SPNavigationNode node in web.Navigation.QuickLaunch)
                                                             {
                                                                 if (node.Title == parentHeading)
                                                                 {
                                                                     web.AllowUnsafeUpdates = true;
                                                                     var node1 = new SPNavigationNode(title, url, isExternal);
                                                                     if (position == QuickLaunchPosition.AddAtLast)
                                                                         node.Children.AddAsLast(node1);
                                                                     else
                                                                         node.Children.AddAsFirst(node1);
                                                                     web.Update();
                                                                     web.AllowUnsafeUpdates = false;
                                                                     web.Update();
                                                                     break;
                                                                 }
                                                             }
                                                         }
                                                     });
        }

        public static void CreateQuickLaunchMenuHeading(SPWeb spWeb, string title, string url, bool isExternal, QuickLaunchPosition position)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate
                                                     {
                                                         using (var site = new SPSite(spWeb.Url))
                                                         {
                                                             SPWeb web = site.AllWebs[spWeb.ID];
                                                             web.AllowUnsafeUpdates = true;
                                                             var node1 = new SPNavigationNode(title, url,
                                                                                                           isExternal);
                                                             if (position == QuickLaunchPosition.AddAtLast)
                                                                 web.Navigation.QuickLaunch.AddAsLast(node1);
                                                             else
                                                                 web.Navigation.QuickLaunch.AddAsFirst(node1);
                                                             web.Update();
                                                             web.AllowUnsafeUpdates = false;
                                                             web.Update();
                                                         }
                                                     });
        }

        public static void CreatTopLinkBarHeading(SPWeb spWeb, string title, string url, bool isExternal, TopLinkMenuPosition position, int positionIndex, bool overWrite)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate
            {
                using (var site = new SPSite(spWeb.Url))
                {
                    SPWeb web = site.AllWebs[spWeb.ID];
                    web.AllowUnsafeUpdates = true;
                    SPNavigationNodeCollection navigation = web.Navigation.TopNavigationBar;
                    if (navigation.Count > 0)
                    {
                        if (overWrite)
                        {
                            DeleteLink(navigation, title);
                        }
                    }
                    if (navigation.Count > 0)
                    {
                        if (position == TopLinkMenuPosition.None)
                        {
                            if (navigation[0].Children.Count > positionIndex)
                            {
                                var node = new SPNavigationNode(title, url, isExternal);
                                SPNavigationNode previousNode = navigation[0].Children[positionIndex];
                                navigation[0].Children.Add(node, previousNode);
                            }
                        }
                        else
                        {
                            var node = new SPNavigationNode(title, url, isExternal);
                            if (position == TopLinkMenuPosition.AddAtFist)
                            {
                                navigation[0].Children.AddAsFirst(node);
                            }
                            else
                            {
                                navigation[0].Children.AddAsLast(node);
                            }
                        }
                        web.Update();
                        web.AllowUnsafeUpdates = false;
                        web.Update();
                    }
                }
            });
        }

        public static void DeleteLink(SPNavigationNodeCollection navigation, string title)
        {
            bool isUpdatesMade = true;
            while (isUpdatesMade)
            {
                isUpdatesMade = false;
                foreach (SPNavigationNode node in navigation)
                {
                    if (node.Title == title)
                    {
                        navigation.Delete(node);
                        isUpdatesMade = true;
                    }
                    DeleteLink(node.Children, title);
                }
            }
        }

        public static void CreatTopLinkBarSubMenu(SPWeb spWeb, string headerTitle, string title, string url, bool isExternal, bool overWrite)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate
            {
                using (var site = new SPSite(spWeb.Url))
                {
                    SPWeb web = site.AllWebs[spWeb.ID];
                    web.AllowUnsafeUpdates = true;
                    SPNavigationNodeCollection navigation = web.Navigation.TopNavigationBar;
                    if (navigation.Count > 0)
                    {
                        if (overWrite)
                        {
                            foreach (SPNavigationNode node in navigation)
                            {
                                if (node.Title.ToLower() == headerTitle.ToLower())
                                {
                                    foreach (SPNavigationNode nodechild in node.Children)
                                    {
                                        node.Children.Delete(nodechild);
                                        web.Update();
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    if (navigation.Count > 0)
                    {
                        var node = new SPNavigationNode(title, url, isExternal);
                        foreach (SPNavigationNode tnode in navigation)
                        {
                            if (tnode.Title.ToLower() == headerTitle.ToLower())
                            {
                                tnode.Children.AddAsLast(node);
                                web.Update();
                                web.AllowUnsafeUpdates = false;
                                web.Update();
                                break;
                            }
                        }
                    }
                }
            });
        }
    }

    public class VirtualDirectories
    {
        public int Port;
        public string Path;
    }

    public enum QuickLaunchPosition
    {
        AddAtFist,
        AddAtLast
    }

    public enum TopLinkMenuPosition
    {
        AddAtFist,
        AddAtLast,
        None
    }
}