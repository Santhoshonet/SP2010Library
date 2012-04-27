using System;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;

namespace SP2010Library
{
    public class WebParts
    {
        public static void ValidateAndCreateList(SPWeb rootWeb, String listName, String webPartName, int zoneIndex)
        {
            try
            {
                SPList listToCheck;
                try
                {
                    listToCheck = rootWeb.Lists[listName];
                }
                catch (Exception)
                {
                    listToCheck = null;
                }
                if (listToCheck == null)
                {
                    CreateListAndInsertWebPart(rootWeb, listName, webPartName, zoneIndex);
                }
                else
                {
                    try
                    {
                        if (zoneIndex == 0)
                        {
                            rootWeb.Lists.Delete(listToCheck.ID);
                            rootWeb.Update();
                        }
                        CreateListAndInsertWebPart(rootWeb, listName, webPartName, zoneIndex);
                    }
                    catch (Exception)
                    {
                        return;
                    }
                }
            }
            catch (Exception)
            {
                return;
            }
        }

        public static void AddWebPartToPage(SPWeb web, string pageUrl, string webPartName, string zoneId, int zoneIndex)
        {
            // adding Web part if not exists on the workspace
            using (WebPart webPart = CreateWebPart(web, webPartName))
            {
                Microsoft.SharePoint.WebPartPages.SPLimitedWebPartManager manager;
                try
                {
                    manager = web.GetLimitedWebPartManager(pageUrl, PersonalizationScope.Shared);
                }
                catch (Exception)
                {
                    manager = null;
                }
                if (manager != null)
                {
                    Boolean webPartdonotExist = true;
                    if (zoneIndex == 0)
                    {
                        while (true)
                        {
                            Boolean exitFlag = true;
                            for (int i = 0; i < manager.WebParts.Count; i++)
                            {
                                if (manager.WebParts[i].Title == webPartName)
                                {
                                    webPartdonotExist = false;
                                }
                                else
                                {
                                    exitFlag = false;
                                    manager.DeleteWebPart(manager.WebParts[i]);
                                    break;
                                }
                            }
                            if (exitFlag)
                                break;
                        }
                    }
                    if (webPartdonotExist)
                    {
                        webPart.ChromeType = PartChromeType.None;
                        manager.AddWebPart(webPart, zoneId, zoneIndex);
                    }
                    //  return webPart.ID;
                    manager.Dispose();
                }
            }
        }

        public static void AddWebPartToPage(SPWeb web, string pageUrl, string webPartName, string zoneId, int zoneIndex, bool flag)
        {
            try
            {
                // adding Web part if not exists on the workspace
                using (WebPart webPart = CreateWebPart(web, webPartName))
                {
                    Microsoft.SharePoint.WebPartPages.SPLimitedWebPartManager manager;
                    try
                    {
                        manager = web.GetLimitedWebPartManager(pageUrl, PersonalizationScope.Shared);
                    }
                    catch (Exception)
                    {
                        manager = null;
                    }
                    if (manager != null)
                    {
                        webPart.ChromeType = PartChromeType.None;
                        manager.AddWebPart(webPart, zoneId, zoneIndex);
                        //  return webPart.ID;
                        manager.Dispose();
                    }
                }
            }
            catch (Exception)
            {
                return;
            }
        }

        public static void CreateListAndInsertWebPart(SPWeb rootWeb, String listName, String webPartName, int zoneIndex)
        {
            try
            {
                if (zoneIndex == 0)
                    rootWeb.Lists.Add(listName, "", SPListTemplateType.GenericList);
                SPList listPage = rootWeb.Lists[listName];
                listPage.OnQuickLaunch = true;
                rootWeb.Update();
                listPage.Update();
                AddWebPartToPage(rootWeb, "Lists/" + listName + "/AllItems.aspx", webPartName, "", zoneIndex);
            }
            catch (Exception)
            {
                return;
            }
        }

        public static WebPart CreateWebPart(SPWeb web, string webPartName)
        {
            var query = new SPQuery
            {
                Query =
                    String.Format(
                    "<Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='string'>{0}</Value></Eq></Where>",
                    webPartName)
            };
            SPList webPartGallery = web.ParentWeb == null ? web.GetCatalog(SPListTemplateType.WebPartCatalog) : web.ParentWeb.GetCatalog(SPListTemplateType.WebPartCatalog);
            SPListItemCollection webParts = webPartGallery.GetItems(query);
            if (webParts.Count > 0)
            {
                string typeName = webParts[0].GetFormattedValue("ows_WebPartTypeName");
                string assemblyName = webParts[0].GetFormattedValue("ows_WebPartAssembly");
                if (assemblyName != "")
                {
                    System.Runtime.Remoting.ObjectHandle webPartHandle = Activator.CreateInstance(assemblyName, typeName);
                    {
                        var webPart = (WebPart)webPartHandle.Unwrap();
                        webPart.Title = webPartName;
                        return webPart;
                    }
                }
                return null;
            }
            return null;
        }

        public static String ValidateAndCreatePage(SPWeb rootWeb, String documentName, String documentDesc, String pageName, String webPartName)
        {
            try
            {
                SPList list;
                try
                {
                    list = rootWeb.Lists[documentName];
                }
                catch
                {
                    list = null;
                }
                if (list == null)
                {
                    rootWeb.Lists.Add(documentName, documentDesc, SPListTemplateType.DocumentLibrary);
                    try
                    {
                        list = rootWeb.Lists[documentName];
                    }
                    catch (Exception)
                    {
                        list = null;
                    }
                }
                if (list != null)
                {
                    string postInformation = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + "<Method>" + "<SetList Scope=\"Request\">"
                                                + list.ID + "</SetList>" + "<SetVar Name=\"ID\">New</SetVar>" + "<SetVar Name=\"Cmd\">NewWebPage</SetVar>"
                                                + "<SetVar Name=\"Type\">WebPartPage</SetVar>" + "<SetVar Name=\"WebPartPageTemplate\">1</SetVar>"
                                                + "<SetVar Name=\"Title\">" + pageName + "</SetVar>" + "<SetVar Name=\"Overwrite\">true</SetVar>" + "</Method>";
                    rootWeb.ProcessBatchData(postInformation);
                    try
                    {
                        foreach (SPListItem listItem in list.Items)
                        {
                            if (listItem.DisplayName.ToLower() == pageName.ToLower())
                            {
                                AddWebPartToPage(rootWeb, listItem.Url, webPartName, "", 0);
                                return "Process Completed";
                            }
                        }
                    }
                    catch (Exception)
                    {
                        return "Unable to Create a Page in the Document Library";
                    }
                }
                return "Unable to Create Document Library. Verify the Permission.";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}