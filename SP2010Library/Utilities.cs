using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Xml;
using Microsoft.Office.Server;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace SP2010Library
{
    public class Utilities
    {
        #region MOSS Web Pages Related functions

        public static String GetDefaultZoneUri(SPSite spSite)
        {
            if (spSite != null)
            {
                try
                {
                    SPAlternateUrlCollectionManager urlAccMgr = spSite.WebApplication.Farm.AlternateUrlCollections;
                    var incomingUri = new Uri(spSite.Url);
                    Uri responseUri = urlAccMgr.RebaseUriWithAlternateUri(incomingUri, SPUrlZone.Default);
                    return responseUri.OriginalString;
                }
                catch (Exception)
                {
                    string returnUrl = string.Empty;
                    SPSecurity.RunWithElevatedPrivileges(delegate
                                                             {
                                                                 using (var site = new SPSite(spSite.Url))
                                                                 {
                                                                     SPAlternateUrlCollectionManager urlAccMgr =
                                                                         site.WebApplication.Farm.
                                                                             AlternateUrlCollections;
                                                                     var incomingUri = new Uri(spSite.Url);
                                                                     Uri responseUri =
                                                                         urlAccMgr.RebaseUriWithAlternateUri(
                                                                             incomingUri, SPUrlZone.Default);
                                                                     returnUrl = responseUri.OriginalString;
                                                                 }
                                                             });
                    return returnUrl;
                }
            }
            return string.Empty;
        }

        public static String GetUrlByZone(SPSite spSite, SPUrlZone zone)
        {
            try
            {
                SPAlternateUrlCollectionManager urlAccMgr = spSite.WebApplication.Farm.AlternateUrlCollections;
                Uri responseUri = urlAccMgr.RebaseUriWithAlternateUri(new Uri(spSite.Url), zone);
                return responseUri.OriginalString;
            }
            catch (Exception)
            {
                string returnUrl = string.Empty;
                SPSecurity.RunWithElevatedPrivileges(delegate
                {
                    using (var site = new SPSite(spSite.Url))
                    {
                        SPAlternateUrlCollectionManager urlAccMgr = site.WebApplication.Farm.AlternateUrlCollections;
                        Uri responseUri = urlAccMgr.RebaseUriWithAlternateUri(new Uri(spSite.Url), zone);
                        returnUrl = responseUri.OriginalString;
                    }
                });
                return returnUrl;
            }
        }

        public static String GetDefaultZoneUri(SPSite spSite, String fullUrl)
        {
            try
            {
                SPAlternateUrlCollectionManager urlAccMgr = spSite.WebApplication.Farm.AlternateUrlCollections;
                var incomingUri = new Uri(fullUrl);
                Uri responseUri = urlAccMgr.RebaseUriWithAlternateUri(incomingUri, SPUrlZone.Default);
                return responseUri.OriginalString;
            }
            catch (Exception)
            {
                string returnUrl = string.Empty;
                SPSecurity.RunWithElevatedPrivileges(delegate
                {
                    using (var site = new SPSite(spSite.Url))
                    {
                        SPAlternateUrlCollectionManager urlAccMgr = site.WebApplication.Farm.AlternateUrlCollections;
                        var incomingUri = new Uri(fullUrl);
                        Uri responseUri = urlAccMgr.RebaseUriWithAlternateUri(incomingUri, SPUrlZone.Default);
                        returnUrl = responseUri.OriginalString;
                    }
                });
                return returnUrl;
            }
        }

        public static string GetSspName(SPSite site)
        {
            var serverContext = ServerContext.GetContext(site);
            PropertyInfo srpProp = serverContext.GetType().GetProperty("SharedResourceProvider", BindingFlags.NonPublic | BindingFlags.Instance);
            object sharedResourceProvider = srpProp.GetValue(serverContext, null);
            PropertyInfo srpName = sharedResourceProvider.GetType().GetProperty("Name", BindingFlags.Public | BindingFlags.Instance);
            var sspName = (string)srpName.GetValue(sharedResourceProvider, null);
            return sspName;
        }

        public static string GetSspName(string siteUrl)
        {
            using (var site = new SPSite(siteUrl))
            {
                var serverContext = ServerContext.GetContext(site);
                PropertyInfo srpProp = serverContext.GetType().GetProperty("SharedResourceProvider", BindingFlags.NonPublic | BindingFlags.Instance);
                object sharedResourceProvider = srpProp.GetValue(serverContext, null);
                PropertyInfo srpName = sharedResourceProvider.GetType().GetProperty("Name", BindingFlags.Public | BindingFlags.Instance);
                var sspName = (string)srpName.GetValue(sharedResourceProvider, null);
                return sspName;
            }
        }

        public static SPGroup GetSiteGroup(SPWeb web, string name)
        {
            return web.SiteGroups.Cast<SPGroup>().FirstOrDefault(@group => @group.Name.ToLower() == name.ToLower());
        }

        #endregion MOSS Web Pages Related functions

        #region Web Config Related Functions

        private static bool _isValueExists;
        private static List<bool> _nodeAttributesFound;

        public static void UpdateWebConfigForAjax(String webConfigFilePath)
        {
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/configSections", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<sectionGroup name=""system.web.extensions"" type=""System.Web.Configuration.SystemWebExtensionsSectionGroup,System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31BF3856AD364E35"" />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/configSections/sectionGroup[@name='system.web.extensions']", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<sectionGroup name='scripting' type='System.Web.Configuration.ScriptingSectionGroup,System.Web.Extensions, Version=3.5.0.0, Culture=neutral,PublicKeyToken=31BF3856AD364E35'/>");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/configSections/sectionGroup[@name='system.web.extensions']/sectionGroup[@name='scripting']", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<section name='scriptResourceHandler' type='System.Web.Configuration.ScriptingScriptResourceHandlerSection,System.Web.Extensions, Version=3.5.0.0, Culture=neutral,PublicKeyToken=31BF3856AD364E35' requirePermission='false' allowDefinition='MachineToApplication'/>");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/configSections/sectionGroup[@name='system.web.extensions']/sectionGroup[@name='scripting']", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<sectionGroup name='webServices' type='System.Web.Configuration.ScriptingWebServicesSectionGroup, System.Web.Extensions, Version=3.5.0.0, Culture=neutral,PublicKeyToken=31BF3856AD364E35'/>");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/configSections/sectionGroup[@name='system.web.extensions']/sectionGroup[@name='scripting']/sectionGroup[@name='webServices']", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<section name='jsonSerialization' type='System.Web.Configuration.ScriptingJsonSerializationSection,System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31BF3856AD364E35' requirePermission='false' allowDefinition='Everywhere' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/configSections/sectionGroup[@name='system.web.extensions']/sectionGroup[@name='scripting']/sectionGroup[@name='webServices']", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<section name='profileService' type='System.Web.Configuration.ScriptingProfileServiceSection, System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31BF3856AD364E35' requirePermission='false' allowDefinition='MachineToApplication' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/configSections/sectionGroup[@name='system.web.extensions']/sectionGroup[@name='scripting']/sectionGroup[@name='webServices']", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<section name='authenticationService' type='System.Web.Configuration.ScriptingAuthenticationServiceSection, System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31BF3856AD364E35' requirePermission='false' allowDefinition='MachineToApplication' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/configSections/sectionGroup[@name='system.web.extensions']/sectionGroup[@name='scripting']/sectionGroup[@name='webServices']", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<section name='roleService' type='System.Web.Configuration.ScriptingRoleServiceSection, System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31BF3856AD364E35' requirePermission='false' allowDefinition='MachineToApplication' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web/pages", SPWebConfigModification.SPWebConfigModificationType.EnsureSection, @"<controls/>");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web/pages/controls", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add tagPrefix='asp' namespace='System.Web.UI' assembly='System.Web.Extensions, Version=3.5.0.0, Culture=neutral,PublicKeyToken=31BF3856AD364E35'/>");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web/pages/controls", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add tagPrefix='asp' namespace='System.Web.UI.WebControls' assembly='System.Web.Extensions, Version=3.5.0.0, Culture=neutral,PublicKeyToken=31BF3856AD364E35' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web/compilation/assemblies", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, "<add assembly='System.Core, Version=3.5.0.0, Culture=neutral, PublicKeyToken=B77A5C561934E089' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web/compilation/assemblies", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add assembly='System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31BF3856AD364E35' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web/compilation/assemblies", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add assembly='System.Data.DataSetExtensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=B77A5C561934E089' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web/compilation/assemblies", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add assembly='System.Xml.Linq, Version=3.5.0.0, Culture=neutral, PublicKeyToken=B77A5C561934E089' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web/httpHandlers", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add verb='*' path='*.asmx' validate='false' type='System.Web.Script.Services.ScriptHandlerFactory, System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31BF3856AD364E35' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web/httpHandlers", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add verb='*' path='*_AppService.axd' validate='false' type='System.Web.Script.Services.ScriptHandlerFactory,System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31BF3856AD364E35' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web/httpHandlers", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add verb='GET,HEAD' path='ScriptResource.axd' type='System.Web.Handlers.ScriptResourceHandler,System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31BF3856AD364E35' validate='false' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web/httpModules", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add name='ScriptModule' type='System.Web.Handlers.ScriptModule, System.Web.Extensions, Version=3.5.0.0, Culture=neutral,PublicKeyToken=31BF3856AD364E35' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/SharePoint/SafeControls", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<SafeControl Assembly='System.Web.Silverlight, Version=2.0.5.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35' Namespace='System.Web.UI.SilverlightControls' TypeName='*' Safe='True' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/SharePoint/SafeControls", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<SafeControl Assembly='System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'  Namespace='System.Web.UI' TypeName='*' Safe='True' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<system.web.extensions />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web.extensions", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<scripting />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web.extensions/scripting", SPWebConfigModification.SPWebConfigModificationType.EnsureSection, @"<webServices />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<system.webServer />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.webServer", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<validation validateIntegratedModeConfiguration='false' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.webServer", SPWebConfigModification.SPWebConfigModificationType.EnsureSection, @"<modules />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.webServer/modules", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<remove name='ScriptModule' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.webServer/modules", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add name='ScriptModule' preCondition='managedHandler' type='System.Web.Handlers.ScriptModule, System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31BF3856AD364E35' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.webServer", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<handlers />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.webServer/handlers", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<remove name='WebServiceHandlerFactory-Integrated' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.webServer/handlers", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<remove name='ScriptHandlerFactory' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.webServer/handlers", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<remove name='ScriptHandlerFactoryAppServices' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.webServer/handlers", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<remove name='ScriptResource' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.webServer/handlers", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add name='ScriptHandlerFactory' verb='*' path='*.asmx' preCondition='integratedMode' type='System.Web.Script.Services.ScriptHandlerFactory,System.Web.Extensions, Version=3.5.0.0, Culture=neutral,PublicKeyToken=31BF3856AD364E35' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.webServer/handlers", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add name='ScriptHandlerFactoryAppServices' verb='*' path='*_AppService.axd' preCondition='integratedMode' type='System.Web.Script.Services.ScriptHandlerFactory,System.Web.Extensions, Version=3.5.0.0, Culture=neutral,PublicKeyToken=31BF3856AD364E35' />");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.webServer/handlers", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add name='ScriptResource' preCondition='integratedMode' verb='GET,HEAD' path='ScriptResource.axd' type='System.Web.Handlers.ScriptResourceHandler,System.Web.Extensions, Version=3.5.0.0, Culture=neutral,PublicKeyToken=31BF3856AD364E35' />");
            SetAttributeValueInWebConfig(webConfigFilePath, "configuration/system.web/trust", "level", "Full");
        }

        public static void UpdateWebConfigForTelerikDefaults(String webConfigFilePath)
        {
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web/httpHandlers", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add path='Telerik.Web.UI.WebResource.axd' type='Telerik.Web.UI.WebResource' verb='*' validate='false'/>");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web/httpHandlers", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add path='Telerik.Web.UI.DialogHandler.aspx' type='Telerik.Web.UI.DialogHandler' verb='*' validate='false'/>");
            AddEntryIntoWebConfig(webConfigFilePath, "configuration/system.web/httpHandlers", SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode, @"<add path='Telerik.Web.UI.SpellCheckHandler.axd' type='Telerik.Web.UI.SpellCheckHandler' verb='*' validate='false'/>");
        }

        public static void AddEntryIntoWebConfig(String webConfigFilePath, String path, SPWebConfigModification.SPWebConfigModificationType type, String value)
        {
            XmlTextReader reader = null;
            XmlNode compareNode = null;
            string[] nodePath = null;
            try
            {
                reader = new XmlTextReader(webConfigFilePath);
                nodePath = path.Split('/');
                var notificationDataXmlDocument = new XmlDocument();
                notificationDataXmlDocument.LoadXml(value);
                compareNode = notificationDataXmlDocument.DocumentElement;
                _nodeAttributesFound = new List<bool>();
                if (compareNode != null)
                    if (compareNode.Attributes != null)
#pragma warning disable 168
                        foreach (object t in compareNode.Attributes)
#pragma warning restore 168
                            _nodeAttributesFound.Add(false);
            }
            catch (Exception)
            {
                _isValueExists = false;
            }
            if (nodePath != null)
                if (nodePath.GetUpperBound(0) > -1 && compareNode != null)
                {
                    while (reader.NodeType != XmlNodeType.Element)
                        reader.Read();
                    _isValueExists = false;
                    IsXmlEntryFoundInWebConfig(path, reader.ReadSubtree(), 0, compareNode);
                    try
                    {
                        reader.Close();
                    }
                    catch
                    {
                        reader.Close();
                    }
                    if (_isValueExists == false)
                        AddEntrytoWebConfig(webConfigFilePath, path, value);
                }
        }

        public static void IsXmlEntryFoundInWebConfig(String path, XmlReader rootNode, int index, XmlNode nodeToCompare)
        {
            if (_isValueExists == false)
            {
                string[] nodePath = path.Split('/');
                while (rootNode.Read())
                {
                    switch (rootNode.NodeType)
                    {
                        case XmlNodeType.Element: // The node is an element.
                            string rootNodeName;
                            string attributeNameTag = string.Empty;
                            if (nodePath[index].IndexOf("[@name='", StringComparison.Ordinal) > 0)
                            {
                                rootNodeName = nodePath[index].Substring(0, nodePath[index].IndexOf("[@name='", StringComparison.Ordinal)).Trim();
                                attributeNameTag = nodePath[index].Substring(nodePath[index].IndexOf("[@name='", StringComparison.Ordinal) + 8, nodePath[index].Length - nodePath[index].IndexOf("[@name='") - 10).Trim();
                            }
                            else
                                rootNodeName = nodePath[index];
                            if (rootNode.Name.ToLower() == rootNodeName.ToLower())
                            {
                                var attributes = rootNode.GetAttribute("name");
                                if (attributes != null && (attributeNameTag == string.Empty || attributes.ToLower() == attributeNameTag.ToLower()))
                                {
                                    if (nodePath.GetUpperBound(0) > index)
                                    {
                                        index += 1;
                                        IsXmlEntryFoundInWebConfig(path, rootNode.ReadSubtree(), index, nodeToCompare);
                                    }
                                    else if (nodePath.GetUpperBound(0) == index)
                                    {
                                        XmlReader comparisonReader = rootNode.ReadSubtree();
                                        while (comparisonReader.Read())
                                        {
                                            if (comparisonReader.NodeType == XmlNodeType.Element)
                                            {
                                                if (comparisonReader.Name.ToLower() == nodeToCompare.Name.ToLower())
                                                {
                                                    int i = 0;
                                                    if (nodeToCompare.Attributes != null)
                                                        foreach (XmlAttribute attribute in nodeToCompare.Attributes)
                                                        {
                                                            if (comparisonReader.MoveToAttribute(attribute.Name))
                                                            {
                                                                string str1 = attribute.Value.Trim();
                                                                string str2 = comparisonReader.Value.Trim();
                                                                str1 = str1.Replace(" ", "");
                                                                str2 = str2.Replace(" ", "");
                                                                if (str1.ToLower() == str2.ToLower())
                                                                {
                                                                    _nodeAttributesFound[i] = true;
                                                                }
                                                                else
                                                                    break;
                                                            }
                                                            i += 1;
                                                        }
                                                    bool isItAllTrue = _nodeAttributesFound.All(f => f);
                                                    if (isItAllTrue == false)
                                                    {
                                                        for (i = 0; i < _nodeAttributesFound.Count; i++)
                                                            _nodeAttributesFound[i] = false;
                                                    }
                                                    else
                                                        _isValueExists = true;
                                                }
                                            }
                                        }
                                        try
                                        {
                                            rootNode.ReadEndElement();
                                        }
                                        catch (Exception)
                                        {
                                            rootNode.Read();
                                        }
                                    }
                                    else
                                        return;
                                }
                            }
                            break;
                    }
                }
            }
        }

        public static void AddEntrytoWebConfig(String webConfigFilePath, String path, String value)
        {
            var myXmlDocument = new XmlDocument();
            myXmlDocument.Load(webConfigFilePath);
            XmlNode updateRequiredNode = GotoNodePosition(path, myXmlDocument);
            if (updateRequiredNode != null)
            {
                var notificationDataXmlDocument = new XmlDocument();
                notificationDataXmlDocument.LoadXml(value);
                XmlNode newNode = notificationDataXmlDocument.DocumentElement;
                if (newNode != null)
                {
                    XmlNode newElement = myXmlDocument.CreateNode(XmlNodeType.Element, newNode.Name, "");
                    if (newNode.Attributes != null)
                        foreach (XmlAttribute attrib in newNode.Attributes)
                        {
                            XmlAttribute attribute = myXmlDocument.CreateAttribute(attrib.Name);
                            attribute.Value = newNode.Attributes[attrib.Name].Value;
                            if (newElement.Attributes != null) newElement.Attributes.Append(attribute);
                        }
                    updateRequiredNode.AppendChild(newElement);
                }
            }
            myXmlDocument.Save(webConfigFilePath);
        }

        public static void SetAttributeValueInWebConfig(String webConfigFilePath, String path, String attributeName, String attributeValue)
        {
            var myXmlDocument = new XmlDocument();
            myXmlDocument.Load(webConfigFilePath);
            XmlNode updateRequiredNode = GotoNodePosition(path, myXmlDocument);
            if (updateRequiredNode != null)
            {
                if (updateRequiredNode.Attributes != null)
                    foreach (XmlAttribute attrib in updateRequiredNode.Attributes)
                    {
                        if (attrib.Name.ToLower() == attributeName.ToLower())
                            attrib.Value = attributeValue;
                    }
            }
            myXmlDocument.Save(webConfigFilePath);
        }

        private static XmlNode GotoNodePosition(String path, XmlDocument myXmlDocument)
        {
            int index;
            string[] nodePath = path.Split('/');
            XmlNode rootnode = myXmlDocument.DocumentElement;
            XmlNode targetnode = rootnode;
            XmlNode updateRequiredNode = rootnode;
            for (index = 1; index <= nodePath.GetUpperBound(0); index++)
            {
                string rootNodeName;
                string attributeNameTag = string.Empty;
                if (nodePath[index].IndexOf("[@name='", StringComparison.Ordinal) > 0)
                {
                    rootNodeName = nodePath[index].Substring(0, nodePath[index].IndexOf("[@name='", StringComparison.Ordinal)).Trim();
                    attributeNameTag = nodePath[index].Substring(nodePath[index].IndexOf("[@name='", StringComparison.Ordinal) + 8, nodePath[index].Length - nodePath[index].IndexOf("[@name='", StringComparison.Ordinal) - 10).Trim();
                }
                else
                    rootNodeName = nodePath[index];
                while (true)
                {
                    XmlNode testnode = GetNode(targetnode, rootNodeName, attributeNameTag);
                    if (testnode == null)
                    {
                        //if (NodeIdx < Rootnode.ChildNodes.Count)
                        //    Targetnode = Rootnode.ChildNodes[NodeIdx];
                        //else
                        // return;
                        break;
                    }
                    updateRequiredNode = testnode;
                    targetnode = testnode;
                    break;
                }
            }
            return updateRequiredNode;
        }

        private static XmlNode GetNode(XmlNode rootnode, string rootNodeName, string nodeNameValue)
        {
            foreach (XmlNode childNode in rootnode.ChildNodes)
            {
                if (nodeNameValue != string.Empty)
                {
                    try
                    {
                        if (childNode.Attributes != null && childNode.Attributes["name"].Value.ToLower() == nodeNameValue.ToLower())
                            return childNode;
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
                else if (childNode.Name.ToLower() == rootNodeName.ToLower())
                    return childNode;
            }
            return null;
        }

        #endregion Web Config Related Functions

        #region Email Tempalate Related Functions

        public static object GetMailTemplateData(String filePath, string dataSectionName)
        {
            string beginTag = "<%" + dataSectionName + "%>";
            string endTag = "<%/" + dataSectionName + "%>";
            var strReader = new StreamReader(filePath);
            string emailData = strReader.ReadToEnd();
            var mailTemp = new MailTemplate
                               {
                                   Header = emailData.Substring(0, emailData.IndexOf(beginTag, StringComparison.Ordinal) - beginTag.Length),
                                   Footer = emailData.Substring(emailData.IndexOf(endTag, StringComparison.Ordinal) + endTag.Length)
                               };
            int tempIdx = emailData.IndexOf(beginTag, StringComparison.Ordinal) + beginTag.Length;
            mailTemp.Body = emailData.Substring(tempIdx, emailData.IndexOf(endTag, StringComparison.Ordinal) - tempIdx);
            try
            {
                strReader.Close();
                strReader.Dispose();
            }
            catch (Exception)
            {
                return mailTemp;
            }
            return mailTemp;
        }

        public static MailMessage EmbedTemplateImages(String mailBody)
        {
            var mail = new MailMessage();
            try
            {
                mail.Body = mailBody;
                int index = 0;
                while (mailBody.IndexOf("<%EmbedImage", StringComparison.Ordinal) > 0)
                {
                    String strBeforeImage = mailBody.Substring(0, mailBody.IndexOf("<%EmbedImage", StringComparison.Ordinal));
                    String strEmbedImgeTag = mailBody.Substring(mailBody.IndexOf("<%EmbedImage", StringComparison.Ordinal), mailBody.IndexOf("%>", StringComparison.Ordinal) - strBeforeImage.Length + 2);
                    String strAfterImage = mailBody.Substring(mailBody.IndexOf(strEmbedImgeTag, StringComparison.Ordinal) + strEmbedImgeTag.Length);
                    String imageTag;
                    String imagePath = strEmbedImgeTag.Substring(strEmbedImgeTag.IndexOf("Path=", StringComparison.Ordinal) + 5);
                    imagePath = imagePath.Substring(0, imagePath.Length - 2);
                    if (File.Exists(imagePath))
                    {
                        try
                        {
                            var mailLogo = new Attachment(imagePath) { ContentId = ("EmbeddedImage" + index) };
                            mail.Attachments.Add(mailLogo);
                            imageTag = "<img src='cid:" + "EmbeddedImage" + index + "' border='0'>";
                        }
                        catch (Exception)
                        {
                            imageTag = "";
                        }
                    }
                    else
                        imageTag = "";
                    mailBody = strBeforeImage + imageTag + strAfterImage;
                    // Limitation upto 10 images
                    index += 1;
                    if (index > 10)
                        break;
                }
                mail.Body = mailBody;
                return mail;
            }
            catch (Exception)
            {
                return mail;
            }
        }

        public static void UpdateMailDataPart(Object genericListEmailDataPart, MailTemplate mailTemplate)
        {
            var dataPart = (List<EmailDataPart>)genericListEmailDataPart;
            int i = 0;
            foreach (EmailDataPart dp in dataPart)
            {
                i += 1;
                mailTemplate.Body = mailTemplate.Body.Replace("<%Header" + i + "%>", dp.HeaderData);
                mailTemplate.Body = mailTemplate.Body.Replace("<%Value" + i + "%>", dp.Value);
            }
            while (i <= 20)
            {
                i += 1;
                mailTemplate.Body = mailTemplate.Body.Replace("<%Display" + i + "%>", "none");
                mailTemplate.Body = mailTemplate.Body.Replace("<%Header" + i + "%>", "");
                mailTemplate.Body = mailTemplate.Body.Replace("<%Value" + i + "%>", "");
            }
        }

        public static MailTemplate UpdateMailDataPart(string headerData, string value, int headerNumber, MailTemplate mailTemplate)
        {
            mailTemplate.Body = mailTemplate.Body.Replace("<%Display" + headerNumber + "%>", "");
            mailTemplate.Body = mailTemplate.Body.Replace("<%Header" + headerNumber + "%>", headerData);
            mailTemplate.Body = mailTemplate.Body.Replace("<%Value" + headerNumber + "%>", value);
            return mailTemplate;
        }

        public static MailTemplate CloseMailDataPart(int headerNumber, MailTemplate mailTemplate)
        {
            while (headerNumber <= 20)
            {
                headerNumber += 1;
                mailTemplate.Body = mailTemplate.Body.Replace("<%Display" + headerNumber + "%>", "none");
                mailTemplate.Body = mailTemplate.Body.Replace("<%Header" + headerNumber + "%>", "");
                mailTemplate.Body = mailTemplate.Body.Replace("<%Value" + headerNumber + "%>", "");
            }
            return mailTemplate;
        }

        public static EmailDataPart GetDataPart(string headerData, string value)
        {
            var tempEmailDataPart = new EmailDataPart { HeaderData = headerData, Value = value };
            return tempEmailDataPart;
        }

        #endregion Email Tempalate Related Functions

        #region SecurityEncryption Related functions

        public static string Base64Decode(string sData)
        {
            var encoder = new System.Text.UTF8Encoding();
            System.Text.Decoder utf8Decode = encoder.GetDecoder();
            byte[] todecodeByte = Convert.FromBase64String(sData);
            int charCount = utf8Decode.GetCharCount(todecodeByte, 0, todecodeByte.Length);
            var decodedChar = new char[charCount];
            utf8Decode.GetChars(todecodeByte, 0, todecodeByte.Length, decodedChar, 0);
            var result = new String(decodedChar); return result;
        }

        public static string Base64Encode(string sData)
        {
            try
            {
                byte[] encDataByte = System.Text.Encoding.UTF8.GetBytes(sData);
                string encodedData = Convert.ToBase64String(encDataByte);
                return encodedData;
            }
            catch (Exception ex)
            {
                throw new Exception("Error in base64Encode" + ex.Message);
            }
        }

        #endregion SecurityEncryption Related functions

        #region Properties

        public static string ProjWssuidProperty { get { return "MSPWAPROJUID"; } }

        #endregion Properties
    }
}