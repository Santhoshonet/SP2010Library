using System;
using System.Linq;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace SP2010Library
{
    public class Feature
    {
        public static void ActivateFeatureInAllSites(string siteCollectionUrl, string featureName)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate
            {
                using (var site = new SPSite(siteCollectionUrl))
                {
                    SPWebApplication webApp = site.WebApplication;
                    foreach (SPSite subsite in webApp.Sites)
                    {
                        try
                        {
                            // Activate Site Collection Features
                            SPFeatureDefinitionCollection featureDefinition = SPFarm.Local.FeatureDefinitions;

                            SPFeatureDefinition feature = (from SPFeatureDefinition f
                                          in featureDefinition
                                                           where f.Scope == SPFeatureScope.Site && f.GetTitle(System.Globalization.CultureInfo.CurrentCulture) == featureName
                                                           select f).SingleOrDefault();
                            if (feature != null)
                            {
                                subsite.Features.Add(feature.Id, false);
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Exception occured: {0}\r\n{1}", e.Message, e.StackTrace);
                        }
                    }
                }
            });
        }

        public static void ActivateFeatureInAllWebs(string siteCollectionUrl, string featureName)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate
            {
                using (var site = new SPSite(siteCollectionUrl))
                {
                    SPWebApplication webApp = site.WebApplication;
                    foreach (SPSite subsite in webApp.Sites)
                    {
                        try
                        {
                            // Activate Site Collection Features
                            SPFeatureDefinitionCollection featureDefinition = SPFarm.Local.FeatureDefinitions;

                            SPFeatureDefinition feature = (from SPFeatureDefinition f
                                          in featureDefinition
                                                           where f.Scope == SPFeatureScope.Web && f.GetTitle(System.Globalization.CultureInfo.CurrentCulture) == featureName
                                                           select f).SingleOrDefault();
                            if (feature != null)
                            {
                                foreach (SPWeb spWeb in subsite.AllWebs)
                                {
                                    try
                                    {
                                        spWeb.Features.Add(feature.Id, false);
                                    }
                                    catch (Exception)
                                    {
                                        continue;
                                    }
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Exception occured: {0}\r\n{1}", e.Message, e.StackTrace);
                        }
                    }
                }
            });
        }

        public static void ActivateFeatureInAllWebApplications(string featureName)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate
            {
                try
                {
                    // Activate Site Collection Features
                    SPFeatureDefinitionCollection featureDefinition = SPFarm.Local.FeatureDefinitions;

                    SPFeatureDefinition feature = (from SPFeatureDefinition f
                                  in featureDefinition
                                                   where f.Scope == SPFeatureScope.WebApplication && f.GetTitle(System.Globalization.CultureInfo.CurrentCulture) == featureName
                                                   select f).SingleOrDefault();
                    if (feature != null)
                    {
                        SPFarm farm = SPFarm.Local;
                        var service = farm.Services.GetValue<SPWebService>("");
                        foreach (SPWebApplication webApp in service.WebApplications)
                        {
                            webApp.Features.Add(feature.Id, false);
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception occured: {0}\r\n{1}", e.Message, e.StackTrace);
                }
            });
        }
    }
}