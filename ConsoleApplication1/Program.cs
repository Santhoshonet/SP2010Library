using Microsoft.SharePoint;

namespace ConsoleApplication1
{
    internal class Program
    {
        private static void Main()
        {
            string noeSiteUrl = "http://wdv-sea-sptdev1:777";
            SPSecurity.RunWithElevatedPrivileges(delegate
                                                     {
                                                         using (var site = new SPSite(noeSiteUrl))
                                                         {
                                                             ILinkProjectsLibrary.Templates.CopyAllListTemplateAcrossSites("https://universe.univartest.com", site, true);
                                                         }
                                                         noeSiteUrl = "http://wdv-sea-sptdev1:777/process";
                                                         using (var site = new SPSite(noeSiteUrl))
                                                         {
                                                             ILinkProjectsLibrary.Templates.CopyAllListTemplateAcrossSites("https://universe.univartest.com/process", site, true);
                                                         }
                                                         noeSiteUrl = "http://wdv-sea-sptdev1:777/product";
                                                         using (var site = new SPSite(noeSiteUrl))
                                                         {
                                                             ILinkProjectsLibrary.Templates.CopyAllListTemplateAcrossSites("https://universe.univartest.com/product", site, true);
                                                         }
                                                         noeSiteUrl = "http://wdv-sea-sptdev1:777/practice";
                                                         using (var site = new SPSite(noeSiteUrl))
                                                         {
                                                             ILinkProjectsLibrary.Templates.CopyAllListTemplateAcrossSites("https://universe.univartest.com/procece", site, true);
                                                         }
                                                     });
        }
    }
}