using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;

namespace ConsoleCSOM
{
    class SharepointInfo
    {
        public string SiteUrl { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                using (var clientContextHelper = new ClientContextHelper())
                {
                    ClientContext ctx = GetContext(clientContextHelper);
                    ctx.Load(ctx.Web);
                    await ctx.ExecuteQueryAsync();

                    string ownerEmail = "phong@adminvn.onmicrosoft.com";
                    string user5Email = "user5@adminvn.onmicrosoft.com";

                    Console.WriteLine($"Site {ctx.Web.Title}");

                    await GetDefaultSecurityGroupAsync(ctx);
                    await CreateNewPermissionLevelAsync(ctx);
                    await CreateNewGroupAsync(ctx, ownerEmail, user5Email);
                    await CheckSubSiteInheritedGroupAsync(ctx);
                }

                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
            }
        }

        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }
        
        private static async Task GetDefaultSecurityGroupAsync(ClientContext context)
        {
            var web = context.Web;
            context.Load(web, w => w.AssociatedMemberGroup, w => w.Title);
            await context.ExecuteQueryAsync();

            Console.WriteLine("Associated groups for Site \"" + web.Title + "\"");

            Console.WriteLine("*************************************************");

            Console.WriteLine("Member Group: " + web.AssociatedMemberGroup.Title);
        }

        private static async Task CreateNewPermissionLevelAsync(ClientContext context)
        {
            BasePermissions perm = new BasePermissions();
            perm.Set(PermissionKind.CreateAlerts);
            perm.Set(PermissionKind.ManageLists);
            perm.Set(PermissionKind.ViewListItems);
            perm.Set(PermissionKind.ViewPages);
            perm.Set(PermissionKind.Open);

            RoleDefinitionCreationInformation creationInfo = new RoleDefinitionCreationInformation();
            creationInfo.BasePermissions = perm;
            creationInfo.Description = "A role with create alerts and manage list permission";
            creationInfo.Name = "Alert Manager Role";
            creationInfo.Order = 0;
            context.Web.RoleDefinitions.Add(creationInfo);

            await context.ExecuteQueryAsync();
        }

        private static async Task CreateNewGroupAsync(ClientContext context, string ownerEmail, string userEmail)
        {
            var alertRole = context.Web.RoleDefinitions.GetByName("Alert Manager Role");
            var owner = context.Web.EnsureUser(ownerEmail);
            var user = context.Web.EnsureUser(userEmail);

            var group = new GroupCreationInformation();
            group.Title = "Test Group CSOM";
            var newGroup = context.Web.SiteGroups.Add(group);

            context.Web.RoleAssignments.Add(newGroup, new RoleDefinitionBindingCollection(context) { alertRole });
            newGroup.Owner = owner;
            newGroup.Users.AddUser(user);
            newGroup.Update();
            await context.ExecuteQueryAsync();
        }

        private static async Task CheckSubSiteInheritedGroupAsync(ClientContext context)
        {
            var subSite = context.Site.OpenWeb("FA/");
            var groupInherited = subSite.SiteGroups.GetByName("Test Group CSOM");
            context.Load(groupInherited, g => g.Title);
            await context.ExecuteQueryAsync();
            Console.WriteLine(groupInherited.Title);
        }
    }
}
