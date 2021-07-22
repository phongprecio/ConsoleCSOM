using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client.UserProfiles;
using System.Collections.Generic;
using System.Text.RegularExpressions;

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
        private const string TestText = "TestText";
        private const string TestEmail = "TestEmail";
        private const string TestDate = "TestDate";
        private const string TestBool = "TestBool";
        private const string TestInteger = "TestInteger";
        private const string TestPerson = "TestPerson";
        private const string TestTax = "TestTax";
        private const string TestTaxMul = "TestTaxMultiple";

        static async Task Main(string[] args)
        {
            try
            {
                List<string> testProperties = new List<string>() { TestText, TestEmail, TestDate, TestBool, TestInteger, TestPerson, TestTax, TestTaxMul };
                using (var clientContextHelper = new ClientContextHelper())
                {
                    ClientContext ctx = GetContext(clientContextHelper);
                    ctx.Load(ctx.Web);
                    await ctx.ExecuteQueryAsync();

                    Console.WriteLine($"Site {ctx.Web.Title}");

                    // Get the people manager instance and initialize the account name.
                    PeopleManager peopleManager = new PeopleManager(ctx);
                    PersonProperties personProperties = peopleManager.GetMyProperties();
                    ctx.Load(personProperties, p => p.AccountName);
                    ctx.ExecuteQuery();

                    foreach (var property in testProperties)
                    {
                        await UpdatePropertyValueAsync(ctx, peopleManager, personProperties, property);
                    }
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

        private static async Task UpdatePropertyValueAsync(ClientContext context, PeopleManager peopleManager, PersonProperties personProperties, string propertyName)
        {
            string value = String.Empty;

            switch (propertyName)
            {
                case TestText:
                    Console.WriteLine($"Input {propertyName} :");
                    value = Console.ReadLine();
                    // Update property for the user using account name from the user profile.
                    peopleManager.SetSingleValueProfileProperty(personProperties.AccountName, propertyName, value);
                    await context.ExecuteQueryAsync();
                    Console.WriteLine($"{propertyName} updated to {value}");
                    break;
                case TestEmail:
                    bool isMatchRegex = false;
                    string regex = @"\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)\Z";
                    do
                    {
                        Console.WriteLine($"Input {propertyName} :");
                        value = Console.ReadLine();
                        isMatchRegex = Regex.IsMatch(value, regex, RegexOptions.IgnoreCase);

                    } while (!isMatchRegex);

                    peopleManager.SetSingleValueProfileProperty(personProperties.AccountName, propertyName, value);
                    await context.ExecuteQueryAsync();
                    Console.WriteLine($"{propertyName} updated to {value}");
                    break;
                case TestDate:
                    Console.WriteLine($"Input {propertyName} (mm/DD/yyyy):");
                    value = Console.ReadLine();

                    peopleManager.SetSingleValueProfileProperty(personProperties.AccountName, propertyName, value);
                    await context.ExecuteQueryAsync();
                    Console.WriteLine($"{propertyName} updated to {value}");
                    break;
                case TestBool:
                    ConsoleKey response;
                    do
                    {
                        Console.Write("Input bool option [y/n] ");
                        response = Console.ReadKey(false).Key;
                        if (response != ConsoleKey.Enter)
                            Console.WriteLine(); // break new line

                        if (response == ConsoleKey.Y)
                            value = "true";
                        else if (response == ConsoleKey.N)
                            value = "false";

                    } while (response != ConsoleKey.Y && response != ConsoleKey.N);

                    peopleManager.SetSingleValueProfileProperty(personProperties.AccountName, propertyName, value);
                    await context.ExecuteQueryAsync();
                    Console.WriteLine($"{propertyName} updated to {value}");
                    break;
                case TestInteger:
                    GetIntegerValueFromConsole(propertyName, out value);

                    peopleManager.SetSingleValueProfileProperty(personProperties.AccountName, propertyName, value);
                    await context.ExecuteQueryAsync();
                    Console.WriteLine($"{propertyName} updated to {value}");
                    break;
                case TestPerson:
                    Console.WriteLine($"Input {propertyName} email:");
                    value = Console.ReadLine();
                    var user = context.Web.EnsureUser(value);
                    context.Load(user);
                    await context.ExecuteQueryAsync();

                    peopleManager.SetMultiValuedProfileProperty(personProperties.AccountName, propertyName, new List<string> { user.LoginName });
                    await context.ExecuteQueryAsync();
                    Console.WriteLine($"{propertyName} updated to {value}");
                    break;
                case TestTax:
                    Console.WriteLine($"Term values: ");
                    var terms = await GetTermsInTermSetAsync(context);
                    var selectedItem = GetIntegerValueFromConsole(propertyName, out value);

                    peopleManager.SetSingleValueProfileProperty(personProperties.AccountName, propertyName, terms[selectedItem].Name);
                    await context.ExecuteQueryAsync();
                    Console.WriteLine($"{propertyName} updated to {terms[selectedItem].Name}");
                    break;
                case TestTaxMul:
                    Console.WriteLine($"Term values: ");
                    var termSelected = new List<string>();
                    var termList = await GetTermsInTermSetAsync(context);
                    Console.WriteLine("Please input at format: 0,1,2");
                    var selectedString = Console.ReadLine();
                    var selectedItems = selectedString.Split(",").Select(Int32.Parse).ToList();
                    foreach (var item in selectedItems)
                    {
                        termSelected.Add(termList[item].Name);
                    }

                    peopleManager.SetMultiValuedProfileProperty(personProperties.AccountName, propertyName, termSelected);
                    await context.ExecuteQueryAsync();
                    Console.WriteLine($"{propertyName} updated to {selectedString}");
                    break;
            }
        }

        private static async Task<List<Term>> GetTermsInTermSetAsync(ClientContext context)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            // Get the term store default
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get Term Group
            TermGroup termGroup = termStore.Groups.GetByName("Test");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("CITY");
            var terms = termSet.GetAllTerms();
            context.Load(terms);
            await context.ExecuteQueryAsync();
            foreach(var term in terms)
            {
                Console.WriteLine(term.Name);
            }
            return terms.ToList();
        }

        private static int GetIntegerValueFromConsole(string propertyName, out string value)
        {
            bool isInteger = false;
            int numberValue;
            do
            {
                Console.Write($"Input {propertyName} number : ");
                value = Console.ReadLine();
                var isNumeric = int.TryParse(value, out numberValue);
                if (isNumeric && numberValue > 0)
                    isInteger = true;

            } while (!isInteger);

            return numberValue;
        }
    }
}
