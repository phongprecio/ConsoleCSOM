using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.Generic;

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

                    Console.WriteLine($"Site {ctx.Web.Title}");

                    // await SimpleCamlQueryAsync(ctx);
                    //await CsomTermSetAsync(ctx);
                    //await CreateListCsomTestAsync(ctx);
                    //await CreateTermSetAsync(ctx);
                    //await CreateFieldsAsync(ctx);
                    //await CreateContentTypeAsync(ctx);
                    //await UpdateAboutDefaultAsync(ctx);
                    //await AddTermToTermSetAsync(ctx);
                    //await UpdateCityDefaultAsync(ctx);
                    //await AddSampleDataAsync(ctx);
                    //await DeleteContentType(ctx);
                    //await CamlQueryGetListAboutAsync(ctx);
                    //await CreateViewByCityAsync(ctx);
                    //await UpdateBatchDataAsync(ctx);
                    //await CreateFolderAsync(ctx);
                    //await AddFieldAuthorAsync(ctx);
                    //await CreateTaxonomyMultipleValueAndSetContentTypeAsync(ctx);
                    await CreateDataForTaxonomiFieldMultipleAsync(ctx);
                    //await CreateDocumentListSync(ctx);
                    //await CreateFolderInDocumentListAsync(ctx);
                    //await CreateDocumentInFolderAsync(ctx);
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

        // Create List CSOM Test
        private static async Task CreateListCsomTestAsync(ClientContext context)
        {
            Web web = context.Web;
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.TemplateType = (int)ListTemplateType.GenericList;
            creationInfo.Title = "CSOM Test";
            web.Lists.Add(creationInfo);

            await context.ExecuteQueryAsync();
        }

        private static async Task CreateTermSetAsync(ClientContext context)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            // Get the term store default
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get Term Group
            TermGroup termGroup = termStore.Groups.GetByName("Site Collection - adminvn.sharepoint.com-sites-TrainingSharePoint");
            // Create new term set
            termGroup.CreateTermSet("city-phong", new Guid("ef96ce9b-27fc-44dc-b3f2-6d73e6b3c02e"), 1033);

            await context.ExecuteQueryAsync();
        }

        private static async Task CreateFieldsAsync(ClientContext context)
        {
            Web web = context.Web;

            string fieldAbout = @"<Field Name='About' DisplayName='About' Type='Text' Group='Custom Columns' />";
            web.Fields.AddFieldAsXml(fieldAbout, true, AddFieldOptions.DefaultValue);

            // TODO duplicate name City, so create City2 instead
            //string fieldCity = @"<Field Name='City' Type='TaxonomyFieldType' Group='Base Columns' />";
            //web.Fields.AddFieldAsXml(fieldCity, true, AddFieldOptions.DefaultValue);

            string fieldSchema = "<Field Type='TaxonomyFieldType' DisplayName='City2' Name='City2' Hidden='False'/>";
            web.Fields.AddFieldAsXml(fieldSchema, true, AddFieldOptions.AddToDefaultContentType);

            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName("Site Collection - adminvn.sharepoint.com-sites-TrainingSharePoint");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("city-phong");

            Field field = context.Web.Fields.GetByInternalNameOrTitle("City2");

            context.Load(termStore, t => t.Id);
            context.Load(termSet, t => t.Id);
            context.Load(field);
            await context.ExecuteQueryAsync();

            TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(field);
            taxonomyField.SspId = termStore.Id;
            taxonomyField.TermSetId = termSet.Id;
            taxonomyField.TargetTemplate = String.Empty;
            taxonomyField.Update();

            await context.ExecuteQueryAsync();
        }

        private static async Task CreateContentTypeAsync(ClientContext context)
        {
            Web web = context.Web;
            // Create a Content Type Information object.
            ContentTypeCreationInformation newCt = new ContentTypeCreationInformation();
            // Set the name for the content type.
            newCt.Name = "CSOM Test content type";
            newCt.Id = "0x0100DA5D705795149F44B45648D8D24EAC58";
            // Set content type to be available from specific group.
            newCt.Group = "List Content Types";
            // Create the content type.
            web.ContentTypes.Add(newCt);
            await context.ExecuteQueryAsync();

            // Add content type to existed List
            ContentType contentType = web.ContentTypes.GetById("0x0100DA5D705795149F44B45648D8D24EAC58");

            // Get Site field and add to content type
            FieldCollection fields = web.Fields;
            Field aboutField = fields.GetByInternalNameOrTitle("About");
            Field cityField = fields.GetByInternalNameOrTitle("City2");
            context.Load(cityField);
            context.Load(aboutField);
            await context.ExecuteQueryAsync();

            FieldLinkCreationInformation fldLinkAbout = new FieldLinkCreationInformation();
            fldLinkAbout.Field = aboutField;
            FieldLinkCreationInformation fldLinkCity = new FieldLinkCreationInformation();
            fldLinkCity.Field = cityField;

            contentType.FieldLinks.Add(fldLinkAbout);
            contentType.FieldLinks.Add(fldLinkCity);
            contentType.Update(true);

            // Get Existing List and update
            List csomList = web.Lists.GetByTitle("CSOM Test");
            csomList.ContentTypes.AddExistingContentType(contentType);
            csomList.Update();
            context.Web.Update();
            await context.ExecuteQueryAsync();
        }

        private static async Task AddSampleDataAsync(ClientContext context)
        {
            Web web = context.Web;

            List list = web.Lists.GetByTitle("CSOM Test");
            context.Load(list);
            await context.ExecuteQueryAsync();
            TaxonomyField taxField = context.CastTo<TaxonomyField>(context.Web.Fields.GetByTitle("City2")); ;
            var taxFieldValue = new TaxonomyFieldValue()
            {
                WssId = -1, // alway let it -1
                Label = "Ho Chi Minh",
                TermGuid = "44649e15-7612-432e-b742-147eee391f9b"
            };

            for (var i = 0; i < 5; i++)
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = list.AddItem(itemCreateInfo);
                oListItem["Title"] = "Item" + i;
                oListItem["About"] = "About" + i;
                oListItem["ContentTypeId"] = "0x0100DA5D705795149F44B45648D8D24EAC58"; // Set content type Id

                taxField.SetFieldValueByValue(oListItem, taxFieldValue);
                oListItem.Update();
            }

            await context.ExecuteQueryAsync();
        }

        private static async Task UpdateAboutDefaultAsync(ClientContext context)
        {
            Web web = context.Web;

            List list = web.Lists.GetByTitle("CSOM Test");

            var field = web.Fields.GetByTitle("About");
            field.DefaultValue = "About default";
            field.UpdateAndPushChanges(true);
            await context.ExecuteQueryAsync();

            // Add new data
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem oListItem = list.AddItem(itemCreateInfo);
            oListItem["Title"] = "New Item 2.6";
            oListItem["ContentTypeId"] = "0x0100DA5D705795149F44B45648D8D24EAC58"; // Set content type Id
            oListItem.Update();

            ListItemCreationInformation itemCreateInfo2 = new ListItemCreationInformation();
            ListItem oListItem2 = list.AddItem(itemCreateInfo2);
            oListItem2["Title"] = "New Item 2.7";
            oListItem2["ContentTypeId"] = "0x0100DA5D705795149F44B45648D8D24EAC58"; // Set content type Id
            oListItem2.Update();
            await context.ExecuteQueryAsync();
        }

        private static async Task DeleteContentType(ClientContext context)
        {
            ContentType contentType = context.Web.ContentTypes.GetById("0x0100DA5D705795149F44B45648D8D24EAE20");
            contentType.DeleteObject();
            await context.ExecuteQueryAsync();
        }

        private static async Task AddTermToTermSetAsync(ClientContext context)
        {
            Web web = context.Web;
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            // Get the term store default
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get TermGroup
            TermGroup termGroup = termStore.Groups.GetByName("Site Collection - adminvn.sharepoint.com-sites-TrainingSharePoint");
            // Get TermSet
            TermSet termSet = termGroup.TermSets.GetByName("city-phong");
            termSet.CreateTerm("Ho Chi Minh", 1033, new Guid("44649e15-7612-432e-b742-147eee391f9b"));
            termSet.CreateTerm("Stockholm", 1033, new Guid("1ce8d481-fec5-45ff-8b58-6c96996f6aa7"));
            await context.ExecuteQueryAsync();
        }

        private static async Task UpdateCityDefaultAsync(ClientContext context)
        {
            var taxfield = context.CastTo<TaxonomyField>(context.Web.Fields.GetByTitle("City2"));
            context.Load(taxfield);
            await context.ExecuteQueryAsync();

            var defaultValue = new TaxonomyFieldValue();
            defaultValue.WssId = -1;
            defaultValue.Label = "Ho Chi Minh";
            defaultValue.TermGuid = "44649e15-7612-432e-b742-147eee391f9b";
            //retrieve validated taxonomy field value
            var validatedValue = taxfield.GetValidatedString(defaultValue);
            context.ExecuteQuery();
            //set default value for a taxonomy field
            taxfield.DefaultValue = validatedValue.Value;
            taxfield.UpdateAndPushChanges(true);
            await context.ExecuteQueryAsync();

            //  Add new data
            List list = context.Web.Lists.GetByTitle("CSOM Test");

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem oListItem = list.AddItem(itemCreateInfo);
            oListItem["Title"] = "Test City 1";
            oListItem["ContentTypeId"] = "0x0100DA5D705795149F44B45648D8D24EAC58"; // Set content type Id
            oListItem.Update();

            ListItemCreationInformation itemCreateInfo2 = new ListItemCreationInformation();
            ListItem oListItem2 = list.AddItem(itemCreateInfo2);
            oListItem2["Title"] = "Test City 2";
            oListItem2["ContentTypeId"] = "0x0100DA5D705795149F44B45648D8D24EAC58"; // Set content type Id
            oListItem2.Update();
            await context.ExecuteQueryAsync();
        }

        private static async Task CamlQueryGetListAboutAsync(ClientContext context)
        {
            var list = context.Web.Lists.GetByTitle("CSOM Test");

            var items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <Where>
                                        <Neq>
                                          <FieldRef Name='About'/>
                                          <Value Type='Text'>Default about</Value>
                                        </Neq>
                                    </Where>
                                </Query>
                                <RowLimit>20</RowLimit>
                            </View>"
                //example for site: https://omniapreprod.sharepoint.com/sites/test-site-duc-11111/
            });

            context.Load(items);
            await context.ExecuteQueryAsync();
        }

        private static async Task CreateViewByCityAsync(ClientContext context)
        {
            string viewQuery = @"<Where><Eq><FieldRef Name='City2' /><Value Type='Text'>Ho Chi Minh</Value></Eq></Where><OrderBy><FieldRef Name='Modified' Ascending='False' /></OrderBy> ";

            var list = context.Web.Lists.GetByTitle("CSOM Test");
            ViewCollection viewColl = list.Views;
            ViewCreationInformation creationInfo = new ViewCreationInformation();
            creationInfo.Title = "CSOM New View 2";
            creationInfo.RowLimit = 50;
            creationInfo.ViewFields = new string[] { "ID", "Title", "City2", "About" };
            creationInfo.ViewTypeKind = ViewType.None;
            creationInfo.SetAsDefaultView = true;
            creationInfo.Query = viewQuery;
            viewColl.Add(creationInfo);

            await context.ExecuteQueryAsync();
        }

        private static async Task UpdateBatchDataAsync(ClientContext context)
        {
            var list = context.Web.Lists.GetByTitle("CSOM Test");
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = @"<View><Query><Where><IsNotNull><FieldRef Name='About' /></IsNotNull></Where><OrderBy><FieldRef Name='Modified' Ascending='False' /></OrderBy> </Query></View>";

            var items = list.GetItems(camlQuery);
            context.Load(items);
            await context.ExecuteQueryAsync();
            if (items.Count >= 1)
            {
                items[0]["About"] = "Update Script";
                items[0]["Title"] = "Update Script";
                items[0].Update();
                if (items.Count >= 2)
                {
                    items[1]["About"] = "Update Script";
                    items[1].Update();
                }
            }
            await context.ExecuteQueryAsync();
        }

        private static async Task CreateFolderAsync(ClientContext context)
        {
            var list = context.Web.Lists.GetByTitle("CSOM Test");
            //Enable Folder creation for the list
            list.EnableFolderCreation = true;
            list.Update();

            var currentFolder = list.RootFolder;
            context.Load(currentFolder, x => x.ServerRelativeUrl);
            await context.ExecuteQueryAsync();

            // Create the folder and sub folder
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
            itemCreateInfo.FolderUrl = currentFolder.ServerRelativeUrl;

            ListItem folder1 = list.AddItem(itemCreateInfo);
            folder1["Title"] = "Folder A1";
            folder1.Update();

            currentFolder = folder1.Folder;
            context.Load(currentFolder, x => x.ServerRelativeUrl);
            await context.ExecuteQueryAsync();

            ListItemCreationInformation itemCreateSubFolderInfo = new ListItemCreationInformation();
            itemCreateSubFolderInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
            itemCreateSubFolderInfo.FolderUrl = currentFolder.ServerRelativeUrl;

            ListItem folder2 = list.AddItem(itemCreateSubFolderInfo);
            folder2["Title"] = "Folder A2";
            folder2.Update();

            context.Load(folder2, x => x.Folder.ServerRelativeUrl);
            await context.ExecuteQueryAsync();

            // Add data sub folder
            ListItemCreationInformation subFolderItemCreateInfo1 = new ListItemCreationInformation();
            subFolderItemCreateInfo1.FolderUrl = folder2.Folder.ServerRelativeUrl;
            ListItem oListItem1 = list.AddItem(subFolderItemCreateInfo1);
            oListItem1["Title"] = "Item 1";
            oListItem1["About"] = "Folder test";
            oListItem1["ContentTypeId"] = "0x0100DA5D705795149F44B45648D8D24EAC58"; // Set content type Id
            oListItem1.Update();

            ListItemCreationInformation subFolderItemCreateInfo2 = new ListItemCreationInformation();
            subFolderItemCreateInfo2.FolderUrl = folder2.Folder.ServerRelativeUrl;
            ListItem oListItem2 = list.AddItem(subFolderItemCreateInfo2);
            oListItem2["Title"] = "Item 2";
            oListItem2["About"] = "Folder test";
            oListItem2["ContentTypeId"] = "0x0100DA5D705795149F44B45648D8D24EAC58";
            oListItem2.Update();

            ListItemCreationInformation subFolderItemCreateInfo3 = new ListItemCreationInformation();
            subFolderItemCreateInfo3.FolderUrl = folder2.Folder.ServerRelativeUrl;
            ListItem oListItem3 = list.AddItem(subFolderItemCreateInfo3);
            oListItem3["Title"] = "Item 3";
            oListItem3["About"] = "Folder test";
            oListItem3["ContentTypeId"] = "0x0100DA5D705795149F44B45648D8D24EAC58";
            oListItem3.Update();

            await context.ExecuteQueryAsync();
        }

        private static async Task AddFieldAuthorAsync(ClientContext context)
        {
            Web web = context.Web;

            var list = web.Lists.GetByTitle("CSOM Test");

            string field = @"<Field Name='CSOMTestAuthor' DisplayName='CSOM Test Author' Type='User' Group='Custom Columns' />";
            list.Fields.AddFieldAsXml(field, true, AddFieldOptions.DefaultValue);

            // Update data
            var allItemQuery = CamlQuery.CreateAllItemsQuery();
            var items = list.GetItems(allItemQuery);
            context.Load(items);
            await context.ExecuteQueryAsync();

            var currentUser = web.CurrentUser;

            foreach(var item in items)
            {
                item["CSOM_x0020_Test_x0020_Author"] = currentUser;
                item.Update();
            }
            await context.ExecuteQueryAsync();
        }

        private static async Task CreateTaxonomyMultipleValueAndSetContentTypeAsync(ClientContext context)
        {
            Web web = context.Web;
            List list = web.Lists.GetByTitle("CSOM Test");

            string fieldSchema = "<Field Type='TaxonomyFieldTypeMulti' DisplayName='Cities' Name='Cities' />";
            web.Fields.AddFieldAsXml(fieldSchema, true, AddFieldOptions.AddToDefaultContentType);

            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            TermGroup termGroup = termStore.Groups.GetByName("Site Collection - adminvn.sharepoint.com-sites-TrainingSharePoint");
            TermSet termSet = termGroup.TermSets.GetByName("city-phong");

            Field field = context.Web.Fields.GetByInternalNameOrTitle("Cities");

            context.Load(termStore, t => t.Id);
            context.Load(termSet);
            context.Load(field);
            await context.ExecuteQueryAsync();

            TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(field);
            taxonomyField.SspId = termStore.Id;
            taxonomyField.TermSetId = termSet.Id;
            taxonomyField.TargetTemplate = String.Empty;
            taxonomyField.AllowMultipleValues = true;
            taxonomyField.UpdateAndPushChanges(true);

            await context.ExecuteQueryAsync();

            // Add field to existed content type
            ContentType contentType = web.ContentTypes.GetById("0x0100DA5D705795149F44B45648D8D24EAC58");

            // Get Site field and add to content type
            FieldCollection fields = web.Fields;
            Field cityField = fields.GetByInternalNameOrTitle("Cities");
            context.Load(cityField);
            await context.ExecuteQueryAsync();

            FieldLinkCreationInformation fldLinkCity = new FieldLinkCreationInformation();
            fldLinkCity.Field = cityField;

            contentType.FieldLinks.Add(fldLinkCity);
            contentType.Update(true);
            await context.ExecuteQueryAsync();
        }

        private static async Task CreateDataForTaxonomiFieldMultipleAsync(ClientContext context)
        {
            Web web = context.Web;
            List list = web.Lists.GetByTitle("CSOM Test");

            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            TermGroup termGroup = termStore.Groups.GetByName("Site Collection - adminvn.sharepoint.com-sites-TrainingSharePoint");
            TermSet termSet = termGroup.TermSets.GetByName("city-phong");

            context.Load(termSet);
            await context.ExecuteQueryAsync();

            // TODO: Update termValueString

            for (var i = 0; i < 3; i++)
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = list.AddItem(itemCreateInfo);
                oListItem["Title"] = "Item test" + i;
                oListItem["ContentTypeId"] = "0x0100DA5D705795149F44B45648D8D24EAC58";
                oListItem.Update();
                context.Load(oListItem);
                await context.ExecuteQueryAsync();

                string termValueString = String.Empty;

                TaxonomyField taxFieldTypeMultiple = context.CastTo<TaxonomyField>(context.Web.Fields.GetByTitle("Cities"));
                var existingTerms = oListItem["Cities"] as TaxonomyFieldValueCollection; // Error here
                foreach (TaxonomyFieldValue tv in existingTerms)
                {
                    termValueString += tv.WssId + ";#" + tv.Label + "|" + tv.TermGuid + ";#";
                }

                TaxonomyFieldValueCollection termValues = new TaxonomyFieldValueCollection(context, termValueString, taxFieldTypeMultiple);

                taxFieldTypeMultiple.SetFieldValueByValueCollection(oListItem, termValues);
                oListItem.Update();
            }
            await context.ExecuteQueryAsync();
        }

        private static async Task CreateDocumentListSync(ClientContext context)
        {
            Web web = context.Web;
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.TemplateType = (int)ListTemplateType.DocumentLibrary;
            creationInfo.Title = "Document Test";
            web.Lists.Add(creationInfo);

            await context.ExecuteQueryAsync();

            // Add content type to existed List
            ContentType contentType = web.ContentTypes.GetById("0x0100DA5D705795149F44B45648D8D24EAC58");

            List csomList = web.Lists.GetByTitle("Document Test");
            csomList.ContentTypes.AddExistingContentType(contentType);
            csomList.Update();
            await context.ExecuteQueryAsync();
        }

        private static async Task CreateFolderInDocumentListAsync(ClientContext context)
        {
            var list = context.Web.Lists.GetByTitle("Document Test");

            // Create the folder and sub folder
            //var currentFolder = list.RootFolder;
            //currentFolder = currentFolder.Folders.Add("Folder 1");
            //currentFolder.Folders.Add("Folder 2");

            //context.Load(currentFolder, x => x.ServerRelativeUrl);
            //await context.ExecuteQueryAsync();

            // TODO: can not add item not type of document in list
            TaxonomyField taxField = context.CastTo<TaxonomyField>(context.Web.Fields.GetByTitle("City2")); ;
            var taxFieldValue = new TaxonomyFieldValue()
            {
                WssId = -1, // alway let it -1
                Label = "Ho Chi Minh",
                TermGuid = "44649e15-7612-432e-b742-147eee391f9b"
            };

            var folder = context.Web.GetFolderByServerRelativeUrl("Document Test/Folder 1/Folder 2");
            context.Load(folder, t => t.ServerRelativeUrl);
            await context.ExecuteQueryAsync();

            //Add data sub folder
            for (var i = 0; i < 5; i++)
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                itemCreateInfo.FolderUrl = folder.ServerRelativeUrl;
                ListItem oListItem = list.AddItem(itemCreateInfo);
                oListItem["Title"] = "Item" + i;
                oListItem["About"] = "Folder test";
                oListItem["ContentTypeId"] = "0x0100DA5D705795149F44B45648D8D24EAC58"; // Set content type Id

                taxField.SetFieldValueByValue(oListItem, taxFieldValue);
                oListItem.Update();
            }
            await context.ExecuteQueryAsync();
        }

        private static async Task CreateDocumentInFolderAsync(ClientContext context)
        {
            //var folder = context.Web.GetFolderByServerRelativeUrl("Document Test/Folder 1/Folder 2");
            //context.Load(folder, t => t.ServerRelativeUrl);
            //await context.ExecuteQueryAsync();

            Web web = context.Web;
            List list = web.Lists.GetByTitle("Document Test");

            context.Load(list.RootFolder, f => f.ServerRelativeUrl);
            context.ExecuteQuery();

            var file = new FileCreationInformation();
            file.Content = System.IO.File.ReadAllBytes(@"D:\Document\Training\Document.docx");
            file.Overwrite = true;
            file.Url =  list.RootFolder.ServerRelativeUrl + "/Document.docx";
            File uploadfile = list.RootFolder.Files.Add(file); // Replace by other folder
            uploadfile.ListItemAllFields["Title"] = "This is new Test Document";
            uploadfile.ListItemAllFields.Update();
            context.Load(uploadfile);
            await context.ExecuteQueryAsync();
        }

        private static async Task CsomTermSetAsync(ClientContext ctx)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName("Test");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("Test Term Set");

            var terms = termSet.GetAllTerms();

            ctx.Load(terms);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CsomLinqAsync(ClientContext ctx)
        {
            var fieldsQuery = from f in ctx.Web.Fields
                              where f.InternalName == "Test" ||
                                    f.TypeAsString == "TaxonomyFieldTypeMulti" ||
                                    f.TypeAsString == "TaxonomyFieldType"
                              select f;

            var fields = ctx.LoadQuery(fieldsQuery);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task SimpleCamlQueryAsync(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle("Documents");

            var allItemsQuery = CamlQuery.CreateAllItemsQuery();
            var allFoldersQuery = CamlQuery.CreateAllFoldersQuery();

            var items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy>
                                </Query>
                                <RowLimit>20</RowLimit>
                            </View>",
                FolderServerRelativeUrl = "/sites/test-site-duc-11111/Shared%20Documents/2"
                //example for site: https://omniapreprod.sharepoint.com/sites/test-site-duc-11111/
            });

            ctx.Load(items);
            await ctx.ExecuteQueryAsync();
        }
    }
}
