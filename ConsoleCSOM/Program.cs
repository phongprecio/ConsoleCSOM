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

                    string userEmail = "user5@adminvn.onmicrosoft.com";
                    string ownerEmail = "phong@adminvn.onmicrosoft.com";

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
                    //await CamlQueryGetListAboutAsync(ctx);
                    //await CreateViewByCityAsync(ctx);
                    //await UpdateBatchDataAsync(ctx);
                    //await CreateFolderAsync(ctx);
                    //await AddFieldAuthorAsync(ctx, userEmail);
                    //await CreateTaxonomyMultipleValueAndSetContentTypeAsync(ctx);
                    //await CreateDataForTaxonomiFieldMultipleAsync(ctx);
                    //await CreateDocumentListSync(ctx);
                    //await CreateFolderInDocumentListAsync(ctx);
                    //await CreateDocumentInFolderAsync(ctx);
                    //await GetListInFolder2Async(ctx);
                    //await CreateViewShowFolderAsync(ctx);
                    //await GetUserAsync(ctx, userEmail);

                    // Permissions
                    await CreateNewPermissionLevelAsync(ctx);
                    await CreateNewGroupAsync(ctx, ownerEmail, userEmail);
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

            // Duplicate name City, so create City2 instead
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
            await context.ExecuteQueryAsync();
        }

        private static async Task AddSampleDataAsync(ClientContext context)
        {
            Web web = context.Web;

            List list = web.Lists.GetByTitle("CSOM Test");
            context.Load(list);
            await context.ExecuteQueryAsync();
            var taxFieldValue = new TaxonomyFieldValue()
            {
                WssId = -1, // alway let it -1
                Label = "Ho Chi Minh",
                TermGuid = "44649e15-7612-432e-b742-147eee391f9b"
            };

            for (var i = 1; i <= 5; i++)
            {
                CreateListItem(context, String.Empty, $"Sample Item {i}", $"About {i}", taxFieldValue);
            }

            await context.ExecuteQueryAsync();
        }

        private static async Task UpdateAboutDefaultAsync(ClientContext context)
        {
            var field = context.Web.Fields.GetByTitle("About");
            field.DefaultValue = "About default";
            field.UpdateAndPushChanges(true);
            await context.ExecuteQueryAsync();

            // Add new data
            CreateListItem(context, String.Empty, "New Item 2.6", String.Empty, null);
            CreateListItem(context, String.Empty, "New Item 2.7", String.Empty, null);
            await context.ExecuteQueryAsync();
        }

        private static async Task AddTermToTermSetAsync(ClientContext context)
        {
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

            // Add new data
            CreateListItem(context, String.Empty, "Test City 1", String.Empty, null);
            CreateListItem(context, String.Empty, "Test City 2", String.Empty, null);
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
            string viewQuery = @"<Where><Eq><FieldRef Name='City2' LookupId='TRUE'/><Value Type='Integer'>1</Value></Eq></Where><OrderBy><FieldRef Name='Modified' Ascending='False' /></OrderBy> ";

            var list = context.Web.Lists.GetByTitle("CSOM Test");
            ViewCollection viewColl = list.Views;
            ViewCreationInformation creationInfo = new ViewCreationInformation();
            creationInfo.Title = "CSOM New View";
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
            camlQuery.ViewXml = @"<View><Query><Where><Eq><FieldRef Name='About' /><Value Type='Text'>About default</Value></Eq></Where><OrderBy><FieldRef Name='Modified' Ascending='False' /></OrderBy> </Query></View>";

            var items = list.GetItems(camlQuery);
            context.Load(items);
            await context.ExecuteQueryAsync();

            for (var i = 0; i < items.Count; i++)
            {
                items[i]["About"] = "Update Script";
                items[i].Update();
                if ((i + 1) % 2 == 0)
                    await context.ExecuteQueryAsync();
            }

            await context.ExecuteQueryAsync();
        }

        private static async Task CreateFolderAsync(ClientContext context)
        {
            var list = context.Web.Lists.GetByTitle("CSOM Test");
            //Enable Folder creation for the list
            list.EnableFolderCreation = true;
            list.Update();

            await context.ExecuteQueryAsync();

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
            context.Load(list, x => x.RootFolder.ServerRelativeUrl);
            await context.ExecuteQueryAsync();
            var folderUrl = list.RootFolder.ServerRelativeUrl + "/Folder A1" + "/Folder A2";
            for (var i = 1; i <= 3; i++)
            {
                CreateListItem(context, folderUrl, $"Item F{i}", String.Empty, null);
            }

            await context.ExecuteQueryAsync();
        }

        private static async Task AddFieldAuthorAsync(ClientContext context, string userEmail)
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

            var user = web.EnsureUser(userEmail);

            foreach(var item in items)
            {
                item["CSOM_x0020_Test_x0020_Author"] = user;
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
            string termValueString = $"-1;#Stockholm|1ce8d481-fec5-45ff-8b58-6c96996f6aa7;#-1;#Term1|0ac925f1-5c5a-422f-ba1f-a98d0ac3b67e";

            // TODO truyen them Id vao param de lay thay vi dung title
            TaxonomyField taxFieldTypeMultiple = context.CastTo<TaxonomyField>(context.Web.Fields.GetByTitle("Cities"));

            TaxonomyFieldValueCollection termValues = new TaxonomyFieldValueCollection(context, termValueString, taxFieldTypeMultiple);

            for (var i = 1; i <= 3; i++)
            {
                CreateListItem(context, String.Empty, $"Item Cities {i}", String.Empty, null, taxFieldTypeMultiple, termValues);
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
            var currentFolder = list.RootFolder;
            currentFolder = currentFolder.Folders.Add("Folder 1");
            currentFolder.Folders.Add("Folder 2");

            context.Load(currentFolder, x => x.ServerRelativeUrl);
            await context.ExecuteQueryAsync();

            var folder = context.Web.GetFolderByServerRelativeUrl("Document Test/Folder 1/Folder 2");
            context.Load(folder, t => t.ServerRelativeUrl);
            await context.ExecuteQueryAsync();

            var url = folder.ServerRelativeUrl + "/Document.docx";
            var url2 = folder.ServerRelativeUrl + "/Testdocument.docx";
            var content = System.IO.File.ReadAllBytes(@"D:\Document\Training\Document.docx");
            var content2 = System.IO.File.ReadAllBytes(@"D:\Document\Training\Testdocument.docx");

            string termValueString = $"-1#Stockholm|1ce8d481-fec5-45ff-8b58-6c96996f6aa7";

            TaxonomyField taxFieldTypeMultiple = context.CastTo<TaxonomyField>(context.Web.Fields.GetByTitle("Cities"));

            TaxonomyFieldValueCollection termValues = new TaxonomyFieldValueCollection(context, termValueString, taxFieldTypeMultiple);

            CreateFile(context, url, folder, content, $"This is new Test Document 1", "About", null, taxFieldTypeMultiple, termValues);
            CreateFile(context, url2, folder, content2, $"This is new Test Document 2", "Test", null, taxFieldTypeMultiple, termValues);
            await context.ExecuteQueryAsync();
        }

        private static async Task GetListInFolder2Async(ClientContext context)
        {
            var list = context.Web.Lists.GetByTitle("Document Test");

            var folder = context.Web.GetFolderByServerRelativeUrl("Document Test/Folder 1/Folder 2");
            context.Load(folder, t => t.ServerRelativeUrl);
            await context.ExecuteQueryAsync();

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = @"
                                    <View>
                                        <Query>
                                            <Where>
                                                <In>
                                                    <FieldRef LookupId='TRUE' Name='Cities'/>
                                                    <Values>
                                                        <Value Type='Integer'>2</Value>
                                                    </Values>
                                                </In>
                                            </Where>
                                        </Query>
                                    </View>
                                ";
            camlQuery.FolderServerRelativeUrl = folder.ServerRelativeUrl;
            var items = list.GetItems(camlQuery);
            context.Load(items);
            await context.ExecuteQueryAsync();
        }

        private static async Task CreateViewShowFolderAsync(ClientContext context)
        {
            string viewQuery = @"<Where>
                                    <BeginsWith>
                                        <FieldRef Name='FSObjType' />
                                        <Value Type='Integer'>1</Value>
                                    </BeginsWith>
                                </Where>";

            var list = context.Web.Lists.GetByTitle("Document Test");
            ViewCollection viewColl = list.Views;
            ViewCreationInformation creationInfo = new ViewCreationInformation();
            creationInfo.Title = "Folders";
            creationInfo.RowLimit = 50;
            creationInfo.ViewFields = new string[] { "Type", "Name", "Title", "City2", "About" };
            creationInfo.ViewTypeKind = ViewType.None;
            creationInfo.SetAsDefaultView = true;
            creationInfo.Query = viewQuery;
            viewColl.Add(creationInfo);

            await context.ExecuteQueryAsync();
        }

        private static async Task GetUserAsync(ClientContext context, string userEmail)
        {
            var user = context.Web.EnsureUser(userEmail);
            context.Load(user);
            await context.ExecuteQueryAsync();
        }

        private static void CreateListItem(ClientContext context, string folderUrl, string title, string about, TaxonomyFieldValue taxFieldValue)
        {
            CreateListItem(context, folderUrl, title, about, taxFieldValue, null, null);
        }

        private static void CreateListItem(ClientContext context, string folderUrl, string title, string about, TaxonomyFieldValue taxFieldValue, TaxonomyField taxFieldTypeMultiple, TaxonomyFieldValueCollection taxonomyFieldValues)
        {
            var list = context.Web.Lists.GetByTitle("CSOM Test");
            TaxonomyField cityField = context.CastTo<TaxonomyField>(context.Web.Fields.GetByTitle("City2"));

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

            if (!string.IsNullOrEmpty(folderUrl))
                itemCreateInfo.FolderUrl = folderUrl;

            ListItem oListItem = list.AddItem(itemCreateInfo);
            oListItem["ContentTypeId"] = "0x0100DA5D705795149F44B45648D8D24EAC58"; // Set content type Id

            if (!string.IsNullOrEmpty(title))
                oListItem["Title"] = title;

            if (!string.IsNullOrEmpty(about))
                oListItem["About"] = about;

            if (taxFieldValue != null)
                cityField.SetFieldValueByValue(oListItem, taxFieldValue);

            if (taxFieldTypeMultiple != null && taxonomyFieldValues != null)
                taxFieldTypeMultiple.SetFieldValueByValueCollection(oListItem, taxonomyFieldValues);

            oListItem.Update();
        }

        private static void CreateFile(ClientContext context, string url, Folder folderUploadFile, byte[] contents, string title, string about, TaxonomyFieldValue taxFieldValue, TaxonomyField taxFieldTypeMultiple, TaxonomyFieldValueCollection taxonomyFieldValues)
        {
            TaxonomyField cityField = context.CastTo<TaxonomyField>(context.Web.Fields.GetByTitle("City2"));
            var file = new FileCreationInformation();
            file.Content = contents;
            file.Overwrite = true;
            file.Url = url;
            File uploadfile = folderUploadFile.Files.Add(file);

            if (!string.IsNullOrEmpty(title))
                uploadfile.ListItemAllFields["Title"] = title;

            if (!string.IsNullOrEmpty(about))
                uploadfile.ListItemAllFields["About"] = about;

            if (taxFieldValue != null)
                cityField.SetFieldValueByValue(uploadfile.ListItemAllFields, taxFieldValue);

            if (taxFieldTypeMultiple != null && taxonomyFieldValues != null)
                taxFieldTypeMultiple.SetFieldValueByValueCollection(uploadfile.ListItemAllFields, taxonomyFieldValues);

            uploadfile.ListItemAllFields.Update();
            context.Load(uploadfile);
        }

        private static async Task CreateDocumentInFolderAsync(ClientContext context)
        {
            List list = context.Web.Lists.GetByTitle("Document Test");

            context.Load(list.RootFolder, f => f.ServerRelativeUrl);
            context.ExecuteQuery();

            var taxFieldValue = new TaxonomyFieldValue()
            {
                WssId = -1, // alway let it -1
                Label = "Ho Chi Minh",
                TermGuid = "44649e15-7612-432e-b742-147eee391f9b"
            };

            var contents = System.IO.File.ReadAllBytes(@"D:\Document\Training\Document.docx");
            var url =  list.RootFolder.ServerRelativeUrl + "/Document.docx";

            CreateFile(context, url, list.RootFolder, contents, "This is new Test Document", String.Empty, taxFieldValue, null, null);
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


        // Permissions
        private static async Task GetDefaultSecurityGroup(ClientContext context)
        {
            var subSite = context.Site.OpenWeb("FA/");
            var siteGroups = subSite.SiteGroups;
            
        }

        private static async Task CreateNewPermissionLevelAsync(ClientContext context)
        {
            BasePermissions perm = new BasePermissions();
            perm.Set(PermissionKind.CreateAlerts);
            perm.Set(PermissionKind.ManageLists);
            perm.Set(PermissionKind.ViewListItems);
            perm.Set(PermissionKind.ViewPages);
            perm.Set(PermissionKind.Open);
            perm.Set(PermissionKind.OpenItems);

            RoleDefinitionCreationInformation creationInfo = new RoleDefinitionCreationInformation();
            creationInfo.BasePermissions = perm;
            creationInfo.Description = "A role with create and manage alerts permission";
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
