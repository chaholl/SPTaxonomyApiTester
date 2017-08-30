using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace SpTaxonomyApiTester
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                try
                {
                    ConnectToWeb(args[0]);
                }
                catch (Exception e)
                {
                    var oldColor = Console.ForegroundColor;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(e);
                    Console.ForegroundColor = oldColor;
                }
            }
            else
            {
                Console.WriteLine("Please supply a SharePoint Online URL to check");
            }

            Console.WriteLine("Press return to exit");
            Console.ReadLine();
        }

        private static void ConnectToWeb(string url)
        {
            using (var ctx = TokenHelper.GetAppOnlyClientContextForUrl(url))
            {
                if (ctx == null) throw new Exception("Unable to create client context");
                var web = ctx.Web;
                var props = web.AllProperties;
                ctx.Load(web);
                ctx.Load(props);
                ctx.ExecuteQuery();
                Console.WriteLine($"Connected to {ctx.Web.Url}");

                //ListInstalledApps(ctx);

                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(web.Context);
                TermStoreCollection termStores = taxonomySession.TermStores;

                web.Context.Load(termStores, t => t.Include(v => v.DefaultLanguage, v => v.Id, v => v.Name));
                web.Context.ExecuteQuery();

                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine("######################################################");
                Console.WriteLine("   Term Stores");
                Console.WriteLine("######################################################");

                var termGuid = new Guid("B6C16ADB-ACB6-4097-81A8-4EBA3AE4C89E");
                var term2Guid = new Guid("31D4092B-34D2-456D-B057-7976ECEA60D7");
                var testGroupid = new Guid("D2D3E7DF-C90A-430A-A062-81B3C0941BF8");
                var termsetGuid = new Guid("DA4CCAFE-F797-4558-90D1-DCFB14A281BE");

                foreach (var termstore in termStores)
                {
                    try
                    {
                        Console.WriteLine("{0} pointing to {1}", termstore.Id, termstore.Name);

                        WriteMessage("Creating termset in site collection group.....");

                        var siteGroup = termstore.GetSiteCollectionGroup(ctx.Site, true);
                        var newTermSet = siteGroup.CreateTermSet("Some new termset", termsetGuid, termstore.DefaultLanguage);
                        ctx.Load(newTermSet);
                        ctx.ExecuteQuery();

                        WriteSuccess("Done");

                        WriteMessage("Getting new termset from termstore.....");

                        var existingTermset = termstore.GetTermSet(termsetGuid);
                        ctx.Load(existingTermset);
                        ctx.ExecuteQuery();

                        if (existingTermset == null)
                        {
                            WriteFailure("Not found");

                            try
                            {
                                WriteMessage("No termset found. Creating group and adding termset...");

                                var newGroup = termstore.CreateGroup("Test Group", testGroupid);
                                var newTermSet2 = newGroup.CreateTermSet("Some new termset", termsetGuid, termstore.DefaultLanguage);
                                ctx.Load(newGroup);
                                ctx.Load(newTermSet2);
                                ctx.ExecuteQuery();

                                WriteSuccess("Done");
                            }
                            catch (Exception ex)
                            {
                                WriteFailure($"Failed");
                                WriteException(ex);
                            }
                        }
                        else
                        {
                            WriteSuccess("Found");
                        }

                        WriteMessage("Creating a term in site collection group....");

                        var newterm = existingTermset.CreateTerm("Some new term", termstore.DefaultLanguage, termGuid);
                        ctx.Load(newterm);
                        ctx.ExecuteQuery();

                        WriteSuccess("Done");

                        WriteMessage("Finding term in termset.....");
                        var existingterm = termstore.GetTerm(termGuid);
                        ctx.Load(existingterm);
                        ctx.ExecuteQuery();

                        if (existingterm == null)
                        {
                            WriteFailure("Not found");
                        }
                        else
                        {
                            WriteSuccess("Found");
                        }

                        WriteMessage("Create a term with a specific GUID.......");

                        var newterm2 = existingTermset.CreateTerm("Some new term 2", termstore.DefaultLanguage, term2Guid);
                        ctx.Load(newterm2);
                        ctx.ExecuteQuery();

                        WriteSuccess("Done");

                        try
                        {
                            WriteMessage("Create another term with the same GUID.......");

                            var anotherTermWithTheSameGuid = existingTermset.CreateTerm("Some new term 3", termstore.DefaultLanguage, term2Guid);
                            ctx.Load(anotherTermWithTheSameGuid);
                            ctx.ExecuteQuery();
                            WriteSuccess("Done");
                        }
                        catch (Exception ex)
                        {
                            WriteFailure("Failed");
                            WriteException(ex);
                        }

                        CreateSmallBatchWithDuplicate(ctx, term2Guid, termstore, existingTermset);

                        //CreateLargeBatch(ctx, termstore, existingTermset);

                        //ManySmallRequests(ctx, termstore, siteGroup);

                    }
                    finally
                    {
                        var existingTermset = termstore.GetTermSet(termsetGuid);
                        existingTermset.DeleteObject();
                        ctx.ExecuteQuery();
                        Console.WriteLine("Cleaned stuff up");
                    }

                }
            }
        }

        private static void CreateSmallBatchWithDuplicate(ClientContext ctx, Guid term2Guid, TermStore termstore, TermSet existingTermset)
        {
            try
            {
                WriteMessage("Creating a batch of 20 terms where one has a duplicate ID......");
                for (int x = 0; x < 19; x++)
                {
                    var smallBatchTerm = existingTermset.CreateTerm($"SB Term {x}", termstore.DefaultLanguage, Guid.NewGuid());
                    ctx.Load(smallBatchTerm);
                }
                var termWithTheSameGuid = existingTermset.CreateTerm("SB Term 20", termstore.DefaultLanguage, term2Guid);
                ctx.Load(termWithTheSameGuid);
                ctx.ExecuteQuery();
                WriteSuccess("Done");
            }
            catch (Exception ex)
            {
                WriteFailure("Failed. No way to know which term has a problem");
                WriteException(ex);
            }
        }

        private static void CreateLargeBatch(ClientContext ctx, TermStore termstore, TermSet existingTermset)
        {
            try
            {
                WriteMessage("Creating 10,000 terms in batches of 100.......");
                for (int x = 0; x < 100; x++)
                {
                    var someTerm = existingTermset.CreateTerm($"Term {x}", termstore.DefaultLanguage, Guid.NewGuid());
                    ctx.Load(someTerm);
                    for (int level1 = 0; level1 < 100; level1++)
                    {
                        var level1Term = someTerm.CreateTerm($"Level 1 term {level1}", termstore.DefaultLanguage, Guid.NewGuid());
                        ctx.Load(level1Term);
                    }
                    ctx.ExecuteQuery();
                }
                WriteSuccess("Done");
            }
            catch (Exception ex)
            {
                WriteFailure("Failed");
                WriteException(ex);
            }
        }

        private static void ManySmallRequests(ClientContext ctx, TermStore termstore, TermGroup sitegroup)
        {
            var newTermSet = sitegroup.CreateTermSet("ManySmallRequests", Guid.NewGuid(), termstore.DefaultLanguage);
            ctx.Load(newTermSet);
            ctx.ExecuteQuery();

            try
            {
                WriteMessage("Creating 100,000 terms individually.......");
                for (int x = 0; x < 10; x++)
                {
                    var someTerm = newTermSet.CreateTerm($"Term {x}", termstore.DefaultLanguage, Guid.NewGuid());
                    ctx.Load(someTerm);
                    ctx.ExecuteQuery();
                    for (int level1 = 0; level1 < 100; level1++)
                    {
                        var level1Term = someTerm.CreateTerm($"Level 1 term {level1}", termstore.DefaultLanguage, Guid.NewGuid());
                        ctx.Load(level1Term);
                        ctx.ExecuteQuery();
                        for (int level2 = 0; level2 < 100; level2++)
                        {
                            var level2Term = level1Term.CreateTerm($"Level 2 term {level2}", termstore.DefaultLanguage, Guid.NewGuid());
                            ctx.Load(level2Term);
                            ctx.ExecuteQuery();
                        }
                    }                    
                }
                WriteSuccess("Done");
            }
            catch (Exception ex)
            {
                WriteFailure("Failed");
                WriteException(ex);
            }

            newTermSet.DeleteObject();
            ctx.ExecuteQuery();
            
        }

        private static void ListInstalledApps(ClientContext ctx)
        {
            var apps = AppCatalog.GetAppInstances(ctx, ctx.Web);
            ctx.Load(apps);
            ctx.ExecuteQuery();

            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("######################################################");
            Console.WriteLine("   Installed Add-Ins");
            Console.WriteLine("######################################################");
            //print info
            foreach (var app in apps)
            {
                Console.WriteLine("Name: {0}, ProductId:{1}, InstanceId: {2}, Status: {3}", app.Title, app.ProductId, app.Id, app.Status);
            }
        }

        private static void WriteMessage(string message)
        {
            var color = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write(message);
            Console.ForegroundColor = color;
        }

        private static void WriteSuccess(string message)
        {
            var color = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(message);
            Console.ForegroundColor = color;
        }

        private static void WriteFailure(string message)
        {
            var color = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(message);
            Console.ForegroundColor = color;
        }

        private static void WriteException(Exception ex)
        {
            var color = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(ex);
            Console.ForegroundColor = color;
        }
    }
}
