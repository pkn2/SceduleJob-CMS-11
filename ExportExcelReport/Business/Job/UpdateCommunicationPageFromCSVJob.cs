using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.Configuration.Attributes;
using EPiServer;
using EPiServer.Core;
using EPiServer.DataAccess;
using EPiServer.Framework.Blobs;
using EPiServer.PlugIn;
using EPiServer.Scheduler;
using EPiServer.Security;
using EPiServer.ServiceLocation;
using EPiServer.Web.Routing;
using ExportExcelReport.Models.Media;
using ExportExcelReport.Models.Pages;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using EPiServer.Globalization;
using EPiServer.Filters;

namespace ExportExcelReport.Business.Job
{
    [ScheduledPlugIn(DisplayName = "Update Communication Page From CSV")]
    public class UpdateCommunicationPageFromCSVJob : ScheduledJobBase
    {
        private bool _stopSignaled;
        private IBlobFactory blobFactory;
        private IContentRepository contentRepository;

        public UpdateCommunicationPageFromCSVJob(IBlobFactory iblob, IContentRepository iContentRepository)
        {
            IsStoppable = true;
            blobFactory = iblob;
            contentRepository = iContentRepository;
        }

        public override void Stop()
        {
            _stopSignaled = true;
        }

        public override string Execute()
        {
            OnStatusChanged(String.Format("Starting execution of {0}", this.GetType()));

            var startPage = contentRepository.Get<StartPage>(ContentReference.StartPage);
            var contentVersionRepository = ServiceLocator.Current.GetInstance<IContentVersionRepository>();

            var folderReference = startPage.FolderReference;
            if (folderReference == null)
            {
                return "Please select Folder reference.";
            }

            var csvFile = contentRepository.GetChildren<GenericMedia>(folderReference)
                .Where(x => x.Name == "Import_All_Communication_Page.csv").FirstOrDefault();
            if(csvFile== null)
            {
                return $"Please add a csv file inside that folder reference. And the file name should be \"Import_All_Communication_Page.csv\"";
            }
            var blob = this.blobFactory.GetBlob(csvFile.BinaryData.ID);

            var update = 0;
            var skip = 0;
            var error = 0;
            var skipPageIds = new List<string>();
            var errorPageIds = new List<string>();

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                BadDataFound = null,
            };
            using (var stream = blob.OpenRead())
            {
                using (var reader = new StreamReader(stream))
                using (var csv = new CsvReader(reader, config))
                {
                    csv.Read();
                    csv.ReadHeader();
                    while (csv.Read())
                    {
                        var page = new ProductPage();
                        try
                        {
                            var record = csv.GetRecord<ImportDataModel>();

                            var pageRef = new PageReference(record.ID);
                            page = contentRepository.Get<ProductPage>(pageRef);

                            var latestVersionb = contentVersionRepository.LoadCommonDraft(pageRef, page.Language.Name).Status;

                            if (!string.IsNullOrEmpty(record.Test)
                                && !string.IsNullOrWhiteSpace(record.Test)
                                && page.Test != record.Test
                                && latestVersionb == VersionStatus.Published)
                            {
                                var isUpdated = UpdateCommunicationPage(page, record);
                                if (isUpdated)
                                {
                                    update++;
                                }
                                else
                                {
                                    error++;
                                    errorPageIds.Add(page.ContentLink.ID.ToString());
                                }
                            }
                            else
                            {
                                skip++;
                                skipPageIds.Add(page.ContentLink.ID.ToString());
                            }
                        }
                        catch
                        {
                            error++;
                            errorPageIds.Add(page.ContentLink.ID.ToString());
                        }
                                         
                    }
                    return $"Import Successfully. Update: {update}, Skipped: {skip}, Error: {error}, Skipped Page Ids: {string.Join(",", skipPageIds)}, Error Page Ids: {string.Join(",", errorPageIds)}";
                }
            }

            if(_stopSignaled)
            {
                return "Stop of job was called";
            }
            return $"Successful";
        }

        public bool UpdateCommunicationPage(ProductPage page, ImportDataModel record)
        {
            //var isUpdated = false;
            try
            {
                var clonePage = page.CreateWritableClone() as ProductPage;
                clonePage.Test = record.Test;
                contentRepository.Save(clonePage, SaveAction.Publish, AccessLevel.NoAccess);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    public class ImportDataModel
    {
        public string Name { get; set; }

        public int ID { get; set; }

        public string FriendlyUrl { get; set; }

        public string Test { get; set; }
    }
}
