using CsvHelper;
using EPiServer;
using EPiServer.Core;
using EPiServer.Find;
using EPiServer.Framework.Blobs;
using EPiServer.PlugIn;
using EPiServer.Scheduler;
using EPiServer.ServiceLocation;
using EPiServer.Web.Routing;
using ExportExcelReport.Models.Media;
using ExportExcelReport.Models.Pages;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web.Http.Routing;

namespace ExportExcelReport.Business.Job
{
    [ScheduledPlugIn(DisplayName = "ExportReport")]
    public class ExportReport : ScheduledJobBase
    {
        private bool _stopSignaled;
        private IBlobFactory blobFactory;

        public ExportReport(IBlobFactory iblob)
        {
            IsStoppable = true;
            blobFactory = iblob;
        }

        public override void Stop()
        {
            _stopSignaled = true;
        }

        public override string Execute()
        {
            //Call OnStatusChanged to periodically notify progress of job for manually started jobs
            OnStatusChanged(String.Format("Starting execution of {0}", this.GetType()));

            var contentRepository = ServiceLocator.Current.GetInstance<IContentRepository>();
            var urlHelper = ServiceLocator.Current.GetInstance<UrlResolver>();
            
            var startPage = contentRepository.Get<StartPage>(ContentReference.StartPage);

            var folderReference = startPage.FolderReference;
            if (folderReference == null)
            {
                return "Please select Folder reference.";
            }

            var reportFile = contentRepository.GetDefault<GenericMedia>(folderReference);
            reportFile.Name = $"Export_All_Communication_Page.csv";

            var blob = this.blobFactory.CreateBlob(reportFile.BinaryDataContainer, ".csv");

            var getAllPage = new GetAllPage();
            var allComunicationPage = getAllPage.AllComunicationPage().ToList();

            using (var stream = blob.OpenWrite())
            {
                var writer = new StreamWriter(stream);
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    csv.WriteRecords(allComunicationPage.Select(x => new { x.Name, x.ContentLink.ID, FriendlyUrl = urlHelper.GetVirtualPath(x.ContentLink).VirtualPath, x.Test }));
                    writer.Flush();
                }
            }

            reportFile.BinaryData = blob;

            var res = contentRepository.Publish(reportFile, EPiServer.Security.AccessLevel.NoAccess);
            //For long running jobs periodically check if stop is signaled and if so stop execution
            if (_stopSignaled)
            {
                return "Stop of job was called";
            }

            return $"Export Successfully. Total no of communication Page: {allComunicationPage.Count}";
        }
    }
}
