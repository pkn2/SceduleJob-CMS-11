using EPiServer.Find;
using EPiServer.Find.Cms;
using EPiServer.Find.Framework;
using ExportExcelReport.Models.Pages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportExcelReport.Business.Job
{
    public class GetAllPage
    {
        public IEnumerable<ProductPage> AllComunicationPage()
        {
            int totalNumberOfPages;

            var query = SearchClient.Instance.Search<ProductPage>()
                                        .Take(1000);

            var batch = query.GetContentResult();
            foreach (var page in batch)
            {
                yield return page;
            }

            totalNumberOfPages = batch.TotalMatching;

            var nextBatchFrom = 1000;
            while (nextBatchFrom < totalNumberOfPages)
            {
                batch = query.Skip(nextBatchFrom).GetContentResult();
                foreach (var page in batch)
                {
                    yield return page;
                }
                nextBatchFrom += 1000;
            }
        }
    }
   
}