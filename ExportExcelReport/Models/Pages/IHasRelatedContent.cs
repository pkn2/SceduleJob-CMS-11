using EPiServer.Core;

namespace ExportExcelReport.Models.Pages
{
    public interface IHasRelatedContent
    {
        ContentArea RelatedContentArea { get; }
    }
}
