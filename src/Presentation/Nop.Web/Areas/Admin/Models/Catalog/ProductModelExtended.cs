using Nop.Web.Framework.Mvc.ModelBinding;

namespace Nop.Web.Areas.Admin.Models.Catalog
{
    /// <summary>
    /// Extended property
    /// </summary>
    public partial class ProductModel
    {
        [NopResourceDisplayName("Admin.Catalog.Products.Fields.RefundablePrice")]
        public decimal RefundablePrice { get; set; }
    }
}
