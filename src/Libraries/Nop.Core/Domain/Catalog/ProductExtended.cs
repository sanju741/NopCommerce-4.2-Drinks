namespace Nop.Core.Domain.Catalog
{
    /// <summary>
    /// Product entity extended to have refundable price
    /// </summary>
    public partial class Product
    {
        /// <summary>
        /// Refundable Price
        /// </summary>
        public decimal RefundablePrice { get; set; }
    }
}
