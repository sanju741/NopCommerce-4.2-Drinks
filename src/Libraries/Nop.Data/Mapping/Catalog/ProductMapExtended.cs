using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using Nop.Core.Domain.Catalog;

namespace Nop.Data.Mapping.Catalog
{
    /// <summary>
    /// Map the new refundable price column
    /// </summary>
    public partial class ProductMapExtended : NopEntityTypeConfiguration<Product>
    {
        public override void Configure(EntityTypeBuilder<Product> builder)
        {
            builder.Property(product => product.RefundablePrice).HasColumnType("decimal(18, 4)");
        }
    }
}
