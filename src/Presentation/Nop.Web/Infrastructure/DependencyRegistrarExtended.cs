using Autofac;
using Nop.Core.Configuration;
using Nop.Core.Infrastructure;
using Nop.Core.Infrastructure.DependencyManagement;
using Nop.Web.Areas.Admin.Factories;

namespace Nop.Web.Infrastructure
{
    public class DependencyRegistrarExtended: IDependencyRegistrar
    {
        public void Register(ContainerBuilder builder, ITypeFinder typeFinder, NopConfig config)
        {
            builder.RegisterType<ProductModelFactoryExtended>().As<IProductModelFactory>().InstancePerLifetimeScope();
        }

        public int Order => 3;
    }
}
