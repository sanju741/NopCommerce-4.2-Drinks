using System.IO;

namespace Nop.Services.ExportImport
{
    public partial interface IImportManager
    {
        /// <summary>
        /// Migrate existing products in different excel format to the system
        /// </summary>
        /// <param name="stream"></param>
        void MigrateProductsFromXlsx(Stream stream);

    }
}