using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using Microsoft.Extensions.DependencyInjection;
using Nop.Core;
using Nop.Core.Data;
using Nop.Core.Domain.Catalog;
using Nop.Core.Domain.Customers;
using Nop.Core.Domain.Media;
using Nop.Core.Domain.Tax;
using Nop.Core.Domain.Vendors;
using Nop.Core.Infrastructure;
using Nop.Services.Catalog;
using Nop.Services.Customers;
using Nop.Services.Directory;
using Nop.Services.ExportImport.Help;
using Nop.Services.Localization;
using Nop.Services.Logging;
using Nop.Services.Media;
using Nop.Services.Messages;
using Nop.Services.Security;
using Nop.Services.Seo;
using Nop.Services.Shipping;
using Nop.Services.Shipping.Date;
using Nop.Services.Stores;
using Nop.Services.Tax;
using Nop.Services.Vendors;
using OfficeOpenXml;
using StackExchange.Profiling.Internal;

namespace Nop.Services.ExportImport
{
    public partial class ImportManager
    {
        private readonly ICustomerService _customerService;

        public ImportManager(
            CatalogSettings catalogSettings,
            ICategoryService categoryService,
            ICountryService countryService,
            ICustomerActivityService customerActivityService,
            IDataProvider dataProvider,
            IDateRangeService dateRangeService,
            IEncryptionService encryptionService,
            IHttpClientFactory httpClientFactory,
            ILocalizationService localizationService,
            ILogger logger,
            IManufacturerService manufacturerService,
            IMeasureService measureService,
            INewsLetterSubscriptionService newsLetterSubscriptionService,
            INopFileProvider fileProvider,
            IPictureService pictureService,
            IProductAttributeService productAttributeService,
            IProductService productService,
            IProductTagService productTagService,
            IProductTemplateService productTemplateService,
            IServiceScopeFactory serviceScopeFactory,
            IShippingService shippingService,
            ISpecificationAttributeService specificationAttributeService,
            IStateProvinceService stateProvinceService,
            IStoreContext storeContext,
            IStoreMappingService storeMappingService,
            IStoreService storeService,
            ITaxCategoryService taxCategoryService,
            IUrlRecordService urlRecordService,
            IVendorService vendorService,
            IWorkContext workContext,
            MediaSettings mediaSettings,
            VendorSettings vendorSettings,
            ICustomerService customerService) : this(
             catalogSettings,
             categoryService,
             countryService,
             customerActivityService,
             dataProvider,
             dateRangeService,
             encryptionService,
             httpClientFactory,
             localizationService,
             logger,
             manufacturerService,
             measureService,
             newsLetterSubscriptionService,
             fileProvider,
             pictureService,
             productAttributeService,
             productService,
             productTagService,
             productTemplateService,
             serviceScopeFactory,
             shippingService,
             specificationAttributeService,
             stateProvinceService,
             storeContext,
             storeMappingService,
             storeService,
             taxCategoryService,
             urlRecordService,
             vendorService,
             workContext,
             mediaSettings,
             vendorSettings)
        {
            _customerService = customerService;
        }

        protected Category CreateCategoryAndSlug(string categoryName, int parentId = 0)
        {
            var category = new Category
            {
                Name = categoryName,
                ParentCategoryId = parentId,
                Published = true,
                DisplayOrder = 0,
                IncludeInTopMenu = true,
                AllowCustomersToSelectPageSize = true,
                PageSizeOptions = "6,3,9"
            };
            //save category and its slug
            _categoryService.InsertCategory(category);

            string categorySeName = _urlRecordService.ValidateSeName(category, "", category.Name, true);
            _urlRecordService.SaveSlug(category, categorySeName, 0);
            return category;
        }

        /// <summary>
        /// Import products from XLSX file
        /// </summary>
        /// <param name="stream">Stream</param>
        public virtual void MigrateProductsFromXlsx(Stream stream)
        {
            using (var xlPackage = new ExcelPackage(stream))
            {
                // get the first worksheet in the workbook
                var worksheet = xlPackage.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                    throw new NopException("No worksheet found");


                var metadata = PrepareImportProductDataForMigration(worksheet);

                if (_catalogSettings.ExportImportSplitProductsFile && metadata.CountProductsInFile > _catalogSettings.ExportImportProductsCountInOneFile)
                {
                    MigrateProductsFromSplitedXlsx(worksheet, metadata);
                    return;
                }

                //performance optimization, load all categories in one SQL request
                IList<Category> allCategoryList = _categoryService.GetAllCategories(showHidden: true, loadCacheableCopy: false);


                //performance optimization, load all manufacturers in one SQL request
                var allManufacturers = _manufacturerService.GetAllManufacturers(showHidden: true);

                var specificationAttributesInImport = new List<string>
                {
                    "Type",
                    "Segment",
                    "Group",
                    "SegmentID",
                    "SubSegment",
                    "Number of units",
                    "Volume (in liter)",
                    "Unit"
                };
                //Specification attributes
                List<SpecificationAttribute> existingSpecificationAttributes = _specificationAttributeService
                    .GetSpecificationAttributes().ToList();

                foreach (string attributeInImport in specificationAttributesInImport)
                {
                    if (existingSpecificationAttributes.Any(x => x.Name.Equals(attributeInImport)))
                        continue;

                    //insert specification attribute
                    var specificationAttribute = new SpecificationAttribute()
                    {
                        Name = attributeInImport,
                        DisplayOrder = 0,
                    };
                    _specificationAttributeService.InsertSpecificationAttribute(specificationAttribute);

                    //insert specification attribute option
                    var specificationAttributeOption = new SpecificationAttributeOption
                    {
                        SpecificationAttributeId = specificationAttribute.Id,
                        Name = attributeInImport,
                        DisplayOrder = 0,
                    };
                    _specificationAttributeService.InsertSpecificationAttributeOption(specificationAttributeOption);

                    //update the existing specification attributes 
                    existingSpecificationAttributes.Add(specificationAttribute);
                }

                //tier price customer role mapping
                List<CustomerRole> allB2BRoles = _customerService.GetAllCustomerRoles()
                    .Where(x => x.SystemName.Contains("B2BLevel-")).ToList();

                var tierPriceDictionary = new Dictionary<string, string>()
                {
                    {"B2BLevel-1", "SellingPrice1"},
                    {"B2BLevel-2", "SellingPrice2"},
                    {"B2BLevel-3", "SellingPrice3"},
                    {"B2BLevel-4", "SellingPrice4"},
                    {"B2BLevel-5", "SellingPrice5"},
                    {"B2BLevel-6", "SellingPrice6"},
                    {"B2BLevel-7", "SellingPrice7"},
                    {"B2BLevel-8", "SellingPrice8"},
                    {"B2BLevel-9", "SellingPrice9"},
                };

                //tax category
                IList<TaxCategory> allTaxCategories = _taxCategoryService.GetAllTaxCategories();
                string taxCategoryPrefix = "Drinks {0}%";

                //todo product pictures
                ////product to import images
                //var productPictureMetadata = new List<ProductPictureMetadata>();

                Product lastLoadedProduct = null;

                //common value preparation
                int simpleTemplateId = _productTemplateService.GetAllProductTemplates()
                    .First(x => x.Name == "Simple product").Id;
                for (var iRow = 2; iRow < metadata.EndRow; iRow++)
                {
                    metadata.Manager.ReadFromXlsx(worksheet, iRow);

                    var product = new Product();
                    //base value from excel
                    foreach (var property in metadata.Manager.GetProperties)
                    {
                        switch (property.PropertyName)
                        {
                            case "Name":
                                product.Name = property.StringValue;
                                break;
                            case "ArtikelID":
                                product.Sku = property.StringValue;
                                break;
                            case "SellingPrice0":
                                product.Price = property.DecimalValue;
                                break;
                            case "Exise":
                                product.AdminComment = $"Exise: {property.StringValue}";
                                break;
                            case "Price Refundable packaging":
                                product.RefundablePrice = property.DecimalValue;
                                break;
                        }
                    }
                    //the product already exist in the system no need to process further
                    if (_productService.GetProductBySku(product.Sku) != null)
                        continue;

                    //default values
                    product.CreatedOnUtc = DateTime.UtcNow;
                    product.UpdatedOnUtc = DateTime.UtcNow;
                    product.ProductType = ProductType.SimpleProduct;
                    product.ParentGroupedProductId = 0;
                    product.VisibleIndividually = true;
                    product.ProductTemplateId = simpleTemplateId;
                    product.AllowCustomerReviews = true;
                    product.Published = true;
                    product.IsShipEnabled = true;
                    product.StockQuantity = 1000;
                    product.NotifyAdminForQuantityBelow = 1;
                    product.OrderMinimumQuantity = 1;
                    product.OrderMaximumQuantity = int.MaxValue - 1;
                    product.ManageInventoryMethod = ManageInventoryMethod.ManageStock;
                    product.BackorderMode = BackorderMode.NoBackorders;
                    //disable shipping by default
                    product.IsShipEnabled = false;

                    //tax
                    var taxProperty = metadata.Manager.GetProperty("Percent VAT");
                    if (taxProperty != null && !string.IsNullOrWhiteSpace(taxProperty.StringValue))
                    {
                        var taxPercent = taxProperty.DecimalValue;

                        string taxCategoryName = string.Format(taxCategoryPrefix, taxPercent);
                        var taxCategory = allTaxCategories.FirstOrDefault(x => x.Name.Equals(taxCategoryName));
                        if (taxCategory == null)
                        {
                            taxCategory = new TaxCategory
                            {
                                Name = taxCategoryName,
                            };
                            _taxCategoryService.InsertTaxCategory(taxCategory);
                            allTaxCategories.Add(taxCategory);
                        }
                        product.TaxCategoryId = taxCategory.Id;
                    }

                    //Insert the product
                    _productService.InsertProduct(product);

                    //quantity change history
                    _productService.AddStockQuantityHistoryEntry(product, product.StockQuantity, product.StockQuantity,
                        product.WarehouseId, _localizationService.GetResource("Admin.StockQuantityHistory.Messages.ImportProduct.Edit"));

                    //search engine name
                    var seName = _urlRecordService.GetSeName(product, 0);
                    _urlRecordService.SaveSlug(product, _urlRecordService.ValidateSeName(product, seName, product.Name, true), 0);

                    //categories
                    var tempProperty = metadata.Manager.GetProperty("Brand");
                    if (tempProperty != null && !string.IsNullOrWhiteSpace(tempProperty.StringValue))
                    {
                        var categoryName = tempProperty.StringValue.Trim();
                        int parentCategoryId = 0;

                        //great grand parent category
                        tempProperty = metadata.Manager.GetProperty("Group");
                        var greatGrandParentName = tempProperty.StringValue.Trim();
                        var greatGrandParent = allCategoryList
                            .FirstOrDefault(x =>
                                x.Name.Equals(greatGrandParentName) && x.ParentCategoryId == parentCategoryId);
                        if (greatGrandParent == null)
                        {
                            var category = CreateCategoryAndSlug(greatGrandParentName, parentCategoryId);

                            //Update existing list and parent category id
                            allCategoryList.Add(category);
                            parentCategoryId = category.Id;
                        }
                        else
                        {
                            parentCategoryId = greatGrandParent.Id;
                        }

                        //grand parent category
                        tempProperty = metadata.Manager.GetProperty("SegmentID");
                        var grandParentName = tempProperty.StringValue.Trim();
                        var grandParent = allCategoryList.FirstOrDefault(x =>
                            x.Name.Equals(grandParentName) && x.ParentCategoryId == parentCategoryId);
                        if (grandParent == null)
                        {
                            var category = CreateCategoryAndSlug(grandParentName, parentCategoryId);

                            // Update existing list and parent category id
                            allCategoryList.Add(category);
                            parentCategoryId = category.Id;
                        }
                        else
                        {
                            parentCategoryId = grandParent.Id;
                        }

                        //parent category
                        tempProperty = metadata.Manager.GetProperty("SubSegment");
                        var parentName = tempProperty.StringValue.Trim();
                        var parent = allCategoryList.FirstOrDefault(x =>
                                         x.Name.Equals(parentName) && x.ParentCategoryId == parentCategoryId);
                        if (parent == null && !string.IsNullOrWhiteSpace(parentName))
                        {
                            var category = CreateCategoryAndSlug(parentName, parentCategoryId);

                            // Update existing list and parent category id
                            allCategoryList.Add(category);
                            parentCategoryId = category.Id;
                        }
                        else if (parent != null)
                        {
                            parentCategoryId = parent.Id;
                        }


                        int? rez = allCategoryList.FirstOrDefault(x => x.Name.Equals(categoryName) && x.ParentCategoryId == parentCategoryId)?.Id;

                        //categories does not exist. Insert new category
                        if (rez == null)
                        {
                            var category = CreateCategoryAndSlug(categoryName, parentCategoryId);

                            rez = category.Id;
                            allCategoryList.Add(category);
                        }

                        //insert product category mapping
                        var productCategory = new ProductCategory
                        {
                            ProductId = product.Id,
                            CategoryId = rez.Value,
                            IsFeaturedProduct = false,
                            DisplayOrder = 1
                        };
                        _categoryService.InsertProductCategory(productCategory);

                    }

                    //manufacturer
                    tempProperty = metadata.Manager.GetProperty("Producer Name");
                    if (tempProperty != null && !string.IsNullOrWhiteSpace(tempProperty.StringValue))
                    {
                        var manufacturerName = tempProperty.StringValue.Trim();
                        //manufacturer mappings

                        var manufacturer = allManufacturers.FirstOrDefault(x => x.Name == manufacturerName);
                        //manufacturer does not exist. Insert the new manufacturer
                        if (manufacturer == null)
                        {
                            manufacturer = new Manufacturer
                            {
                                Name = manufacturerName,
                                Published = true,
                                DisplayOrder = 0,
                                AllowCustomersToSelectPageSize = true,
                                PageSizeOptions = "6,3,9"
                            };
                            //save manufacturer and search engine name
                            _manufacturerService.InsertManufacturer(manufacturer);
                            string manufacturerSeName = _urlRecordService.ValidateSeName(manufacturer, "", manufacturer.Name, true);
                            _urlRecordService.SaveSlug(manufacturer, manufacturerSeName, 0);

                            allManufacturers.Add(manufacturer);
                        }

                        //Insert product manufacturer mapping
                        var productManufacturer = new ProductManufacturer
                        {
                            ProductId = product.Id,
                            ManufacturerId = manufacturer.Id,
                            IsFeaturedProduct = false,
                            DisplayOrder = 1
                        };
                        _manufacturerService.InsertProductManufacturer(productManufacturer);
                    }

                    //Tier prices 
                    foreach (KeyValuePair<string, string> b2BRole in tierPriceDictionary)
                    {
                        tempProperty = metadata.Manager.GetProperty(b2BRole.Value);
                        var price = tempProperty.DecimalValueNullable;

                        if (!price.HasValue || price.Value <= 0)
                            continue;

                        _productService.InsertTierPrice(new TierPrice
                        {
                            CustomerRole = allB2BRoles.First(x => x.SystemName.Equals(b2BRole.Key)),
                            Price = price.Value,
                            ProductId = product.Id,
                            Quantity = 0
                        });
                    }

                    //Specification attribute
                    foreach (var specAttribute in specificationAttributesInImport)
                    {
                        tempProperty = metadata.Manager.GetProperty(specAttribute);
                        string specAttributeValue = tempProperty.StringValue.Trim();
                        if (specAttributeValue.IsNullOrWhiteSpace())
                            continue;

                        _specificationAttributeService.InsertProductSpecificationAttribute(new ProductSpecificationAttribute
                        {
                            ProductId = product.Id,
                            AttributeType = SpecificationAttributeType.CustomText,
                            SpecificationAttributeOption = existingSpecificationAttributes.
                                 First(x => x.Name.Contains(specAttribute)).SpecificationAttributeOptions.First(),
                            CustomValue = specAttributeValue,
                            AllowFiltering = false,
                            ShowOnProductPage = true,
                            DisplayOrder = 0,
                        });
                    }

                    //todo import with picture
                    //var picture1 = DownloadFile(metadata.Manager.GetProperty("Picture1")?.StringValue, downloadedFiles);
                    //var picture2 = DownloadFile(metadata.Manager.GetProperty("Picture2")?.StringValue, downloadedFiles);
                    //var picture3 = DownloadFile(metadata.Manager.GetProperty("Picture3")?.StringValue, downloadedFiles);

                    //productPictureMetadata.Add(new ProductPictureMetadata
                    //{
                    //    ProductItem = product,
                    //    Picture1Path = picture1,
                    //    Picture2Path = picture2,
                    //    Picture3Path = picture3,
                    //    IsNew = true
                    //});

                    lastLoadedProduct = product;

                    //update "HasTierPrices" and "HasDiscountsApplied" properties
                    //_productService.UpdateHasTierPricesProperty(product);
                    //_productService.UpdateHasDiscountsApplied(product);
                }

                //if (_mediaSettings.ImportProductImagesUsingHash && _pictureService.StoreInDb && _dataProvider.SupportedLengthOfBinaryHash > 0)
                //    ImportProductImagesUsingHash(productPictureMetadata, allProductsBySku);
                //else
                //    ImportProductImagesUsingServices(productPictureMetadata);


                //activity log
                _customerActivityService.InsertActivity("ImportProducts", string.Format(_localizationService.GetResource("ActivityLog.ImportProducts"), metadata.CountProductsInFile));
            }
        }

        private ImportProductMetadata PrepareImportProductDataForMigration(ExcelWorksheet worksheet)
        {
            //the columns
            var properties = GetPropertiesByExcelCells<Product>(worksheet);

            var manager = new PropertyManager<Product>(properties, _catalogSettings);

            var productAttributeProperties = new[]
            {
                new PropertyByName<ExportProductAttribute>("AttributeId"),
                new PropertyByName<ExportProductAttribute>("AttributeName"),
                new PropertyByName<ExportProductAttribute>("AttributeTextPrompt"),
                new PropertyByName<ExportProductAttribute>("AttributeIsRequired"),
                new PropertyByName<ExportProductAttribute>("AttributeControlType"),
                new PropertyByName<ExportProductAttribute>("AttributeDisplayOrder"),
                new PropertyByName<ExportProductAttribute>("ProductAttributeValueId"),
                new PropertyByName<ExportProductAttribute>("ValueName"),
                new PropertyByName<ExportProductAttribute>("AttributeValueType"),
                new PropertyByName<ExportProductAttribute>("AssociatedProductId"),
                new PropertyByName<ExportProductAttribute>("ColorSquaresRgb"),
                new PropertyByName<ExportProductAttribute>("ImageSquaresPictureId"),
                new PropertyByName<ExportProductAttribute>("PriceAdjustment"),
                new PropertyByName<ExportProductAttribute>("PriceAdjustmentUsePercentage"),
                new PropertyByName<ExportProductAttribute>("WeightAdjustment"),
                new PropertyByName<ExportProductAttribute>("Cost"),
                new PropertyByName<ExportProductAttribute>("CustomerEntersQty"),
                new PropertyByName<ExportProductAttribute>("Quantity"),
                new PropertyByName<ExportProductAttribute>("IsPreSelected"),
                new PropertyByName<ExportProductAttribute>("DisplayOrder"),
                new PropertyByName<ExportProductAttribute>("PictureId")
            };

            var productAttributeManager = new PropertyManager<ExportProductAttribute>(productAttributeProperties, _catalogSettings);

            var specificationAttributeProperties = new[]
            {
                new PropertyByName<ExportSpecificationAttribute>("AttributeType", p => p.AttributeTypeId),
                new PropertyByName<ExportSpecificationAttribute>("SpecificationAttribute", p => p.SpecificationAttributeId),
                new PropertyByName<ExportSpecificationAttribute>("CustomValue", p => p.CustomValue),
                new PropertyByName<ExportSpecificationAttribute>("SpecificationAttributeOptionId", p => p.SpecificationAttributeOptionId),
                new PropertyByName<ExportSpecificationAttribute>("AllowFiltering", p => p.AllowFiltering),
                new PropertyByName<ExportSpecificationAttribute>("ShowOnProductPage", p => p.ShowOnProductPage),
                new PropertyByName<ExportSpecificationAttribute>("DisplayOrder", p => p.DisplayOrder)
            };

            var specificationAttributeManager = new PropertyManager<ExportSpecificationAttribute>(specificationAttributeProperties, _catalogSettings);

            var endRow = 2;
            var allCategories = new List<string>();
            var allSku = new List<string>();

            //category
            var tempProperty = manager.GetProperty("Brand");
            var categoryCellNum = tempProperty?.PropertyOrderPosition ?? -1;

            //sku
            tempProperty = manager.GetProperty("ArtikelID");
            var skuCellNum = tempProperty?.PropertyOrderPosition ?? -1;

            //manufacturer
            var allManufacturers = new List<string>();
            tempProperty = manager.GetProperty("Producer Name");
            var manufacturerCellNum = tempProperty?.PropertyOrderPosition ?? -1;

            var allStores = new List<string>();
            tempProperty = manager.GetProperty("LimitedToStores");
            var limitedToStoresCellNum = tempProperty?.PropertyOrderPosition ?? -1;

            var allAttributeIds = new List<int>();
            var allSpecificationAttributeOptionIds = new List<int>();

            var attributeIdCellNum = 1 + ExportProductAttribute.ProducAttributeCellOffset;
            var specificationAttributeOptionIdCellNum =
                specificationAttributeManager.GetIndex("SpecificationAttributeOptionId") +
                ExportProductAttribute.ProducAttributeCellOffset;

            var productsInFile = new List<int>();

            //find end of data
            var typeOfExportedAttribute = ExportedAttributeType.NotSpecified;
            while (true)
            {
                var allColumnsAreEmpty = manager.GetProperties
                    .Select(property => worksheet.Cells[endRow, property.PropertyOrderPosition])
                    .All(cell => string.IsNullOrEmpty(cell?.Value?.ToString()));

                if (allColumnsAreEmpty)
                    break;

                if (new[] { 1, 2 }.Select(cellNum => worksheet.Cells[endRow, cellNum])
                        .All(cell => string.IsNullOrEmpty(cell?.Value?.ToString())) &&
                    worksheet.Row(endRow).OutlineLevel == 0)
                {
                    var cellValue = worksheet.Cells[endRow, attributeIdCellNum].Value;
                    SetOutLineForProductAttributeRow(cellValue, worksheet, endRow);
                    SetOutLineForSpecificationAttributeRow(cellValue, worksheet, endRow);
                }

                if (worksheet.Row(endRow).OutlineLevel != 0)
                {
                    var newTypeOfExportedAttribute = GetTypeOfExportedAttribute(worksheet, productAttributeManager, specificationAttributeManager, endRow);

                    //skip caption row
                    if (newTypeOfExportedAttribute != ExportedAttributeType.NotSpecified && newTypeOfExportedAttribute != typeOfExportedAttribute)
                    {
                        typeOfExportedAttribute = newTypeOfExportedAttribute;
                        endRow++;
                        continue;
                    }

                    switch (typeOfExportedAttribute)
                    {
                        case ExportedAttributeType.ProductAttribute:
                            productAttributeManager.ReadFromXlsx(worksheet, endRow,
                                ExportProductAttribute.ProducAttributeCellOffset);
                            if (int.TryParse((worksheet.Cells[endRow, attributeIdCellNum].Value ?? string.Empty).ToString(), out var aid))
                            {
                                allAttributeIds.Add(aid);
                            }

                            break;
                        case ExportedAttributeType.SpecificationAttribute:
                            specificationAttributeManager.ReadFromXlsx(worksheet, endRow, ExportProductAttribute.ProducAttributeCellOffset);

                            if (int.TryParse((worksheet.Cells[endRow, specificationAttributeOptionIdCellNum].Value ?? string.Empty).ToString(), out var saoid))
                            {
                                allSpecificationAttributeOptionIds.Add(saoid);
                            }

                            break;
                    }

                    endRow++;
                    continue;
                }

                if (categoryCellNum > 0)
                {
                    var categoryIds = worksheet.Cells[endRow, categoryCellNum].Value?.ToString() ?? string.Empty;

                    if (!string.IsNullOrEmpty(categoryIds))
                        allCategories.AddRange(categoryIds
                            .Split(new[] { ";", ">>" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim())
                            .Distinct());
                }

                if (skuCellNum > 0)
                {
                    var sku = worksheet.Cells[endRow, skuCellNum].Value?.ToString() ?? string.Empty;

                    if (!string.IsNullOrEmpty(sku))
                        allSku.Add(sku);
                }

                if (manufacturerCellNum > 0)
                {
                    var manufacturerIds = worksheet.Cells[endRow, manufacturerCellNum].Value?.ToString() ??
                                          string.Empty;
                    if (!string.IsNullOrEmpty(manufacturerIds))
                        allManufacturers.AddRange(manufacturerIds
                            .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()));
                }

                if (limitedToStoresCellNum > 0)
                {
                    var storeIds = worksheet.Cells[endRow, limitedToStoresCellNum].Value?.ToString() ??
                                          string.Empty;
                    if (!string.IsNullOrEmpty(storeIds))
                        allStores.AddRange(storeIds
                            .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()));
                }

                //counting the number of products
                productsInFile.Add(endRow);

                endRow++;
            }

            return new ImportProductMetadata
            {
                EndRow = endRow,
                Manager = manager,
                Properties = properties,
                ProductsInFile = productsInFile,
                ProductAttributeManager = productAttributeManager,
                SpecificationAttributeManager = specificationAttributeManager,
                SkuCellNum = skuCellNum,
                AllSku = allSku
            };
        }

        private void MigrateProductsFromSplitedXlsx(ExcelWorksheet worksheet, ImportProductMetadata metadata)
        {
            foreach (var path in SplitProductFile(worksheet, metadata))
            {
                using (var scope = _serviceScopeFactory.CreateScope())
                {
                    // Resolve
                    var importManager = scope.ServiceProvider.GetRequiredService<IImportManager>();

                    using (var sr = new StreamReader(path))
                    {
                        importManager.MigrateProductsFromXlsx(sr.BaseStream);
                    }
                }

                try
                {
                    _fileProvider.DeleteFile(path);
                }
                catch
                {
                    // ignored
                }
            }
        }

    }
}
