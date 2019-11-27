using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Nop.Core.Domain.Catalog;
using Nop.Core.Domain.Customers;
using Nop.Services.Security;
using Nop.Web.Areas.Admin.Infrastructure.Mapper.Extensions;
using Nop.Web.Areas.Admin.Models.Catalog;
using Nop.Web.Framework;
using Nop.Web.Framework.Controllers;
using Nop.Web.Framework.Mvc;
using Nop.Web.Framework.Mvc.Filters;

namespace Nop.Web.Areas.Admin.Controllers
{
    public partial class ProductController
    {

        [HttpPost, ParameterBasedOnFormName("save-continue", "continueEditing")]
        public virtual IActionResult Create(ProductModel model, bool continueEditing)
        {
            if (!_permissionService.Authorize(StandardPermissionProvider.ManageProducts))
                return AccessDeniedView();

            //validate maximum number of products per vendor
            if (_vendorSettings.MaximumProductNumber > 0 && _workContext.CurrentVendor != null
                && _productService.GetNumberOfProductsByVendorId(_workContext.CurrentVendor.Id) >= _vendorSettings.MaximumProductNumber)
            {
                _notificationService.ErrorNotification(string.Format(_localizationService.GetResource("Admin.Catalog.Products.ExceededMaximumNumber"),
                    _vendorSettings.MaximumProductNumber));
                return RedirectToAction("List");
            }

            if (ModelState.IsValid)
            {
                //a vendor should have access only to his products
                if (_workContext.CurrentVendor != null)
                    model.VendorId = _workContext.CurrentVendor.Id;

                //vendors cannot edit "Show on home page" property
                if (_workContext.CurrentVendor != null && model.ShowOnHomepage)
                    model.ShowOnHomepage = false;

                //product
                var product = model.ToEntity<Product>();
                product.CreatedOnUtc = DateTime.UtcNow;
                product.UpdatedOnUtc = DateTime.UtcNow;
                //extended: product will definitely have tier price
                product.HasTierPrices = true;

                _productService.InsertProduct(product);

                //extended: tier prices for B2B
               IEnumerable<CustomerRole> allB2BRoles= _customerService.GetAllCustomerRoles()
                   .Where(x=>x.SystemName.Contains(B2BRole.RoleSuffix));

               foreach (CustomerRole b2BRole in allB2BRoles)
               {
                   _productService.InsertTierPrice(new TierPrice
                   {
                       CustomerRoleId = b2BRole.Id,
                       Price = product.Price,
                       ProductId = product.Id,
                       Quantity = 0
                   });
               }

                //search engine name
                model.SeName = _urlRecordService.ValidateSeName(product, model.SeName, product.Name, true);
                _urlRecordService.SaveSlug(product, model.SeName, 0);

                //locales
                UpdateLocales(product, model);

                //categories
                SaveCategoryMappings(product, model);

                //manufacturers
                SaveManufacturerMappings(product, model);

                //ACL (customer roles)
                SaveProductAcl(product, model);

                //stores
                _productService.UpdateProductStoreMappings(product, model.SelectedStoreIds);

                //discounts
                SaveDiscountMappings(product, model);

                //tags
                _productTagService.UpdateProductTags(product, ParseProductTags(model.ProductTags));

                //warehouses
                SaveProductWarehouseInventory(product, model);

                //quantity change history
                _productService.AddStockQuantityHistoryEntry(product, product.StockQuantity, product.StockQuantity, product.WarehouseId,
                    _localizationService.GetResource("Admin.StockQuantityHistory.Messages.Edit"));

                //activity log
                _customerActivityService.InsertActivity("AddNewProduct",
                    string.Format(_localizationService.GetResource("ActivityLog.AddNewProduct"), product.Name), product);

                _notificationService.SuccessNotification(_localizationService.GetResource("Admin.Catalog.Products.Added"));

                if (!continueEditing)
                    return RedirectToAction("List");

                return RedirectToAction("Edit", new { id = product.Id });
            }

            //prepare model
            model = _productModelFactory.PrepareProductModel(model, null, true);

            //if we got this far, something failed, redisplay form
            return View(model);
        }

        #region Tier prices

        [HttpPost]
        public virtual IActionResult TierPriceList(TierPriceSearchModel searchModel)
        {
            if (!_permissionService.Authorize(StandardPermissionProvider.ManageProducts))
                return AccessDeniedDataTablesJson();

            //try to get a product with the specified id
            var product = _productService.GetProductById(searchModel.ProductId)
                ?? throw new ArgumentException("No product found with the specified id");

            //a vendor should have access only to his products
            if (_workContext.CurrentVendor != null && product.VendorId != _workContext.CurrentVendor.Id)
                return Content("This is not your product");

            //prepare model
            var model = _productModelFactory.PrepareTierPriceListModel(searchModel, product);

            return Json(model);
        }

        public virtual IActionResult TierPriceCreatePopup(int productId)
        {
            if (!_permissionService.Authorize(StandardPermissionProvider.ManageProducts))
                return AccessDeniedView();

            //try to get a product with the specified id
            var product = _productService.GetProductById(productId)
                ?? throw new ArgumentException("No product found with the specified id");

            //prepare model
            var model = _productModelFactory.PrepareTierPriceModel(new TierPriceModel(), product, null);

            return View(model);
        }

        [HttpPost]
        [FormValueRequired("save")]
        public virtual IActionResult TierPriceCreatePopup(TierPriceModel model)
        {
            if (!_permissionService.Authorize(StandardPermissionProvider.ManageProducts))
                return AccessDeniedView();

            //try to get a product with the specified id
            var product = _productService.GetProductById(model.ProductId)
                ?? throw new ArgumentException("No product found with the specified id");

            //a vendor should have access only to his products
            if (_workContext.CurrentVendor != null && product.VendorId != _workContext.CurrentVendor.Id)
                return RedirectToAction("List", "Product");

            if (ModelState.IsValid)
            {
                //fill entity from model
                var tierPrice = model.ToEntity<TierPrice>();
                tierPrice.ProductId = product.Id;
                tierPrice.CustomerRoleId = model.CustomerRoleId > 0 ? model.CustomerRoleId : (int?)null;

                _productService.InsertTierPrice(tierPrice);

                //update "HasTierPrices" property
                _productService.UpdateHasTierPricesProperty(product);

                ViewBag.RefreshPage = true;

                return View(model);
            }

            //prepare model
            model = _productModelFactory.PrepareTierPriceModel(model, product, null, true);

            //if we got this far, something failed, redisplay form
            return View(model);
        }

        public virtual IActionResult TierPriceEditPopup(int id)
        {
            if (!_permissionService.Authorize(StandardPermissionProvider.ManageProducts))
                return AccessDeniedView();

            //try to get a tier price with the specified id
            var tierPrice = _productService.GetTierPriceById(id);
            if (tierPrice == null)
                return RedirectToAction("List", "Product");

            //try to get a product with the specified id
            var product = _productService.GetProductById(tierPrice.ProductId)
                ?? throw new ArgumentException("No product found with the specified id");

            //a vendor should have access only to his products
            if (_workContext.CurrentVendor != null && product.VendorId != _workContext.CurrentVendor.Id)
                return RedirectToAction("List", "Product");

            //prepare model
            var model = _productModelFactory.PrepareTierPriceModel(null, product, tierPrice);

            return View(model);
        }

        [HttpPost]
        public virtual IActionResult TierPriceEditPopup(TierPriceModel model)
        {
            if (!_permissionService.Authorize(StandardPermissionProvider.ManageProducts))
                return AccessDeniedView();

            //try to get a tier price with the specified id
            var tierPrice = _productService.GetTierPriceById(model.Id);
            if (tierPrice == null)
                return RedirectToAction("List", "Product");

            //try to get a product with the specified id
            var product = _productService.GetProductById(tierPrice.ProductId)
                ?? throw new ArgumentException("No product found with the specified id");

            //a vendor should have access only to his products
            if (_workContext.CurrentVendor != null && product.VendorId != _workContext.CurrentVendor.Id)
                return RedirectToAction("List", "Product");

            if (ModelState.IsValid)
            {
                //fill entity from model
                tierPrice = model.ToEntity(tierPrice);
                tierPrice.CustomerRoleId = model.CustomerRoleId > 0 ? model.CustomerRoleId : (int?)null;
                _productService.UpdateTierPrice(tierPrice);

                ViewBag.RefreshPage = true;

                return View(model);
            }

            //prepare model
            model = _productModelFactory.PrepareTierPriceModel(model, product, tierPrice, true);

            //if we got this far, something failed, redisplay form
            return View(model);
        }

        [HttpPost]
        public virtual IActionResult TierPriceDelete(int id)
        {
            if (!_permissionService.Authorize(StandardPermissionProvider.ManageProducts))
                return AccessDeniedView();

            //try to get a tier price with the specified id
            var tierPrice = _productService.GetTierPriceById(id)
                ?? throw new ArgumentException("No tier price found with the specified id");

            //try to get a product with the specified id
            var product = _productService.GetProductById(tierPrice.ProductId)
                ?? throw new ArgumentException("No product found with the specified id");

            //a vendor should have access only to his products
            if (_workContext.CurrentVendor != null && product.VendorId != _workContext.CurrentVendor.Id)
                return Content("This is not your product");

            _productService.DeleteTierPrice(tierPrice);

            //update "HasTierPrices" property
            _productService.UpdateHasTierPricesProperty(product);

            return new NullJsonResult();
        }

        #endregion

        #region Migrate

        [HttpPost]
        public IActionResult ImportExcelForMigration(IFormFile importexcelfileexisting)
        {
            if (!_permissionService.Authorize(StandardPermissionProvider.ManageProducts))
                return AccessDeniedView();

            if (_workContext.CurrentVendor != null && !_vendorSettings.AllowVendorsToImportProducts)
                //a vendor can not import products
                return AccessDeniedView();

            try
            {
                if (importexcelfileexisting != null && importexcelfileexisting.Length > 0)
                {
                    _importManager.MigrateProductsFromXlsx(importexcelfileexisting.OpenReadStream());
                }
                else
                {
                    _notificationService.ErrorNotification(_localizationService.GetResource("Admin.Common.UploadFile"));
                    return RedirectToAction("List");
                }

                _notificationService.SuccessNotification(_localizationService.GetResource("Admin.Catalog.Products.Imported"));
                return RedirectToAction("List");
            }
            catch (Exception exc)
            {
                _notificationService.ErrorNotification(exc);
                return RedirectToAction("List");
            }
        }

        #endregion Migrate

    }
}
