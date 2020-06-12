using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Quantum.Core.Data;
using Quantum.Core.Domain;
using Quantum.Service;
using Quantum.Service.bsViewModels;
using Quantum.Web.Framework.Controllers;
using Quantum.Web.Framework.Paging;
using Quantum.Core.Utility;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using Quantum.Service.ViewModels;
using Quantum.Core.Configuration;
using System.Web.Http;
using System.Data;
using System.IO;
using Quantum.Web.Framework.Document;
using System.Drawing;

namespace Quantum.Web.Controllers
{


    [Authorize]
    public class ReportController : BaseApiController
    {
        protected IbsReportService ReportService;
        protected IConfiguration _iconfiguration;
        protected IUnitOfWorkManager UnitOfWorkManager;

        public ReportController(
            IbsReportService reportService,
            IConfiguration iconfiguration,
            IUnitOfWorkManager unitOfWorkManager)
        {
            ReportService = reportService;
            _iconfiguration = iconfiguration;
            UnitOfWorkManager = unitOfWorkManager;
        }



        [HttpPost]
        public IActionResult GetWorkOrderList ([FromBody] WorkOrderViewModel model)
        {
            if(model.ContractNumber == null || model.ContractNumber == "")
            {
                model.ContractNumber = "-1";
            }
            if(model.ProductionVendor == null)
            {
                model.ProductionVendor = -1;
            }
            if (model.VendorRep == null)
            {
                model.VendorRep = -1;
            }
            model.SellerID = CurrentUserDetails.SellerID;


            var workOrderList = ReportService.GetWorkOrder(model);
            return Ok(workOrderList);
        }

        [HttpPost]
        public IActionResult GetShippingOrderList([FromBody] ShippingOrderViewModel model)
        {
            if (model.ContractNumber == null || model.ContractNumber == "")
            {
                model.ContractNumber = "-1";
            }
            if (model.ProductionVendor == null)
            {
                model.ProductionVendor = -1;
            }
            if (model.VendorRep == null)
            {
                model.VendorRep = -1;
            }
            if (model.Campaign == null || model.Campaign == "")
            {
                model.Campaign = "";
            }
            model.SellerID = CurrentUserDetails.SellerID;


            var shippingOrderList = ReportService.GetShippingOrder(model);
            return Ok(shippingOrderList);
        }

        [HttpPost]
        public IActionResult GetProductionReportList([FromBody] ProductionReportViewModel model)
        {
            if (model.ContractNumber == null || model.ContractNumber == "")
            {
                model.ContractNumber = "-1";
            }
            if (model.ProductionVendor == null)
            {
                model.ProductionVendor = -1;
            }
            if (model.CampaignName == null || model.CampaignName == "")
            {
                model.CampaignName = "";
            }
            if (model.DMA == null || model.DMA == 211)
            {
                model.DMA = -1;
            }
            model.SellerID = CurrentUserDetails.SellerID;


            var productionReportList = ReportService.GetProductionReport(model);
            return Ok(productionReportList);
        }
        [HttpGet]
        public async Task<ActionResult> getReportHeader(string ContractNumber, string ReportFor)
        {
            var result = await ReportService.getReportHeaderForWorkOrder(ContractNumber, ReportFor, CurrentUserDetails.SellerID);
            return Ok(result);

        }

        [HttpPost]
        public async Task<ActionResult> getMappedOwnerProductionVendor([FromBody] WorkOrderViewModel model)
        {
            if (model.ContractNumber == null || model.ContractNumber == "")
            {
                model.ContractNumber = "-1";
            }
            if (model.ProductionVendor == null)
            {
                model.ProductionVendor = -1;
            }
            if (model.VendorRep == null)
            {
                model.VendorRep = -1;
            }
            var result = await ReportService.getMappedOwnerProductionVendor(model, CurrentUserDetails.SellerID);
            return Ok(result);

        }

        [HttpGet]
        public IActionResult getProductionVendorForShippingAndWorkOrder(string ContactID)
        {
            var result = ReportService.getProductionVendorForShippingAndWorkOrder(ContactID);
            return Ok(result);

        }

        [HttpPost]
        public IActionResult ProductionRequirementExcel([FromBody] ProductionReportViewModel model)
        {
            if (model.ContractNumber == null || model.ContractNumber == "")
            {
                model.ContractNumber = "-1";
            }
            if (model.ProductionVendor == null)
            {
                model.ProductionVendor = -1;
            }
            if (model.CampaignName == null || model.CampaignName == "")
            {
                model.CampaignName = "";
            }
            if (model.DMA == null || model.DMA == 211)
            {
                model.DMA = -1;
            }
            model.SellerID = CurrentUserDetails.SellerID;
            var uploadPath = string.Format("uploads\\" + CurrentUserDetails.SellerID + "\\Excel\\");
            var targetDirectory = Path.Combine(_iconfiguration["BManageFolder"], uploadPath);
            if (!Directory.Exists(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
            }

            var filename = "ProductionRequirement" + DateTime.Now.ToShortDateString().Replace('/', '_') + ".xlsx";
            var savePath = Path.Combine(targetDirectory, filename);
            var loadURL = _iconfiguration["BManageUrl"] + uploadPath + filename;
            DataTable prodRequireList = ReportService.ProductionRequirementExport(model);
            DataSet dsExportReport = new DataSet();
            dsExportReport.Tables.Add(prodRequireList);

            ExcelHelper.ListToExcelInMultipleSheets(dsExportReport, savePath);

            return Ok(loadURL);
        }

        [HttpPost]
        public IActionResult shippingOrdersExcel([FromBody] ShippingOrderViewModel model)
        {
            if (model.ContractNumber == null || model.ContractNumber == "")
            {
                model.ContractNumber = "-1";
            }
            if (model.ProductionVendor == null)
            {
                model.ProductionVendor = -1;
            }
            if (model.VendorRep == null)
            {
                model.VendorRep = -1;
            }
            if(model.Campaign == null || model.Campaign == "")
            {
                model.Campaign = "";
            }
            model.SellerID = CurrentUserDetails.SellerID;

            DataTable shipOrderList = ReportService.shippingRequirementExport(model);
            var uploadPath = string.Format("uploads\\" + CurrentUserDetails.SellerID + "\\Excel\\");
            var targetDirectory = Path.Combine(_iconfiguration["BManageFolder"], uploadPath);
            if (!Directory.Exists(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
            }
            var filename = "ShippingOrderExport" + DateTime.Now.ToShortDateString().Replace('/', '_') + ".xlsx";
            var savePath = Path.Combine(targetDirectory, filename);
            var loadURL = _iconfiguration["BManageUrl"] + uploadPath + filename;

            DataSet dsExportReport = new DataSet();
            dsExportReport.Tables.Add(shipOrderList);

            ExcelHelper.ListToExcelInMultipleSheets(dsExportReport, savePath);

            return Ok(loadURL);
        }

        [HttpPost]
        public IActionResult WorkOrderExportToExcel([FromBody] WorkOrderViewModel model)
        {
            if (model.ContractNumber == null || model.ContractNumber == "")
            {
                model.ContractNumber = "-1";
            }
            if (model.ProductionVendor == null)
            {
                model.ProductionVendor = -1;
            }
            if (model.VendorRep == null)
            {
                model.VendorRep = -1;
            }
            model.SellerID = CurrentUserDetails.SellerID;

            //var reportHeader = getReportHeader(model.ContractNumber);

            var uploadPath = string.Format("uploads\\" + CurrentUserDetails.SellerID + "\\Excel\\");
            var targetDirectory = Path.Combine(_iconfiguration["BManageFolder"], uploadPath);

            if (!Directory.Exists(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
            }

            var filename = "workOrderReport" + DateTime.Now.ToShortDateString().Replace('/', '_') + ".xlsx";
            var savePath = Path.Combine(targetDirectory, filename);
            var loadURL = _iconfiguration["BManageUrl"] + uploadPath + filename;
            DataTable MissingSingedDocumentList = ReportService.ExportWorkOrderExcel(model);
            DataTable dtt = new DataTable();

            dtt.Columns.Add(" ");
            dtt.Columns.Add("  ");
            dtt.Columns.Add("   ");
            dtt.Columns.Add("    ");
            dtt.Columns.Add("     ");
            dtt.Columns.Add("      ");
            dtt.Columns.Add("       ");
            dtt.Columns.Add("        ");
            dtt.Columns.Add("         ");

            //dtt.Rows.Add("");
            //dtt.Rows.Add("");

            dtt.Rows.Add("");
            dtt.Rows.Add("");
            dtt.Rows.Add("");
            dtt.Rows.Add("");
            dtt.Rows.Add("");
            dtt.Rows.Add("");

            if (model.OrderType == "Posting")
            { 
            dtt.Rows.Add("POSTING ORDER : INSTRUCTIONS & CONFIRMATION OF RECEIPT OF MATERIALS");
            }
            if (model.OrderType == "TakeDown")
            {
                dtt.Rows.Add("TakeDown ORDER : INSTRUCTIONS & CONFIRMATION OF RECEIPT OF MATERIALS");
            }

            //dtt.Columns.Add(" ");
            //dtt.Columns.Add("  ");
            //dtt.Columns.Add("   ");
            //dtt.Columns.Add("    ");
            //dtt.Columns.Add("     ");
            //dtt.Columns.Add("      ");
            //dtt.Columns.Add("       ");
            //dtt.Columns.Add("        ");
            dtt.Rows.Add("Date : " + DateTime.Now.ToShortDateString());
            dtt.Rows.Add("From : " + model.Name);
            dtt.Rows.Add("Email : " + model.Email);
            dtt.Rows.Add("To : " + model.productionVendorOne + "  " + model.productionVendorTwo);
            dtt.Rows.Add("RE : " + model.JobName);
            dtt.Rows.Add(" ");
            dtt.Rows.Add("Market",/*,"Posting Date"*/model.OrderType + " Date","Unit #", "Media Owner", "Media Type", "Size", "Design", "Date Materials Received");
            DataRow[] rows = MissingSingedDocumentList.Select();
            foreach (DataRow dr in rows)
            {
                dtt.Rows.Add(dr.ItemArray);
            }
            dtt.Rows.Add("");
            dtt.Rows.Add("");
            dtt.Rows.Add("");
            dtt.Rows.Add("");
            dtt.Rows.Add("", " ", "                   ", "ADD CREATIVE IMAGE HERE", "                   ");
            dtt.Rows.Add("","","                   ", "                   ", "                   ");
            dtt.Rows.Add("", "", "                   ", "                   ", "                   ");
            dtt.Rows.Add("");
            dtt.Rows.Add("", " ", "STANDARD PHOTO REQUEST:","BorderTop","BorderTopRight");
            dtt.Rows.Add("", " ", "POSTER: 10% MUST HAVE APPROACH AND CLOSE-UP SHOTS - PHOTO OF EACH CREATIVE","","BorderRight");
            dtt.Rows.Add("   ", " ", "                                      ", "", "BorderRight");
            dtt.Rows.Add("   ", " ", "                                      ", "", "BorderRight");
            dtt.Rows.Add("", " ", "BULLETINS: 100% CLOSE-UP AND APPROACH OF EACH UNIT","BorderBottom","BorderBottomRight");
            dtt.Rows.Add("");
            dtt.Rows.Add("");
            if (model.OrderType == "Posting")
            {
                dtt.Rows.Add("PLEASE SIGN AND EMAIL BACK");
                dtt.Rows.Add("TO CONFIRM THAT THE OUTDOOR CO. ACKNOWLEDGES THE PROPER DESIGN TO BE POSTED ON THE DATE SHOWN ABOVE AND ALSO THE CREATIVE REMOVAL STATUS.");
                dtt.Rows.Add("FILL IN DATE MATERIALS ARE DELIVERED.","IF MATERIALS ARE NOT DELIVERED 5 DAYS PRIOR TO POSTING ,","","","CONTACT US IMMEDIATELY.");
                dtt.Rows.Add("Do not post prior to posting date without confirmation that it is okay to do so.","","", "If you have any questions, please feel free to contact us.");
            }
            if (model.OrderType == "TakeDown")
            {
                dtt.Rows.Add("PLEASE SIGN AND EMAIL BACK");
                dtt.Rows.Add("TO CONFIRM THAT THE OUTDOOR CO. ACKNOWLEDGES THE PROPER DESIGN  TO BE REMOVED AFTER THE TAKEDOWN DATE SHOWN ABOVE AND ");
                dtt.Rows.Add("ALSO TO PROVIDE THE ACTUAL TAKEDOWN DATE.","IF MATERIALS NEED TO BE REMOVED PRIOR TO TAKEDOWN DATE SHOWN ABOVE ,","","","CONTACT US IMMEDIATELY.");
                dtt.Rows.Add("Do not remove prior to takedown date without confirmation that it is okay to do so.","","", "If you have any questions, please feel free to contact us.");
            }
            dtt.Rows.Add(" ");
            dtt.Rows.Add("", " ", "Authorized Signature");
            dtt.Rows.Add("", " ", "Date");
            DataSet ds = new DataSet();
            dtt.TableName = "Work Order Report";
            ds.Tables.Add(dtt);


            ExcelHelper.ExportToExcelWithImageHeader(ds, savePath, model.ImagePath);
            

            return Ok(loadURL);
        }

        [HttpPost]
        public  IActionResult GetContractBillingList([FromBody] ContractBillingExportToExcelSearchViewModel model)
        {
            model.SellerID = CurrentUserDetails.SellerID;
            var contractBillingList =   ReportService.DatatableContractBilling(model);
            return Ok(contractBillingList);
        }

        [HttpPost]
        public async Task<IActionResult> getRevisionReportHeader([FromBody] rivisionHeaderViewModel model)
        {
            var uploadPath = string.Format("uploads\\" + CurrentUserDetails.SellerID + "\\Excel\\");
            var targetDirectory = Path.Combine(_iconfiguration["BManageFolder"], uploadPath);

            if (!Directory.Exists(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
            }

            var filename = "RevisionReport" + DateTime.Now.ToShortDateString().Replace('/', '_') + ".xlsx";
            var savePath = Path.Combine(targetDirectory, filename);
            var loadURL = _iconfiguration["BManageUrl"] + uploadPath + filename;
            var headerList = ReportService.getHeaderForReport(model.ContractNumber, CurrentUserDetails.SellerID);

            DataTable dtt = new DataTable();
            dtt.Columns.Add(" ");
            dtt.Columns.Add("  ");
            dtt.Columns.Add("   ");
            dtt.Columns.Add("    ");
            dtt.Columns.Add("     ");
            dtt.Columns.Add("      ");
            dtt.Columns.Add("       ");
            dtt.Columns.Add("        ");

            dtt.Rows.Add("Revision Report");
            dtt.Rows.Add("Contract Number", headerList.ContractNumber);
            dtt.Rows.Add("Campaign Name", headerList.CampaignName);
            dtt.Rows.Add("Client Company Code", headerList.ClientCompanyCode);
            dtt.Rows.Add("Demographic", headerList.Demographics);
            dtt.Rows.Add("Salesperson on account", headerList.SalesPerson);
            dtt.Rows.Add("");
            dtt.Rows.Add("");

            DataTable getContractList = ReportService.getAllContractRevision(model.ContractNumber);
            DataRow[] BoxRows = getContractList.Select();
            foreach (DataRow drr in BoxRows)
            { 
            var subHeaderList = ReportService.getSubHeaderList(drr.ItemArray[0].ToString(), CurrentUserDetails.SellerID);
            DataTable RevisionReportList = ReportService.getRevisionReportList(drr.ItemArray[0].ToString(), CurrentUserDetails.SellerID);
            dtt.Rows.Add("Contract#: " + subHeaderList.ContractNumber, "Last saved date- " + drr.ItemArray[2].ToString(), "BorderTop", "BorderTop", "BorderTop", "BorderTop", "BorderTop", "TopRightBorder");
            dtt.Rows.Add("", "", "", "", "", "", "", "BorderRight");
            dtt.Rows.Add("Client Total: " , "$" + String.Format("{0:0.00}", subHeaderList.ClientTotals),"","","","", "", "BorderRight");
            dtt.Rows.Add("Vendor Total: " , "$" + String.Format("{0:0.00}", subHeaderList.VendorTotals), "", "", "", "", "", "BorderRight");
            dtt.Rows.Add("Profit Percentage : ", String.Format("{0:0.00}", subHeaderList.ProfitPercent) + "%", "", "", "", "", "", "BorderRight");
            dtt.Rows.Add("","", "", "", "", "", "", "BorderRight");
            dtt.Rows.Add("Client Total Display: " , "$" + String.Format("{0:0.00}", subHeaderList.ClientTotalDisplay), "", "", "", "", "", "BorderRight");
            dtt.Rows.Add("Client Total Production: " , "$" + String.Format("{0:0.00}", subHeaderList.ClientTotalProduction), "", "", "", "", "", "BorderRight");
            dtt.Rows.Add("","", "", "", "", "", "", "BorderRight");
            dtt.Rows.Add("Revision Notes: ", drr.ItemArray[1].ToString(), "", "", "", "", "", "BorderRight");
            dtt.Rows.Add("","", "", "", "", "", "", "BorderRight");
            dtt.Rows.Add("","", "", "", "", "", "", "BorderRight");
            dtt.Rows.Add("","Invoice","Market","Date", "Reference number","Campaign dates", "Amount due","Type");
            DataRow[] rows = RevisionReportList.Select();
            foreach (DataRow dr in rows)
            {
                dtt.Rows.Add("",dr.ItemArray[0], dr.ItemArray[1], dr.ItemArray[2], dr.ItemArray[3], dr.ItemArray[4], "$" + String.Format("{0:0.00}", dr.ItemArray[5]), dr.ItemArray[6]);
            }
            dtt.Rows.Add("BorderBottom", "BorderBottom", "BorderBottom", "BorderBottom", "BorderBottom", "BorderBottom", "BorderBottom", "BottomRightBorder");
            dtt.Rows.Add("");
            }
            DataSet ds = new DataSet();
            dtt.TableName = "Revision Report";
            ds.Tables.Add(dtt);

            ExcelHelper.AppendMultipleSheetWhileDoingExportToExcel(ds, savePath);

            return Ok(loadURL);

        }
    }
}
