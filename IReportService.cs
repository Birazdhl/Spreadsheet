using Quantum.Core.Service;
using Quantum.Service.bsViewModels;
using Quantum.Service.ViewModels;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace Quantum.Service
{
    //Please consult with the team lead before adding/modifying to this service
    public interface IbsReportService : IService
    {
        List<WorkOrderViewModel> GetWorkOrder(WorkOrderViewModel model);

        List<ShippingOrderViewModel> GetShippingOrder(ShippingOrderViewModel model);

        List<ProductionReportViewModel> GetProductionReport(ProductionReportViewModel model);
        Task<IList<ReportHeaderViewModel>> getReportHeaderForWorkOrder(string ContractNumber,string ReportFor, int SellerID);
        string getProductionVendorForShippingAndWorkOrder(string ContactID);
        DataTable ProductionRequirementExport(ProductionReportViewModel model);
        DataTable shippingRequirementExport(ShippingOrderViewModel model);
        DataTable ExportWorkOrderExcel(WorkOrderViewModel model);
        Task<IList<ReportHeaderViewModel>> getMappedOwnerProductionVendor(WorkOrderViewModel model, int SellerID);
        DataTable DatatableContractBilling(ContractBillingExportToExcelSearchViewModel model);
        DataTable getRevisionReportList(string contractNumber, int sellerId);
        rivisionHeaderViewModel getHeaderForReport(string contractNumber, int sellerId);
        rivisionSubHeaderViewModel getSubHeaderList(string contractNumber, int SellerID);
        DataTable getAllContractRevision(string contractNumber);
    }
}