using Dapper;
using Quantum.Core.Data;
using Quantum.Data;
using Quantum.Service.bsViewModels;
using Quantum.Service.ViewModels;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quantum.Service
{
    //Please consult with the team lead before adding/modifying to this service
    public class bsReportService : IbsReportService
    {
        protected IDapperRepository _dapperRepository;
        protected IUnitOfWorkManager UnitOfWorkManager;
        public bsReportService
            (IDapperRepository dapperRepository
            , IUnitOfWorkManager unitOfWorkManager)
        {
            _dapperRepository = dapperRepository;
            UnitOfWorkManager = unitOfWorkManager;
        }

        public List<WorkOrderViewModel> GetWorkOrder(WorkOrderViewModel model)
        {
            string spName = "stp_GetPostingOrderReportEquinox";
            var Params = new DynamicParameters();
            Params.Add("PeriodStart", model.FromDate);
            Params.Add("PeriodEnd", model.ToDate);
            Params.Add("Advertiser", model.ProductionVendor);
            Params.Add("AdvertiserContactID", model.VendorRep);
            Params.Add("SellerID", model.SellerID);
            Params.Add("ContractNumber", model.ContractNumber);
            Params.Add("OrderType", model.OrderType);

            var data = _dapperRepository.ExecuteStoredProc<WorkOrderViewModel>(spName, Params).ToList();
            return data;
        }

        public List<ShippingOrderViewModel> GetShippingOrder(ShippingOrderViewModel model)
        {
            string spName = "stp_GetShippingOrdersReportEquinox";
            var Params = new DynamicParameters();
            Params.Add("PeriodStart", model.FromDateRunReport);
            Params.Add("PeriodEnd", model.ToDateRunReport);
            Params.Add("Advertiser", model.ProductionVendor);
            Params.Add("AdvertiserContactID", model.VendorRep);
            Params.Add("SellerID", model.SellerID);
            Params.Add("ContractNumber", model.ContractNumber);
            Params.Add("OrderType", model.OrderType);
            Params.Add("CampaignName", model.Campaign);

            var data = _dapperRepository.ExecuteStoredProc<ShippingOrderViewModel>(spName, Params).ToList();
            return data;
        }

        public List<ProductionReportViewModel> GetProductionReport(ProductionReportViewModel model)
        {
            string spName = "stp_GetProductionRequirementsReportEquinox";
            var Params = new DynamicParameters();
            Params.Add("PeriodStart", model.FromDateRunReport);
            Params.Add("PeriodEnd", model.ToDateRunReport);
            Params.Add("Advertiser", model.ProductionVendor);
            Params.Add("SellerID", model.SellerID);
            Params.Add("ContractNumber", model.ContractNumber);
            Params.Add("CampaignName", model.CampaignName);
            Params.Add("DMA", model.DMA);

            var data = _dapperRepository.ExecuteStoredProc<ProductionReportViewModel>(spName, Params).ToList();
            return data;
        }

        public async Task<IList<ReportHeaderViewModel>> getReportHeaderForWorkOrder(string ContractNumber,string ReportFor, int SellerID)
        {
            var strSql = new StringBuilder();
            strSql.AppendFormat(@"SELECT DISTINCT COALESCE(ISNULL(FirstName,'') + ' ' + LastName,'') AS Name,Contact.EmailAddress,JobName,org1.OrganizationName AS AgencyName,org2.OrganizationName AS AdvertiserName FROM dbo.bsContractLineItemV2 
                                    LEFT JOIN dbo.bsContract ON bsContract.ContractNumber = bsContractLineItemV2.ContractNumber
                                    LEFT JOIN dbo.Organization org1 ON org1.OrganizationID = bsContract.AgencyOrgID
                                    LEFT JOIN dbo.Organization org2 ON org2.OrganizationID = bsContract.AdvertiserID
                                    LEFT JOIN Contact ON Contact.ContactID = bsContract.ACID
                                    WHERE bsContract.ContractNumber=@ContractNumber AND bsContract.SellerID=@SellerID ");


            //if (ReportFor == "Vendor" || ReportFor == "Media Owner")
            //{
            //    strSql.Append(@"AND bsContractLineItemV2.LineItemStyle='Associated'");
            //}
            //if (ReportFor == "Media Owner")
            //{
            //    strSql.Append(@"AND bsContractLineItemV2.LineItemStyle<>'Associated'");
            //}

            var parameterlist = new DynamicParameters();
            parameterlist.Add("@ContractNumber", ContractNumber);
            parameterlist.Add("@SellerID", SellerID);
            var result = await _dapperRepository.ExecuteQueryAsync<ReportHeaderViewModel>(strSql.ToString(), parameterlist);
            return result.ToList();
        }

    public async Task<IList<ReportHeaderViewModel>> getMappedOwnerProductionVendor(WorkOrderViewModel model, int SellerID)
        {
            var strSql = new StringBuilder();
            strSql.AppendFormat(@"
                                    SELECT * INTO #temp FROM (
                                       SELECT 
		                                    DISTINCT
                                                COALESCE(MediaOwnerOrg.OrganizationName,
                                                         MediaOwnerSelectedFaceOrg.OrganizationName) AS OrganizationName 
            

                                       FROM     dbo.bsContractLineItemV2 CLI

                                                INNER JOIN [bsContract] ON CLI.ContractNumber = bsContract.ContractNumber
                                                INNER  JOIN dbo.bsContractLineItemV2_AssociateCostLineItem acl ON CLI.ContractNumber = acl.ContractNumber
                                                                                                  AND acl.AssociatedMediaLineNumber = CLI.LineNumber
                                                LEFT  JOIN dbo.bsSelectedFace ON bsSelectedFace.ContractNumber = bsContract.ContractNumber
                                                                                 AND bsSelectedFace.LineNumber = CLI.LineNumber
                                                LEFT JOIN dbo.Organization AS MediaOwnerOrg ON MediaOwnerOrg.OrganizationID = CLI.VendorOrgID
                                                LEFT JOIN dbo.Organization AS MediaOwnerSelectedFaceOrg ON MediaOwnerSelectedFaceOrg.OrganizationID = bsSelectedFace.VendorOrgID
                                                LEFT JOIN ( SELECT  CLIV2.ContractNumber ,
                                                                    CLIV2.LineNumber ,
                                                                    CLIV2.CreativeName ,
                                                                    CLIV2.VendorOrgID ,
                                                                    CLIV2.VendorContactID
                                                            FROM    bsContractLineItemV2 CLIV2
                                                            WHERE   CLIV2.LineItemStyle = 'Associated'
                                                                    AND CLIV2.SellerID = @SellerID
                                                          ) AS Creative ON Creative.LineNumber = acl.AssociateLineNumber
                                                                           AND Creative.ContractNumber = CLI.ContractNumber
                                       WHERE    ( CLI.LineItemStyle = 'Media'
                                                  OR CLI.LineItemStyle = 'Open'
                                                  OR CLI.LineItemStyle = 'Face'
                                                  OR CLI.LineItemStyle = 'Associated'
                                                )
                                                AND CLI.SellerID = @SellerID
                                                AND ( CLI.ContractNumber = @ContractNumber
                                                      OR @ContractNumber = '-1'
                                                    )
                                                AND ( Creative.VendorOrgID = @Advertiser
                                                      OR @Advertiser = -1
                                                    )
                                                AND ( Creative.VendorContactID = @AdvertiserContactID
                                                      OR @AdvertiserContactID = -1
                                                    )
                                                AND ( CLI.FromDate > = @PeriodStart
                                                      OR ( bsSelectedFace.FromDate > = @PeriodStart )
                                                      OR @PeriodStart IS NULL
                                                    )
                                                AND ( CLI.ToDate <= @PeriodEnd
                                                      OR ( bsSelectedFace.ToDate <= @PeriodEnd )
                                                      OR @PeriodEnd IS NULL
                                                    )
		
					                                    ) AS temp

				
					                                    SELECT DISTINCT
                                                                         STUFF((SELECT DISTINCT ' , ' +
                                                                              OrganizationName
                                                                        FROM    #temp
                                                                                FOR XML PATH(''), TYPE
                                                                                   ).value('.', 'NVARCHAR(MAX)')
                                                                               ,1,2,'') OrganizationName
                                                                        from #temp t

					                                    DROP TABLE #temp


                                    ");
            var parameterlist = new DynamicParameters();
            parameterlist.Add("@ContractNumber", model.ContractNumber);
            parameterlist.Add("@SellerID", SellerID);
            parameterlist.Add("@Advertiser", model.ProductionVendor);
            parameterlist.Add("@AdvertiserContactID", model.VendorRep);
            parameterlist.Add("@PeriodStart", model.FromDate);
            parameterlist.Add("@PeriodEnd", model.ToDate);
            var result = await _dapperRepository.ExecuteQueryAsync<ReportHeaderViewModel>(strSql.ToString(), parameterlist);
            return result.ToList();
        }
        public string getProductionVendorForShippingAndWorkOrder(string ContactID)
        {
            var sqlStr = new StringBuilder();

            sqlStr.AppendFormat(@"
                                    SELECT FirstName + '' + LastName AS NAME FROM dbo.Contact WHERE ContactID=@ContactID
                                ");

            var parameters = new DynamicParameters();
            parameters.Add("ContactID", ContactID);
            return _dapperRepository.ExecuteQueryFirstOrDefault<string>(sqlStr.ToString(), parameters);
        }

        public DataTable ProductionRequirementExport(ProductionReportViewModel model)
        {
            if (model.Program == null)
            {
                model.Program = getProgramFromContractNumber(model.ContractNumber, model.SellerID);
            }

            var strSQL = new StringBuilder();
            strSQL.AppendFormat(@"
                                    CREATE TABLE #foo (Market NVARCHAR(max), StartDate NVARCHAR(MAX), EndDate NVARCHAR(MAX), MeidaType NVARCHAR(max),UnitID NVARCHAR(max),MediaQuantity NVARCHAR(max),Creative NVARCHAR(max),Company NVARCHAR
                                    (max), Contact NVARCHAR(max),artfilesrecevied NVARCHAR(max),proofsent NVARCHAR(max),proofapproved NVARCHAR(max),PLssent NVARCHAR(max),materialconfirmation NVARCHAR(max),POPsUpload NVARCHAR(max))

                                    INSERT INTO #foo
                                     exec [stp_GetProductionRequirementsReportEquinox] @StartDate, @ToDate,@Advertiser,@SellerID,@ContractNumber,@CampaignName,@DMA

                                     SELECT 'Production Requirements','','','','','','','','','','','','','',''


                                    UNION ALL

                                    SELECT 'Program: ',@Program,'','','','','','','','','','','','',''

                                    UNION ALL

                                    SELECT '','','','','','','','','','','','','','',''

                                    UNION ALL

                                    SELECT 'Market','Start Date','End Date','Media Type','Unit ID','Media Quantity','CREATIVE','Company','Contact','artsfilesrecevied','proofsent','proofsapproved','PIsent','materialconfirmation','POPS Uploaded'

                                    UNION ALL

                                     SELECT Market,CAST(StartDate AS VARCHAR(max)),EndDate,MeidaType,UnitID,CAST(MediaQuantity AS VARCHAR(max)),Creative,Company,Contact,artfilesrecevied,proofsent,proofapproved,PLssent,materialconfirmation,POPsUpload FROM #foo

                                     DROP TABLE #foo  ");

            var parameterList = new[]
            {
                 new SqlParameter("@StartDate", SqlDbType.NVarChar),
                 new SqlParameter("@ToDate", SqlDbType.NVarChar),
                 new SqlParameter("@Advertiser", SqlDbType.Int),
                 new SqlParameter("@SellerID", SqlDbType.NVarChar),
                 new SqlParameter("@ContractNumber", SqlDbType.NVarChar),
                 new SqlParameter("@CampaignName", SqlDbType.NVarChar),
                 new SqlParameter("@DMA", SqlDbType.Int),
                 new SqlParameter("@Program", SqlDbType.NVarChar),

             };
            parameterList[0].Value = GetNullDate(model.FromDateRunReport);
            parameterList[1].Value = GetNullDate(model.ToDateRunReport);
            parameterList[2].Value = model.ProductionVendor;
            parameterList[3].Value = model.SellerID;
            parameterList[4].Value = model.ContractNumber;
            parameterList[5].Value = model.CampaignName;
            parameterList[6].Value = model.DMA;
            parameterList[7].Value = model.Program == null ? "" : model.Program;

            DataTable productionRequirementExportList = _dapperRepository.ExecuteDataTableQueryWithParameter(strSQL.ToString(), parameterList);
            return productionRequirementExportList;
        }

        private string getProgramFromContractNumber(string ContractNumber, int? SellerID)
        {
            var sqlStr = new StringBuilder();

            sqlStr.AppendFormat(@"
                                    SELECT OrganizationName FROM bsContract 
                                    INNER JOIN dbo.Organization ON organization.OrganizationID = AdvertiserID
                                    WHERE ContractNumber=@ContractNumber AND bsContract.SellerID=@SellerID
                                ");

            var parameters = new DynamicParameters();
            parameters.Add("ContractNumber", ContractNumber);
            parameters.Add("SellerID", SellerID);
            return _dapperRepository.ExecuteQueryFirstOrDefault<string>(sqlStr.ToString(), parameters);
        }
        public DataTable shippingRequirementExport(ShippingOrderViewModel model)
        {
            var strSQL = new StringBuilder();
            strSQL.AppendFormat(@"

                                    CREATE TABLE #foo (ContractNumber NVARCHAR(max), Advertiser NVARCHAR(MAX), Market NVARCHAR(MAX), PostingDate NVARCHAR(max),MediaType NVARCHAR(max),Size NVARCHAR
                                    (max), NumberOfUnitsToShip NVARCHAR(max),Creative NVARCHAR(max),Vendor NVARCHAR(max),ShippingAddress NVARCHAR(max),City NVARCHAR(max),State NVARCHAR(max),Zip NVARCHAR(max),Contact NVARCHAR(max),Phone NVARCHAR(max),ShippingType NVARCHAR(max))


                                     INSERT INTO #foo
                                     exec [stp_GetShippingOrdersReportEquinox] @StartDate , @ToDate ,@Advertiser,@AdvertiserContactID,@SellerID,@ContractNumber,@OrderType, @CampaignName

                                     SELECT '','','','',@OrderType + 'Order','','','','','','','','','','',''

        	 UNION ALL

                                    SELECT '','','','','Attention: ','','','','Shipping Instruction:','','','','','Emailed: ','',''

                                    UNION ALL

                                    SELECT 'Contract#','Advertiser','Market','Posting Date','Media Type','Size','#Of Units To Ship','Creative','Vendor','ShippingAddress','City','State','Zip','Contact','Phone','Shipping Type'

                                    UNION ALL

                                     SELECT ContractNumber,Advertiser,Market,PostingDate,MediaType,Size,NumberOfUnitsToShip,Creative,Vendor,ShippingAddress,City,State,Zip,Contact,Phone,ShippingType FROM #foo

                                     DROP TABLE #foo  ");

            var parameterList = new[]
            {
                 new SqlParameter("@StartDate", SqlDbType.DateTime),
                 new SqlParameter("@ToDate", SqlDbType.DateTime),
                 new SqlParameter("@Advertiser", SqlDbType.Int),
                 new SqlParameter("@AdvertiserContactID", SqlDbType.Int),
                 new SqlParameter("@SellerID", SqlDbType.Int),
                 new SqlParameter("@ContractNumber", SqlDbType.NVarChar),
                 new SqlParameter("@OrderType", SqlDbType.NVarChar),
                 new SqlParameter("@CampaignName", SqlDbType.NVarChar)

             };
            parameterList[0].Value = GetNullDate(model.FromDateRunReport);
            parameterList[1].Value = GetNullDate(model.ToDateRunReport);
            parameterList[2].Value = model.ProductionVendor;
            parameterList[3].Value = model.VendorRep;
            parameterList[4].Value = model.SellerID;
            parameterList[5].Value = model.ContractNumber;
            parameterList[6].Value = model.OrderType;
            parameterList[7].Value = model.Campaign;



            DataTable shippingOrderExportList = _dapperRepository.ExecuteDataTableQueryWithParameter(strSQL.ToString(), parameterList);
            return shippingOrderExportList;
        }

        public DataTable ExportWorkOrderExcel(WorkOrderViewModel model)
        {
            string spName = "stp_GetPostingOrderReportEquinox";
            var parameterList = new[]
            {
                 new SqlParameter("PeriodStart", SqlDbType.DateTime),
                 new SqlParameter("PeriodEnd", SqlDbType.DateTime),
                 new SqlParameter("Advertiser", SqlDbType.Int),
                 new SqlParameter("AdvertiserContactID", SqlDbType.Int),
                 new SqlParameter("SellerID", SqlDbType.Int),
                 new SqlParameter("ContractNumber", SqlDbType.NVarChar),
                 new SqlParameter("OrderType", SqlDbType.NVarChar)

        };
            parameterList[0].Value = GetNullDate(model.FromDate);
            parameterList[1].Value = GetNullDate(model.ToDate);
            parameterList[2].Value = model.ProductionVendor;
            parameterList[3].Value = model.VendorRep;
            parameterList[4].Value = model.SellerID;
            parameterList[5].Value = model.ContractNumber;
            parameterList[6].Value = model.OrderType;

            DataTable data = _dapperRepository.ExecuteDataTableProc(spName, parameterList);
            return data;
        }

        //public ReportHeaderViewModel GetReportFor(string ContractNumer, int SellerID)
        //{
        //    var sqlStr = new StringBuilder();

        //    sqlStr.AppendFormat(@"
        //                                SELECT org1.OrganizationName AS AgencyName,org2.OrganizationName AS AdvertiserName FROM dbo.bsContract 
        //                                LEFT JOIN dbo.Organization org1 ON org1.OrganizationID = bsContract.AgencyOrgID
        //                                LEFT JOIN dbo.Organization org2 ON org2.OrganizationID = bsContract.AdvertiserID
        //                                WHERE ContractNumber=@ContractNumber AND bsContract.SellerID=@SellerID

        //                        ");

        //    var parameters = new DynamicParameters();
        //    parameters.Add("ContractNumber", ContractNumer);
        //    parameters.Add("SellerID", SellerID);
        //    return _dapperRepository.ExecuteQueryFirstOrDefault<string>(sqlStr.ToString(), parameters);
        //}

        public DataTable DatatableContractBilling(ContractBillingExportToExcelSearchViewModel model)
        {
            string reportTypeForSp = "";
            string reportTypeSelect = "";
            string status = string.Join(",", model.Status);

            if (model.ReportType == "LineItem")
            {
                reportTypeForSp = " ,[Line Item], Market ,[Description] ,[Vendor Unit ID] ,[Vendor Name],[Start Date] ,[End Date], ";
                model.ReportParameters = "[Line Item], Market ,[Description] ,[Vendor Unit ID] ,[Vendor Name],[Start Date] ,[End Date]";
                reportTypeSelect = " ,[Line Item], Market ,[Description] ,[Vendor Unit ID] ,[Vendor Name], CONVERT(VARCHAR(8),[Start Date],1) AS [Start Date] ,CONVERT(VARCHAR(8),[End Date],1) AS [End Date], ";

            }
            if (model.ReportType == "Vendor")
            {
                reportTypeForSp = " ,[Vendor Name], ";
                model.ReportParameters = "[Vendor Name]";
                reportTypeSelect = " ,[Vendor Name], ";

            }
            if (model.ReportType == "Market")
            {
                reportTypeForSp = " ,Market, ";
                model.ReportParameters = "Market";
                reportTypeSelect = " ,Market, ";
            }
            if (model.ReportType == "Contract" || model.BillingReportType == "SummaryReport")
            {
                reportTypeForSp = " , ";
                model.ReportParameters = "";
                reportTypeSelect = " , ";
            }


            StringBuilder strSQL = new StringBuilder();

            if (model.BillingReportType == "DetailReport")
                strSQL.AppendFormat(@"stp_GetContractBillingDetailReportEquinoxByUsingPivot");

            if (model.BillingReportType == "SummaryReport")
                strSQL.AppendFormat(@"stp_GetContractBillingSummaryReportEquinoxByUsingPivot");

            SqlParameter[] _parameters = new SqlParameter[15];
            _parameters[0] = new SqlParameter("@PeriodStart", model.FromDate);
            _parameters[1] = new SqlParameter("@PeriodEnd", model.ToDate);
            _parameters[2] = new SqlParameter("@EnteredSince", model.EnteredSince);
            _parameters[3] = new SqlParameter("@RevisedSince", model.RevisedSince);
            _parameters[4] = new SqlParameter("@SellerID", model.SellerID);
            _parameters[5] = new SqlParameter("@AccountExecutive", model.AccountExecutive);
            _parameters[6] = new SqlParameter("@AcountCoordinator", model.AccountCoordinator);
            _parameters[7] = new SqlParameter("@Advertiser", model.Advertiser);
            _parameters[8] = new SqlParameter("@Agency", model.AgencyName);
            _parameters[9] = new SqlParameter("@Status", status);
            _parameters[10] = new SqlParameter("@CheckProrate", model.ProRate);
            _parameters[11] = new SqlParameter("@ContractNumber", model.ContractNumber);
            _parameters[12] = new SqlParameter("@ReportType", reportTypeForSp);
            _parameters[13] = new SqlParameter("@ReportParaMeters", model.ReportParameters);
            _parameters[14] = new SqlParameter("@ReportTypeSelect", reportTypeSelect);


            DataTable dt = _dapperRepository.ExecuteDataTableProc(strSQL.ToString(), _parameters);
            //  var list = dt.Select().ToList();
            return dt;

        }

        public DataTable getRevisionReportList(string contractNumber, int sellerId)
        {
            var strSql = new StringBuilder();
            strSql.AppendFormat(@"
                                 SELECT dbo.Invoice.InvoiceID,

												CASE WHEN BillingDetail.FaceID IS NULL
                                                          AND (dbo.BillingDetail.TransactionType = 'Space' OR dbo.BillingDetail.TransactionType ='Print' OR dbo.BillingDetail.TransactionType='Install')
                                                     THEN REPLACE(CASE WHEN MediaLineItemDMA.Name IS NULL
                                                                            OR MediaLineItemDMA.Name = ''
                                                                       THEN bsContractLineItemV2.Market
                                                                       ELSE MediaLineItemDMA.Name
                                                                  END, ',', '')
                                                     WHEN BillingDetail.FaceID IS NOT NULL
                                                          AND (dbo.BillingDetail.TransactionType = 'Space' OR dbo.BillingDetail.TransactionType ='Print' OR dbo.BillingDetail.TransactionType='Install')
                                                     THEN REPLACE(CASE WHEN ( DMA.Name IS NULL
                                                                              OR DMA.Name = ''
                                                                            )
                                                                       THEN ( Site.City + ' ' + Site.State )
                                                                       ELSE DMA.Name
                                                                  END, ',', '')
                                                     ELSE ''
                                                END AS Market,
                                                CONVERT(NVARCHAR(MAX), 
                                                Invoice.InvoiceDate, 101) AS [DATE],
                                                CASE WHEN CHARINDEX('-V', BillingDetail.ContractNumber) > 0 THEN 
                                                LEFT(BillingDetail.ContractNumber, CHARINDEX('-V', BillingDetail.ContractNumber)-1) ELSE BillingDetail.ContractNumber END AS REFERENCENO,
												                                                BillingDetail.PeriodDescription AS CAMPAIGN_DATES,
												                                                (BillingDetail.Amount + ISNULL(dbo.BillingDetail.Tax,0)) AS AmountDue,
                                                CASE WHEN BillingDetail.TransactionSubType = 'Commission' THEN dbo.BillingDetail.TransactionSubType ELSE 
				                                                                                CASE WHEN BillingDetail.TransactionType = 'Space'
                                                                                                     THEN 'Display'  
                                                                                                     WHEN ( dbo.BillingDetail.TransactionType = 'Printing'
                                                                                                            OR dbo.BillingDetail.TransactionType = 'Print'
                                                                                                          ) THEN 'Production'
                                                                                                     ELSE BillingDetail.TransactionType
                                                                                                END
				                                                                                END  AS ITEMID   
                                                 FROM dbo.Invoice 
                                                 INNER JOIN dbo.BillingDetail ON BillingDetail.InvoiceID = Invoice.InvoiceID
                                                 LEFT JOIN dbo.bsSelectedFace ON bsSelectedFace.SelectedFaceID = BillingDetail.SelectedFaceID
                                                LEFT JOIN dbo.bsContractLineItemV2 ON bsContractLineItemV2.ContractNumber = BillingDetail.ContractNumber
                                                                                    AND bsContractLineItemV2.LineNumber = BillingDetail.LineNumber
                                                                                    AND ( bsContractLineItemV2.LineItemStyle = 'Media'
                                                                                        OR bsContractLineItemV2.LineItemStyle = 'Open'
                                                                                        )
                                                LEFT JOIN dbo.Face ON bsSelectedFace.FaceID = Face.FaceID
                                                LEFT JOIN dbo.Site ON Site.SiteID = Face.SiteID
                                                LEFT JOIN dbo.DMA ON DMA.ID = dbo.bsSelectedFace.DMA
                                                LEFT JOIN dbo.DMA MediaLineItemDMA ON MediaLineItemDMA.ID = bsContractLineItemV2.DMA

                                                 WHERE dbo.Invoice.ContractNumber=@ContractNumber AND Invoice.SellerID=@SellerID

                           ");
            SqlParameter[] _parameters = new SqlParameter[2];
            _parameters[0] = new SqlParameter("@ContractNumber", contractNumber);
            _parameters[1] = new SqlParameter("@SellerID", sellerId);
            DataTable dt = _dapperRepository.ExecuteDataTableQueryWithParameter(strSql.ToString(), _parameters);
            return dt;
        }

        public rivisionHeaderViewModel getHeaderForReport(string contractNumber, int sellerId)
        {
            var strSql = new StringBuilder();
            strSql.AppendFormat(@"
                                 SELECT ContractNumber,
                                JobName AS CampaignName,
                                ClientCustomerID AS ClientCompanyCode,
                                 FirstName + ' ' + COALESCE(LastName, '') AS SalesPerson,
                                 ClientDemographicsTargetID +' - '+ dbo.DemographicsTarget.Description AS Demographics 
                                FROM dbo.bsContract 
                                LEFT JOIN dbo.Organization ON AdvertiserID=OrganizationID
                                LEFT JOIN dbo.Contact ON Contact.ContactID = bsContract.ACID
                                LEFT JOIN dbo.DemographicsTarget ON TargetID = ClientDemographicsTargetID
                                WHERE ContractNumber=@ContractNumber AND dbo.bsContract.SellerID=@SellerID
                           ");
            var parameters = new DynamicParameters();
            parameters.Add("ContractNumber", contractNumber);
            parameters.Add("SellerID", sellerId);
            var result =  _dapperRepository.ExecuteQueryFirstOrDefault<rivisionHeaderViewModel>(strSql.ToString(), parameters);
            return result;
        }
     public rivisionSubHeaderViewModel getSubHeaderList(string contractNumber, int SellerID)
        {
            var strSql = new StringBuilder();
            strSql.AppendFormat(@"
                                SELECT  @ContractNumber AS ContractNumber, SUM(ISNULL(BillingDetail.Amount, 0) + ISNULL(BillingDetail.Tax, 0)) AS ClientTotals ,
                                SUM(ISNULL(Payment.Amount, 0) + ISNULL(Payment.Tax, 0)) AS VendorTotals
                                INTO    #temp
                                FROM    dbo.BillingDetail WITH ( NOLOCK )
                                        LEFT  JOIN dbo.Payment WITH ( NOLOCK ) ON Payment.BillingID = BillingDetail.BillingID
                                WHERE   BillingDetail.ContractNumber = @ContractNumber
                                        AND BillingDetail.SellerID = @SellerID
                                        AND Payment.TransactionSubType IS NULL
                                SELECT  ContractNumber,ISNULL(ROUND(ClientTotals, 2), 0) AS ClientTotals ,
                                        ISNULL(ROUND(VendorTotals, 2), 0) AS VendorTotals ,
                                        CASE WHEN ISNULL(VendorTotals, 0) = 0 THEN 0
                                             ELSE ( ( ISNULL(ROUND(ClientTotals, 2), 0)
                                                      - ISNULL(ROUND(VendorTotals, 2), 0) )
                                                    / ISNULL(ROUND(VendorTotals, 2), 0) ) * 100
                                        END AS ProfitPercent ,
                                         (SELECT ISNULL(SUM(ROUND(( ISNULL(Amount, 0) + ISNULL(Tax, 0) ), 2)),0) AS [ClientTotalDisplay]
                                                 FROM   dbo.BillingDetail WITH ( NOLOCK )
                                                 WHERE  ContractNumber = @ContractNumber
                                                        AND TransactionType = 'Space'
                                                        AND BillingDetail.SellerID = @SellerID
                                               ) AS [ClientTotalDisplay] ,
                                        ( SELECT ISNULL(SUM(ROUND(( ISNULL(Amount, 0) + ISNULL(Tax, 0) ), 2)),0) AS [ClientTotalProduction]
                                                 FROM   dbo.BillingDetail WITH ( NOLOCK )
                                                 WHERE  ContractNumber = @ContractNumber
                                                        AND ( TransactionType = 'Print'
                                                              OR TransactionType = 'install'
                                                            )
                                                        AND BillingDetail.SellerID = @SellerID
                                               ) AS [ClientTotalProduction]
                                FROM    #temp
                                DROP TABLE #temp

                           ");
            var parameters = new DynamicParameters();
            parameters.Add("@ContractNumber", contractNumber);
            parameters.Add("@SellerID", SellerID);

            var result = _dapperRepository.ExecuteQueryFirstOrDefault<rivisionSubHeaderViewModel>(strSql.ToString(), parameters);
            return result;
        }

        public DataTable getAllContractRevision(string contractNumber)
        {
            var strSql = new StringBuilder();
            strSql.AppendFormat(@"
                                 SELECT ContractNumber,PrivateNotes,CONVERT(NVARCHAR(MAX), 
                                 ISNULL(TS,CreatedTS), 101) AS [DATE] FROM dbo.bsContract WHERE ContractNumber LIKE @ContractNumber + '%'
                           ");
            //var parameters = new DynamicParameters();
            //parameters.Add("ContractNumber", contractNumber);
            //var result = await _dapperRepository.ExecuteQueryAsync<rivisionSubHeaderViewModel>(strSql.ToString(), parameters);
            //return result.ToArray();

            SqlParameter[] _parameters = new SqlParameter[1];
            _parameters[0] = new SqlParameter("@ContractNumber", contractNumber);
            DataTable dt = _dapperRepository.ExecuteDataTableQueryWithParameter(strSql.ToString(), _parameters);
            return dt;
        }
        private object GetNullDate(DateTime? value)
        {
            if (value == null) { return DBNull.Value; }
            else { return value.Value; }
        }

    }
}
