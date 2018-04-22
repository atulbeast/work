using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Slides.Model
{
    class ProjectModel
    {

        public ProjectModel()
        {


            IList<KeyMilestoneAndActionModel> KeyMileStoneAndAction = new List<KeyMilestoneAndActionModel>();
            IList<IssuesModel> Issues = new List<IssuesModel>();
            IList<RisksModel> Risks = new List<RisksModel>();
        }

        public string ProjectId { get; set; }
         
        public string Country { get; set; }
         
        public string ProjectName { get; set; }
         
        public string Region { get; set; }
         
        public int TypeOfDeal { get; set; }
         
        public int New_or_existingwork { get; set; }
         
        public int Mode { get; set; }
         
        public int NumberOfVehicles { get; set; }
         
        public int Ownership { get; set; }
        public double T_OMEuro { get; set; }
        public double EBIT_Value { get; set; }
        public double EBITDA_Value { get; set; }
        public double EnterpriseValue { get; set; }
        public double LifetimeContractT_O { get; set; }
        public double LifetimeContractEBIT { get; set; }
        public double TotalCAPEX { get; set; }
        public double EBIT_CAPEX { get; set; }
        public double CAPEXXMEuro { get; set; }
        public double CAPEX1MEuro { get; set; }
        public double CAPEX2MEuro { get; set; }
         
        public int MTP { get; set; }
         
        public DateTime DirectorsApproval { get; set; }
         
        public int SuggestedPriority { get; set; }
         
        public int DealStatus { get; set; }
         
        public int DealProbability { get; set; }
         
        public DateTime StartOfOperations { get; set; }
         
        public int CoreContractLength { get; set; }
         
        public int OptionalExtension { get; set; }
         
        public int ReasonForNoBid { get; set; }
         
        public string Comments { get; set; }
        //Additional Details
        public int BusinessUnit { get; set; }
        public string Category { get; set; }
        //Tender 
        public int ProjectStageTender { get; set; }
        public double RevenueX { get; set; }
        public double RevenueX1 { get; set; }
        public double RevenueX2 { get; set; }
        public double EbitX { get; set; }
        public double EbitX1 { get; set; }
        public double EbitX2 { get; set; }
        //m&A
        public int ProjectStageMandA { get; set; }
        public double RevenueFullYearX { get; set; }
        public double RevenueFullYearX1 { get; set; }
        public double RevenueFullYearX2 { get; set; }
        public double EbitFullYearX { get; set; }
        public double EbitFullYearX1 { get; set; }
        public double EbitFullYearX2 { get; set; }
        public double EbitContributionYearX { get; set; }
        public double EbitContributionYearX1 { get; set; }
        public double EbitContributionYearX2 { get; set; }
        public double EbitDAX { get; set; }
        public double EbitDAX1 { get; set; }
        public double EbitDAX2 { get; set; }
        public double EnterpriseValueX { get; set; }
        public double EnterpriseValueX1 { get; set; }
        public double EnterpriseValueX2 { get; set; }
        public double EV_EbitDAX { get; set; }
        public double EV_EbitDAX1 { get; set; }


        //Project Governance
        public string ProjectImage { get; set; }
        public string Description { get; set; }
        public string StrategicRationale { get; set; }

        //Fields related to Project Status
        public int ContractType { get; set; }
        public string ExecutiveBoardMember { get; set; }
        public string CountryManager { get; set; }
        public string ProjectManager { get; set; }
        public string AdditionalTeamMembers { get; set; }
        public string UncoveredTeamResources { get; set; }
        public int TypeOfAcquisition { get; set; }
        public int ProjectStatusCurrentMonth { get; set; }
        public int ProjectStatusPreviousMonth { get; set; }
        public int ProjectRAGCurrentMonth { get; set; }
        public int ProjectRAGPreviousMonth { get; set; }
        public string StatusDescription { get; set; }
        public IList<KeyMilestoneAndActionModel> KeyMileStoneAndAction { get; set; }
        public IList<IssuesModel> Issues { get; set; }
        public IList<RisksModel> Risks { get; set; }
        public bool ITTSummary { get; set; }
        public bool FullContractTranslation { get; set; }
        public bool TenderPresentation { get; set; }
        public bool LegalRiskSummary { get; set; }
        public bool KIRS { get; set; }
        public bool FinancialModel { get; set; }
        public bool InformationPaperNBO { get; set; }
        public bool ProjectStatusLetter { get; set; }
        public bool KeyFindingPhase1 { get; set; }
        public bool KeyFindingPhase2 { get; set; }
        public bool DDReports { get; set; }
        public bool BPValuation { get; set; }
        public bool SPADraft { get; set; }
        public bool ExecCommitteePaper { get; set; }

        //other Fields
        public int Priority { get; set; }
        public bool ManualPriorityOverride { get; set; }
        public int WorkFlowStatus { get; set; }
        public bool ExcludeFromGrowthReport { get; set; }
    }
}
