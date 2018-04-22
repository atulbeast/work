using System;
using System.ComponentModel;
using System.Reflection;

namespace Aspose.Slides.Model
{
    public enum Permissions
    {
        ReadProject,
        ManageProject
    }
    public class EnumValue
    {
        
        public enum TypeOfDeal : int
        {
            [Description("Acquisition(shares)")]
            AcquisitionShares = 1,
            [Description("Acquisition(assets)")]
            AcquisitionAssets = 2,
            [Description("Acquisition(tender + shares)")]
            AcquisitionTenderShares = 3,
            [Description("JV or Co - operation")]
            JvOrCoOperation = 4,
            [Description("Minority stake")]
            MinorityStake = 5,
            [Description("Tender only")]
            TenderOnly = 6
        }

        public enum TenderStatus : int
        {
            [Description("New Tender")]
            NewTender = 1,
            [Description("Existing Tender")]
            ExistingTender = 2,
            [Description("Combination")]
            Combination = 3
        }

        public enum Mode : int
        {
            [Description("Bus")]
            Bus = 1,
            [Description("Rail")]
            Rail = 2,
            [Description("Ambulance")]
            Ambulance = 3,
            [Description("Ferry")]
            Ferry = 4,
            [Description("Combination")]
            Combination = 5,
            [Description("Other")]
            Other = 6
        }

        public enum Ownership : int
        {
            [Description("Private")]
            Private = 1,
            [Description("Public")]
            Public = 2
        }

        public enum Mtp : int
        {
            [Description("Y")]
            Yes = 1,
            [Description("N")]
            No = 2
        }

        public enum SuggestedPriority : int
        {
            [Description("High")]
            High = 1,
            [Description("Medium")]
            Medium = 2,
            [Description("Low")]
            Low = 3
        }

        public enum Priority : int
        {
            [Description("1")]
            One = 1,
            [Description("2")]
            Two = 2,
            [Description("3")]
            Three = 3
        }

        public enum DealStatus : int
        {
            [Description("Won")]
            Won = 1,
            [Description("Lost")]
            Lost = 2,
            [Description("No Bid")]
            NoBid = 3,
            [Description("Awaiting Decision")]
            AwaitingDecision = 4
        }

        public enum DealProbability : int
        {
            [Description("High")]
            High = 1,
            [Description("Medium")]
            Medium = 2,
            [Description("Low")]
            Low = 3
        }

        public enum ReasonForNoBid : int
        {
            [Description("Cancellation of tender / M & A")]
            Cancelation = 1,
            [Description("Prioritisation of resource")]
            PrioritisationOfResource = 2,
            [Description("Competition / likelihood of win")]
            CompetitionLikelihoodOfWin = 3,
            [Description("Unacceptable contract")]
            UnacceptableContract = 4
        }

        public enum BusinessUnit : int
        {
            [Description("NE")]
            NE = 1,
            [Description("SCEE")]
            SCEE = 2,
            [Description("NBD")]
            NBD = 3,
            [Description("UK Bus")]
            UKBus = 4,
            [Description("UK Rail")]
            UKRail = 5
        }

        public enum ProjectStageTenders : int
        {
            [Description("1 - Monitoring")]
            Monitoring = 1,
            [Description("2 - ITT Assessment")]
            ITTAssessment = 2,
            [Description("3 - Bid Preparation")]
            BidPreparation = 3,
            [Description("4 - Bid Submission")]
            BidSubmission = 4,
            [Description("5 - Mobilisation")]
            Mobilisation = 5,
            [Description("6 - Post Tender Review")]
            PostTenderReview = 6
        }

        public enum ProjectStageMAndA : int
        {
            [Description("1 - Scouting")]
            Scouting = 1,
            [Description("2 - Relationship Building")]
            RelationshipBuilding = 2,
            [Description("3 - DD / Valuation")]
            DDValuation = 3,
            [Description("4 - Contract Negotiation")]
            ContractNegotiation = 4,
            [Description("5 - Closing")]
            Closing = 5,
            [Description("6 - PMI")]
            PMI = 6
        }

        public enum ProjectStatus : int
        {
            [Description("Completed")]
            Completed = 1,
            [Description("On track")]
            OnTrack = 2,
            [Description("Issue")]
            Issue = 3,
            [Description("Off Track")]
            OffTrack = 4
        }

        public enum ProjectRag : int
        {
            [Description("Red")]
            Red = 1,
            [Description("Amber")]
            Amber = 2,
            [Description("Green")]
            Green = 3
        }

        public enum ContractType : int
        {
            [Description("Gross Cost")]
            GrossCost = 1,
            [Description("Net Cost")]
            NetCost = 2,
            [Description("Other")]
            Other = 3
        }

        public enum TypeOfAcquisition : int
        {
            [Description("Share Deal")]
            ShareDeal = 1,
            [Description("Asset Deal")]
            AssetDeal = 2
        }

        public enum WorkflowStage : int
        {
            [Description("Initiation")]
            Initiation = 1,
            [Description("Review")]
            Review = 2,
            [Description("Locked")]
            Locked = 3
        }

        public static string GetEnumDescription(Enum value)
        {

          FieldInfo fi = value.GetType().GetField(value.ToString());

            DescriptionAttribute[] attributes =
                (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);

            if (attributes != null && attributes.Length > 0)
                return attributes[0].Description;
            else
                return value.ToString();
        }
    }
}
