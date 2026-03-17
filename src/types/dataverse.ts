/**
 * Dataverse Data Model for Ricoh Bid Management System
 *
 * Table logical names (publisher prefix cr5ab_):
 *   cr5ab_bidrequest        – every incoming bid/opportunity
 *   cr5ab_bidworkspace      – workspace created once a bid is in-flight
 *   cr5ab_bidtype           – lookup: 4 bid types
 *   cr5ab_bidroleassignment – user roles per workspace
 *   cr5ab_approval          – approval records linked to a workspace
 *   cr5ab_bidstatus         – bid status lookup
 *   cr5ab_toritem           – Table of Responsibility row
 *   cr5ab_clarification     – clarification questions / customer comms
 */

// ---------------------------------------------------------------------------
// Enums / option-set values
// ---------------------------------------------------------------------------

export enum BidTypeCode {
  SupplierQuestionnaire = 1,
  SalesLed              = 2,
  SMEToQualify          = 3,
  BidManagement         = 4,
}

export const BidTypeLabel: Record<BidTypeCode, string> = {
  [BidTypeCode.SupplierQuestionnaire]: "Supplier Questionnaire",
  [BidTypeCode.SalesLed]:              "Sales Led",
  [BidTypeCode.SMEToQualify]:          "SME to Qualify",
  [BidTypeCode.BidManagement]:         "Bid Management",
};

export enum BidStatus {
  Draft       = 100000000,
  Submitted   = 100000001,
  InReview    = 100000002,
  Qualified   = 100000003,
  InProgress  = 100000004,
  Won         = 100000005,
  Lost        = 100000006,
  Withdrawn   = 100000007,
  NoBid       = 100000008,
}

export const BidStatusLabel: Record<BidStatus, string> = {
  [BidStatus.Draft]:      "Draft",
  [BidStatus.Submitted]:  "Submitted",
  [BidStatus.InReview]:   "In Review",
  [BidStatus.Qualified]:  "Qualified",
  [BidStatus.InProgress]: "In Progress",
  [BidStatus.Won]:        "Won",
  [BidStatus.Lost]:       "Lost",
  [BidStatus.Withdrawn]:  "Withdrawn",
  [BidStatus.NoBid]:      "No Bid",
};

export type BidStatusBadge = "success" | "warning" | "danger" | "informative" | "subtle";

export const BidStatusColor: Record<BidStatus, BidStatusBadge> = {
  [BidStatus.Draft]:      "subtle",
  [BidStatus.Submitted]:  "informative",
  [BidStatus.InReview]:   "warning",
  [BidStatus.Qualified]:  "informative",
  [BidStatus.InProgress]: "warning",
  [BidStatus.Won]:        "success",
  [BidStatus.Lost]:       "danger",
  [BidStatus.Withdrawn]:  "subtle",
  [BidStatus.NoBid]:      "subtle",
};

// ---------------------------------------------------------------------------
// Opportunity lifecycle stage (pre-qualification pipeline)
// ---------------------------------------------------------------------------

export enum OpportunityStage {
  Identified        = 1,
  PendingOAF        = 2,
  UnderQualification = 3,
  Allocated         = 4,
  NoBid             = 5,
}

export const OpportunityStageLabel: Record<OpportunityStage, string> = {
  [OpportunityStage.Identified]:         "Identified",
  [OpportunityStage.PendingOAF]:         "Pending OAF",
  [OpportunityStage.UnderQualification]: "Under Qualification",
  [OpportunityStage.Allocated]:          "Allocated",
  [OpportunityStage.NoBid]:              "No Bid",
};

export type OpportunityStageBadge = "success" | "warning" | "danger" | "informative" | "subtle";

export const OpportunityStageColor: Record<OpportunityStage, OpportunityStageBadge> = {
  [OpportunityStage.Identified]:         "subtle",
  [OpportunityStage.PendingOAF]:         "warning",
  [OpportunityStage.UnderQualification]: "warning",
  [OpportunityStage.Allocated]:          "success",
  [OpportunityStage.NoBid]:              "danger",
};

// ---------------------------------------------------------------------------
// Bid source / channel
// ---------------------------------------------------------------------------

export enum BidSource {
  Portal       = 1,
  SalesSubmitted = 2,
  Proactive    = 3,
  Other        = 4,
}

export const BidSourceLabel: Record<BidSource, string> = {
  [BidSource.Portal]:         "Portal",
  [BidSource.SalesSubmitted]: "Sales Submitted",
  [BidSource.Proactive]:      "Proactive",
  [BidSource.Other]:          "Other",
};

// ---------------------------------------------------------------------------
// Qualification outcome
// ---------------------------------------------------------------------------

export enum QualificationOutcome {
  BidManagement = 1,
  SalesLed      = 2,
  LightSupport  = 3,
  NoBid         = 4,
}

export const QualificationOutcomeLabel: Record<QualificationOutcome, string> = {
  [QualificationOutcome.BidManagement]: "Bid Management",
  [QualificationOutcome.SalesLed]:      "Sales Led",
  [QualificationOutcome.LightSupport]:  "Light Support",
  [QualificationOutcome.NoBid]:         "No Bid",
};

export const QualificationOutcomeColor: Record<QualificationOutcome, BidStatusBadge> = {
  [QualificationOutcome.BidManagement]: "success",
  [QualificationOutcome.SalesLed]:      "informative",
  [QualificationOutcome.LightSupport]:  "warning",
  [QualificationOutcome.NoBid]:         "danger",
};

// ---------------------------------------------------------------------------
// Clarification status
// ---------------------------------------------------------------------------

export enum ClarificationStatus {
  Open           = 1,
  AnswerReceived = 2,
  Closed         = 3,
}

export const ClarificationStatusLabel: Record<ClarificationStatus, string> = {
  [ClarificationStatus.Open]:           "Open",
  [ClarificationStatus.AnswerReceived]: "Answer Received",
  [ClarificationStatus.Closed]:         "Closed",
};

export const ClarificationStatusColor: Record<ClarificationStatus, "success" | "warning" | "informative"> = {
  [ClarificationStatus.Open]:           "warning",
  [ClarificationStatus.AnswerReceived]: "informative",
  [ClarificationStatus.Closed]:         "success",
};

// ---------------------------------------------------------------------------
// Team / approval / TOR enums (unchanged)
// ---------------------------------------------------------------------------

export enum TeamRole {
  BidManager     = 100000000,
  BidCoordinator = 100000001,
  SME            = 100000002,
  Approver       = 100000003,
  Reviewer       = 100000004,
}

export const TeamRoleLabel: Record<TeamRole, string> = {
  [TeamRole.BidManager]:     "Bid Manager",
  [TeamRole.BidCoordinator]: "Bid Coordinator",
  [TeamRole.SME]:            "Subject Matter Expert",
  [TeamRole.Approver]:       "Approver",
  [TeamRole.Reviewer]:       "Reviewer",
};

export enum ApprovalStatus {
  Pending  = 100000000,
  Approved = 100000001,
  Rejected = 100000002,
}

export enum TorAnsweredStatus {
  Yes     = "Y",
  Partial = "P",
  No      = "N",
}

export const TorAnsweredStatusLabel: Record<TorAnsweredStatus, string> = {
  [TorAnsweredStatus.Yes]:     "Answered",
  [TorAnsweredStatus.Partial]: "Partial",
  [TorAnsweredStatus.No]:      "Not Started",
};

export const TorAnsweredStatusColor: Record<TorAnsweredStatus, "success" | "warning" | "danger"> = {
  [TorAnsweredStatus.Yes]:     "success",
  [TorAnsweredStatus.Partial]: "warning",
  [TorAnsweredStatus.No]:      "danger",
};

// ---------------------------------------------------------------------------
// Base record shape
// ---------------------------------------------------------------------------

export interface DataverseRecord {
  id: string;
  createdOn: string;
  modifiedOn: string;
  createdBy: DataverseUser;
  modifiedBy: DataverseUser;
}

export interface DataverseUser {
  id: string;
  fullName: string;
  email: string;
  azureObjectId?: string;
}

// ---------------------------------------------------------------------------
// cr5ab_bidtype
// ---------------------------------------------------------------------------

export interface BidType extends DataverseRecord {
  cr5ab_code: BidTypeCode;
  cr5ab_name: string;
  cr5ab_description: string;
  cr5ab_routingteam: string;
  cr5ab_routingemail: string;
  cr5ab_slaresponsedays: number;
}

// ---------------------------------------------------------------------------
// cr5ab_bidrequest  (extended with qualification / source fields)
// ---------------------------------------------------------------------------

export interface BidRequest extends DataverseRecord {
  cr5ab_bidreferencenumber: string;
  cr5ab_title: string;
  cr5ab_bidtypeid: Pick<BidType, "id" | "cr5ab_name" | "cr5ab_code">;
  cr5ab_status: BidStatus;

  // Opportunity pipeline stage
  cr5ab_opportunitystage?: OpportunityStage;

  // Source / channel
  cr5ab_source?: BidSource;
  cr5ab_sourceportalname?: string;   // e.g. "Contracts Finder", "Find a Tender"

  // Customer
  cr5ab_customername: string;
  cr5ab_customerindustry?: string;
  cr5ab_estimatedvalue?: number;
  cr5ab_currency: string;

  // Dates
  cr5ab_submissiondeadline: string;
  cr5ab_expectedawarddate?: string;
  cr5ab_contractstartdate?: string;
  cr5ab_contractduration?: number;

  // Content
  cr5ab_description: string;
  cr5ab_scope?: string;
  cr5ab_specialrequirements?: string;
  cr5ab_incumbentvendor?: string;

  // People
  cr5ab_submittedby: DataverseUser;
  cr5ab_assignedto?: DataverseUser;
  cr5ab_routedto?: string;

  // OAF (Opportunity Assessment Form) — stored as JSON string
  cr5ab_oafdata?: string;

  // Qualification decision
  cr5ab_qualificationoutcome?: QualificationOutcome;
  cr5ab_qualificationrationale?: string;
  cr5ab_qualifiedby?: DataverseUser;
  cr5ab_qualifiedon?: string;

  // Relationships
  cr5ab_bidworkspaceid?: Pick<BidWorkspace, "id" | "cr5ab_title">;
  cr5ab_typespecificdata?: string;
}

// ---------------------------------------------------------------------------
// OAF structured data (stored JSON in cr5ab_oafdata)
// ---------------------------------------------------------------------------

export interface OafData {
  servicesInScope: string;
  keyRisks: string;
  hasIncumbent: boolean;
  incumbentName?: string;
  relationshipsInPlace: boolean;
  estimatedBidEffortDays: string;
  recommendedResource: string;
  additionalNotes: string;
}

export const EMPTY_OAF: OafData = {
  servicesInScope: "",
  keyRisks: "",
  hasIncumbent: false,
  incumbentName: "",
  relationshipsInPlace: false,
  estimatedBidEffortDays: "",
  recommendedResource: "",
  additionalNotes: "",
};

// ---------------------------------------------------------------------------
// cr5ab_bidworkspace
// ---------------------------------------------------------------------------

export interface BidWorkspace extends DataverseRecord {
  cr5ab_title: string;
  cr5ab_bidrequestid: Pick<BidRequest, "id" | "cr5ab_title" | "cr5ab_bidreferencenumber">;
  cr5ab_status: BidStatus;
  cr5ab_bidmanagerid: DataverseUser;
  cr5ab_completionpercentage: number;
  cr5ab_sharepointfolderurl?: string;
  cr5ab_teamschannelurl?: string;
  cr5ab_wonlostdate?: string;
  cr5ab_wonlostreason?: string;

  // Expanded relationships
  teamMembers?: BidRoleAssignment[];
  approvals?: BidApproval[];
  documents?: BidDocumentRecord[];
  clarifications?: BidClarification[];
}

// ---------------------------------------------------------------------------
// cr5ab_bidroleassignment
// ---------------------------------------------------------------------------

export interface BidRoleAssignment extends DataverseRecord {
  cr5ab_bidworkspaceid: Pick<BidWorkspace, "id" | "cr5ab_title">;
  cr5ab_userid: DataverseUser;
  cr5ab_role: TeamRole;
  cr5ab_isactive: boolean;
  cr5ab_assigneddate: string;
}

// ---------------------------------------------------------------------------
// cr5ab_approval
// ---------------------------------------------------------------------------

export interface BidApproval extends DataverseRecord {
  cr5ab_bidworkspaceid: Pick<BidWorkspace, "id" | "cr5ab_title">;
  cr5ab_title: string;
  cr5ab_approverstage: number;
  cr5ab_approverid: DataverseUser;
  cr5ab_status: ApprovalStatus;
  cr5ab_requesteddate: string;
  cr5ab_respondeddate?: string;
  cr5ab_comments?: string;
}

// ---------------------------------------------------------------------------
// cr5ab_bidstatus (lookup table)
// ---------------------------------------------------------------------------

export interface BidStatusRecord extends DataverseRecord {
  cr5ab_name: string;
  cr5ab_statuscode: BidStatus;
  cr5ab_description?: string;
}

// ---------------------------------------------------------------------------
// cr5ab_toritem — Table of Responsibility row
// ---------------------------------------------------------------------------

export interface TorItem extends DataverseRecord {
  cr5ab_section: string;
  cr5ab_questionnumber: string;
  cr5ab_questiondetail: string;
  cr5ab_allocatedto?: DataverseUser;
  cr5ab_department?: string;
  cr5ab_supportcontributors?: string;
  cr5ab_scoreweightingthreshold?: string;
  cr5ab_specialinstructions?: string;
  cr5ab_firstdraftdeadline?: string;
  cr5ab_reviewperiod?: string;
  cr5ab_finaldraftdeadline?: string;
  cr5ab_actualdeadline?: string;
  cr5ab_comments?: string;
  cr5ab_unittime?: string;
  cr5ab_answeredstatus: TorAnsweredStatus;
  cr5ab_bidworkspaceid: Pick<BidWorkspace, "id" | "cr5ab_title">;
}

// ---------------------------------------------------------------------------
// cr5ab_clarification — clarification question record
// ---------------------------------------------------------------------------

export interface BidClarification extends DataverseRecord {
  cr5ab_bidworkspaceid: Pick<BidWorkspace, "id" | "cr5ab_title">;
  cr5ab_questionnumber: string;         // e.g. "CQ-01"
  cr5ab_questiontext: string;
  cr5ab_raisedby: DataverseUser;
  cr5ab_raiseddate: string;
  cr5ab_deadline?: string;
  cr5ab_responsetext?: string;
  cr5ab_respondeddate?: string;
  cr5ab_status: ClarificationStatus;
  cr5ab_iscustomerraised: boolean;      // true = customer asked us; false = we asked customer
}

// ---------------------------------------------------------------------------
// cr5ab_biddocument — document linked to a workspace (category-aware)
// ---------------------------------------------------------------------------

export type DocumentCategory = "originals" | "working" | "submission";

export const DocumentCategoryLabel: Record<DocumentCategory, string> = {
  originals:  "Originals",
  working:    "Working Documents",
  submission: "Submission Documents",
};

export interface BidDocumentRecord extends DataverseRecord {
  cr5ab_title: string;
  cr5ab_filename?: string;
  cr5ab_documenttype: string;
  cr5ab_category: DocumentCategory;
  cr5ab_sharepointurl: string;
  cr5ab_filesize?: number;
  cr5ab_version: string;
  cr5ab_uploadedby: DataverseUser;
  cr5ab_bidworkspaceid: Pick<BidWorkspace, "id" | "cr5ab_title">;
}

// ---------------------------------------------------------------------------
// Query helpers
// ---------------------------------------------------------------------------

export interface DataverseQueryOptions {
  filter?: string;
  select?: string[];
  expand?: string[];
  orderBy?: string;
  top?: number;
  skip?: number;
}

export interface PagedResult<T> {
  value: T[];
  totalCount?: number;
  nextLink?: string;
}

// ---------------------------------------------------------------------------
// Power Automate flow payloads
// ---------------------------------------------------------------------------

export interface RouteBidFlowPayload {
  bidRequestId: string;
  bidTypeCode: BidTypeCode;
  submittedById: string;
}

export interface NotifyTeamFlowPayload {
  bidWorkspaceId: string;
  recipientIds: string[];
  messageType: "assignment" | "approval_request" | "deadline_reminder" | "status_change" | "qualification_decision";
  additionalContext?: Record<string, string>;
}
