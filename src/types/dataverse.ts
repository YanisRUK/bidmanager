/**
 * Dataverse Data Model for Ricoh Bid Management System
 *
 * Tables (Dataverse logical names use lowercase with publisher prefix, e.g. ricoh_):
 *   ricoh_bidrequest      – every incoming bid/opportunity
 *   ricoh_bidworkspace    – workspace created once a bid is accepted / in-flight
 *   ricoh_bidtype         – lookup: 4 bid types
 *   ricoh_bidteammember   – junction: user roles per workspace
 *   ricoh_bidapproval     – approval records linked to a workspace
 *   ricoh_biddocument     – SharePoint document references per workspace
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
};

export enum TeamRole {
  BidManager    = 100000000,
  BidCoordinator = 100000001,
  SME           = 100000002,
  Approver      = 100000003,
  Reviewer      = 100000004,
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

// ---------------------------------------------------------------------------
// Base Dataverse record shape
// ---------------------------------------------------------------------------

export interface DataverseRecord {
  /** GUID primary key */
  id: string;
  createdOn: string;   // ISO 8601
  modifiedOn: string;  // ISO 8601
  createdBy: DataverseUser;
  modifiedBy: DataverseUser;
}

// ---------------------------------------------------------------------------
// Dataverse system user (simplified)
// ---------------------------------------------------------------------------

export interface DataverseUser {
  id: string;
  fullName: string;
  email: string;
  /** Azure AD object ID */
  azureObjectId?: string;
}

// ---------------------------------------------------------------------------
// ricoh_bidtype  (mostly a lookup / reference table)
// ---------------------------------------------------------------------------

export interface BidType extends DataverseRecord {
  ricoh_code: BidTypeCode;
  ricoh_name: string;
  ricoh_description: string;
  /** Which team / queue to route to */
  ricoh_routingTeam: string;
  /** Email address for routing notifications */
  ricoh_routingEmail: string;
  /** SLA days to first response */
  ricoh_slaResponseDays: number;
}

// ---------------------------------------------------------------------------
// ricoh_bidrequest  (intake form submission)
// ---------------------------------------------------------------------------

export interface BidRequest extends DataverseRecord {
  ricoh_bidreferencenumber: string;       // auto-generated, e.g. BID-2026-00042
  ricoh_title: string;
  ricoh_bidtypeid: Pick<BidType, "id" | "ricoh_name" | "ricoh_code">;
  ricoh_status: BidStatus;

  // Customer / opportunity details
  ricoh_customername: string;
  ricoh_customerindustry?: string;
  ricoh_estimatedvalue?: number;          // GBP
  ricoh_currency: string;                 // default "GBP"

  // Dates
  ricoh_submissiondeadline: string;       // ISO 8601
  ricoh_expectedawarddate?: string;
  ricoh_contractstartdate?: string;
  ricoh_contractduration?: number;        // months

  // Narrative
  ricoh_description: string;
  ricoh_scope?: string;
  ricoh_specialrequirements?: string;
  ricoh_incumbentvendor?: string;

  // Routing / ownership
  ricoh_submittedby: DataverseUser;
  ricoh_assignedto?: DataverseUser;
  ricoh_routedto?: string;                // team name

  // Linked workspace (set once workspace created)
  ricoh_bidworkspaceid?: Pick<BidWorkspace, "id" | "ricoh_title">;

  // Bid-type-specific fields (stored as JSON in a multiline text column,
  // parsed on the client)
  ricoh_typespecificdata?: string;
}

// ---------------------------------------------------------------------------
// ricoh_bidworkspace  (active bid workspace)
// ---------------------------------------------------------------------------

export interface BidWorkspace extends DataverseRecord {
  ricoh_title: string;
  ricoh_bidrequestid: Pick<BidRequest, "id" | "ricoh_title" | "ricoh_bidreferencenumber">;
  ricoh_status: BidStatus;
  ricoh_bidmanagerid: DataverseUser;

  // Progress (0-100)
  ricoh_completionpercentage: number;

  // SharePoint document library root URL for this workspace
  ricoh_sharepointfolderurl?: string;

  // Teams channel deep-link
  ricoh_teamschannelurl?: string;

  // Win/loss data
  ricoh_wonlostdate?: string;
  ricoh_wonlostreason?: string;

  // Relationships (loaded separately / expanded)
  teamMembers?: BidTeamMember[];
  approvals?: BidApproval[];
  documents?: BidDocument[];
}

// ---------------------------------------------------------------------------
// ricoh_bidteammember
// ---------------------------------------------------------------------------

export interface BidTeamMember extends DataverseRecord {
  ricoh_bidworkspaceid: Pick<BidWorkspace, "id" | "ricoh_title">;
  ricoh_userid: DataverseUser;
  ricoh_role: TeamRole;
  ricoh_isactive: boolean;
  ricoh_assigneddate: string;
}

// ---------------------------------------------------------------------------
// ricoh_bidapproval
// ---------------------------------------------------------------------------

export interface BidApproval extends DataverseRecord {
  ricoh_bidworkspaceid: Pick<BidWorkspace, "id" | "ricoh_title">;
  ricoh_title: string;
  ricoh_approverstage: number;            // 1, 2, 3 …
  ricoh_approverid: DataverseUser;
  ricoh_status: ApprovalStatus;
  ricoh_requesteddate: string;
  ricoh_respondeddate?: string;
  ricoh_comments?: string;
}

// ---------------------------------------------------------------------------
// ricoh_biddocument  (reference to SharePoint document)
// ---------------------------------------------------------------------------

export interface BidDocument extends DataverseRecord {
  ricoh_bidworkspaceid: Pick<BidWorkspace, "id" | "ricoh_title">;
  ricoh_title: string;
  ricoh_documenttype: string;             // e.g. "Response", "Pricing", "Legal"
  ricoh_sharepointurl: string;
  ricoh_uploadedby: DataverseUser;
  ricoh_filesize?: number;                // bytes
  ricoh_version: string;
}

// ---------------------------------------------------------------------------
// API / SDK wrappers
// ---------------------------------------------------------------------------

export interface DataverseQueryOptions {
  filter?: string;    // OData $filter
  select?: string[];  // $select columns
  expand?: string[];  // $expand
  orderBy?: string;   // $orderby
  top?: number;       // $top
  skip?: number;      // $skip
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
  messageType: "assignment" | "approval_request" | "deadline_reminder" | "status_change";
  additionalContext?: Record<string, string>;
}
