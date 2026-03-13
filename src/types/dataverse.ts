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

// ---------------------------------------------------------------------------
// Base Dataverse record shape
// ---------------------------------------------------------------------------

export interface DataverseRecord {
  id: string;
  createdOn: string;
  modifiedOn: string;
  createdBy: DataverseUser;
  modifiedBy: DataverseUser;
}

// ---------------------------------------------------------------------------
// Dataverse system user
// ---------------------------------------------------------------------------

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
// cr5ab_bidrequest
// ---------------------------------------------------------------------------

export interface BidRequest extends DataverseRecord {
  cr5ab_bidreferencenumber: string;
  cr5ab_title: string;
  cr5ab_bidtypeid: Pick<BidType, "id" | "cr5ab_name" | "cr5ab_code">;
  cr5ab_status: BidStatus;

  cr5ab_customername: string;
  cr5ab_customerindustry?: string;
  cr5ab_estimatedvalue?: number;
  cr5ab_currency: string;

  cr5ab_submissiondeadline: string;
  cr5ab_expectedawarddate?: string;
  cr5ab_contractstartdate?: string;
  cr5ab_contractduration?: number;

  cr5ab_description: string;
  cr5ab_scope?: string;
  cr5ab_specialrequirements?: string;
  cr5ab_incumbentvendor?: string;

  cr5ab_submittedby: DataverseUser;
  cr5ab_assignedto?: DataverseUser;
  cr5ab_routedto?: string;

  cr5ab_bidworkspaceid?: Pick<BidWorkspace, "id" | "cr5ab_title">;
  cr5ab_typespecificdata?: string;
}

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
  documents?: BidDocument[];
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
// BidDocument (SharePoint reference — not a Dataverse table, stored as JSON
// in a multiline text column on BidWorkspace or as a separate note record)
// ---------------------------------------------------------------------------

export interface BidDocument {
  id: string;
  cr5ab_title: string;
  cr5ab_documenttype: string;
  cr5ab_sharepointurl: string;
  cr5ab_uploadedby: DataverseUser;
  cr5ab_filesize?: number;
  cr5ab_version: string;
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
  messageType: "assignment" | "approval_request" | "deadline_reminder" | "status_change";
  additionalContext?: Record<string, string>;
}
