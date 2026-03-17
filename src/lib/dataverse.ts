/**
 * Dataverse client
 *
 * Uses @microsoft/power-apps/data to query the real Dataverse tables when
 * running inside the Power Apps host (production / Local Play).
 *
 * Falls back to in-memory mock data when running as a plain Vite dev server
 * (i.e. the localhost URL, not the Local Play URL).
 */

import { getClient } from "@microsoft/power-apps/data";
import type { DataClient } from "@microsoft/power-apps/data";
import type {
  BidRequest,
  BidWorkspace,
  BidType,
  BidRoleAssignment,
  BidApproval,
  BidStatusRecord,
  TorItem,
  BidDocumentRecord,
  BidClarification,
  PagedResult,
  DataverseQueryOptions,
  RouteBidFlowPayload,
  NotifyTeamFlowPayload,
} from "../types/dataverse";
import {
  BidStatus,
  BidTypeCode,
  TeamRole,
  ApprovalStatus,
  TorAnsweredStatus,
  OpportunityStage,
  BidSource,
  QualificationOutcome,
  ClarificationStatus,
} from "../types/dataverse";

// ---------------------------------------------------------------------------
// DataSourcesInfo — maps our table aliases to the Dataverse logical names
// registered in power.config.json under databaseReferences["default.cds"]
// ---------------------------------------------------------------------------

const dataSourcesInfo = {
  cr5ab_bidrequest: {
    tableId: "cr5ab_bidrequests",
    apis: {
      getItems:   { path: "/cr5ab_bidrequests",      method: "GET",   parameters: [] },
      getItem:    { path: "/cr5ab_bidrequests/{id}", method: "GET",   parameters: [{ name: "id", in: "path", required: true, type: "string" }] },
      createItem: { path: "/cr5ab_bidrequests",      method: "POST",  parameters: [] },
      updateItem: { path: "/cr5ab_bidrequests/{id}", method: "PATCH", parameters: [{ name: "id", in: "path", required: true, type: "string" }] },
    },
  },
  cr5ab_bidworkspace: {
    tableId: "cr5ab_bidworkspaces",
    apis: {
      getItems:   { path: "/cr5ab_bidworkspaces",      method: "GET",   parameters: [] },
      getItem:    { path: "/cr5ab_bidworkspaces/{id}", method: "GET",   parameters: [{ name: "id", in: "path", required: true, type: "string" }] },
      createItem: { path: "/cr5ab_bidworkspaces",      method: "POST",  parameters: [] },
      updateItem: { path: "/cr5ab_bidworkspaces/{id}", method: "PATCH", parameters: [{ name: "id", in: "path", required: true, type: "string" }] },
    },
  },
  cr5ab_bidtype: {
    tableId: "cr5ab_bidtypes",
    apis: {
      getItems: { path: "/cr5ab_bidtypes", method: "GET", parameters: [] },
    },
  },
  cr5ab_approval: {
    tableId: "cr5ab_approvals",
    apis: {
      getItems:   { path: "/cr5ab_approvals",      method: "GET",   parameters: [] },
      createItem: { path: "/cr5ab_approvals",      method: "POST",  parameters: [] },
      updateItem: { path: "/cr5ab_approvals/{id}", method: "PATCH", parameters: [{ name: "id", in: "path", required: true, type: "string" }] },
    },
  },
  cr5ab_bidroleassignment: {
    tableId: "cr5ab_bidroleassignments",
    apis: {
      getItems:   { path: "/cr5ab_bidroleassignments", method: "GET",  parameters: [] },
      createItem: { path: "/cr5ab_bidroleassignments", method: "POST", parameters: [] },
    },
  },
  cr5ab_bidstatus: {
    tableId: "cr5ab_bidstatuses",
    apis: {
      getItems: { path: "/cr5ab_bidstatuses", method: "GET", parameters: [] },
    },
  },
  cr5ab_toritem: {
    tableId: "cr5ab_toritems",
    apis: {
      getItems:   { path: "/cr5ab_toritems",      method: "GET",    parameters: [] },
      getItem:    { path: "/cr5ab_toritems/{id}", method: "GET",    parameters: [{ name: "id", in: "path", required: true, type: "string" }] },
      createItem: { path: "/cr5ab_toritems",      method: "POST",   parameters: [] },
      updateItem: { path: "/cr5ab_toritems/{id}", method: "PATCH",  parameters: [{ name: "id", in: "path", required: true, type: "string" }] },
      deleteItem: { path: "/cr5ab_toritems/{id}", method: "DELETE", parameters: [{ name: "id", in: "path", required: true, type: "string" }] },
    },
  },
  cr5ab_clarification: {
    tableId: "cr5ab_clarifications",
    apis: {
      getItems:   { path: "/cr5ab_clarifications",      method: "GET",   parameters: [] },
      createItem: { path: "/cr5ab_clarifications",      method: "POST",  parameters: [] },
      updateItem: { path: "/cr5ab_clarifications/{id}", method: "PATCH", parameters: [{ name: "id", in: "path", required: true, type: "string" }] },
    },
  },
};

// ---------------------------------------------------------------------------
// SDK client instance
// ---------------------------------------------------------------------------

let _client: DataClient | null = null;

function getDataClient(): DataClient {
  if (!_client) {
    _client = getClient(dataSourcesInfo);
  }
  return _client;
}

// ---------------------------------------------------------------------------
// Dev-mode mock data
// ---------------------------------------------------------------------------

const IS_DEV = import.meta.env.DEV;

function nowIso() {
  return new Date().toISOString();
}

const SYSTEM_USER = {
  id: "00000000-0000-0000-0000-000000000001",
  fullName: "Dev User",
  email: "dev@ricoh.co.uk",
};

const USERS = {
  sarah:  { id: "u1", fullName: "Sarah Mitchell",  email: "s.mitchell@ricoh.co.uk" },
  james:  { id: "u2", fullName: "James O'Brien",   email: "j.obrien@ricoh.co.uk" },
  priya:  { id: "u3", fullName: "Priya Sharma",    email: "p.sharma@ricoh.co.uk" },
  tom:    { id: "u4", fullName: "Tom Watkins",     email: "t.watkins@ricoh.co.uk" },
  helen:  { id: "u5", fullName: "Helen Cross",     email: "h.cross@ricoh.co.uk" },
  louise: { id: "u6", fullName: "Louise Brennan",  email: "l.brennan@ricoh.co.uk" },
};

export const MOCK_BID_TYPES: BidType[] = [
  { id: "bt-001", cr5ab_code: BidTypeCode.SupplierQuestionnaire, cr5ab_name: "Supplier Questionnaire", cr5ab_description: "Pre-qualification questionnaire submitted by potential suppliers.", cr5ab_routingteam: "Procurement", cr5ab_routingemail: "procurement@ricoh.co.uk", cr5ab_slaresponsedays: 5, createdOn: nowIso(), modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
  { id: "bt-002", cr5ab_code: BidTypeCode.SalesLed,              cr5ab_name: "Sales Led",              cr5ab_description: "Bids driven by the Sales team in response to an RFP.",              cr5ab_routingteam: "Sales Bids",    cr5ab_routingemail: "salesbids@ricoh.co.uk",    cr5ab_slaresponsedays: 3, createdOn: nowIso(), modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
  { id: "bt-003", cr5ab_code: BidTypeCode.SMEToQualify,          cr5ab_name: "SME to Qualify",         cr5ab_description: "Opportunities routed to an SME for qualification.",                 cr5ab_routingteam: "SME Pool",      cr5ab_routingemail: "sme@ricoh.co.uk",          cr5ab_slaresponsedays: 7, createdOn: nowIso(), modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
  { id: "bt-004", cr5ab_code: BidTypeCode.BidManagement,         cr5ab_name: "Bid Management",         cr5ab_description: "Full bid management lifecycle handled by the Bid Management team.", cr5ab_routingteam: "Bid Management", cr5ab_routingemail: "bidmanagement@ricoh.co.uk", cr5ab_slaresponsedays: 2, createdOn: nowIso(), modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
];

export const MOCK_BID_REQUESTS: BidRequest[] = [
  {
    id: "br-001",
    cr5ab_bidreferencenumber: "BID-2026-00001",
    cr5ab_title: "NHS Digital Printing Framework",
    cr5ab_bidtypeid: { id: "bt-004", cr5ab_name: "Bid Management", cr5ab_code: BidTypeCode.BidManagement },
    cr5ab_status: BidStatus.InProgress,
    cr5ab_opportunitystage: OpportunityStage.Allocated,
    cr5ab_source: BidSource.Portal,
    cr5ab_sourceportalname: "Find a Tender",
    cr5ab_customername: "NHS England",
    cr5ab_customerindustry: "Healthcare",
    cr5ab_estimatedvalue: 4500000,
    cr5ab_currency: "GBP",
    cr5ab_submissiondeadline: "2026-04-15T17:00:00.000Z",
    cr5ab_expectedawarddate: "2026-05-30T00:00:00.000Z",
    cr5ab_description: "National framework for managed print services across NHS trusts.",
    cr5ab_submittedby: SYSTEM_USER,
    cr5ab_assignedto: USERS.sarah,
    cr5ab_routedto: "Bid Management",
    cr5ab_qualificationoutcome: QualificationOutcome.BidManagement,
    cr5ab_qualificationrationale: "Strong existing NHS relationships. Value above threshold. Full bid management required.",
    cr5ab_qualifiedby: USERS.helen,
    cr5ab_qualifiedon: "2026-02-14T11:00:00.000Z",
    cr5ab_oafdata: JSON.stringify({
      servicesInScope: "Managed print services, device management, consumables, service desk. All major NHS trusts in England.",
      keyRisks: "Complex procurement framework requirements. Pricing schedule format is strict. Tight submission timeline.",
      hasIncumbent: true, incumbentName: "Xerox",
      relationshipsInPlace: true,
      estimatedBidEffortDays: "15",
      recommendedResource: "Sarah Mitchell (Bid Manager), James O'Brien (Technical SME)",
      additionalNotes: "Helen to review pricing model before submission.",
    }),
    cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "NHS Digital Printing Framework — Workspace" },
    createdOn: "2026-02-10T09:00:00.000Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "br-002",
    cr5ab_bidreferencenumber: "BID-2026-00002",
    cr5ab_title: "Central Government MFD Refresh",
    cr5ab_bidtypeid: { id: "bt-002", cr5ab_name: "Sales Led", cr5ab_code: BidTypeCode.SalesLed },
    cr5ab_status: BidStatus.Submitted,
    cr5ab_opportunitystage: OpportunityStage.Allocated,
    cr5ab_source: BidSource.SalesSubmitted,
    cr5ab_customername: "HMRC",
    cr5ab_customerindustry: "Government",
    cr5ab_estimatedvalue: 1200000,
    cr5ab_currency: "GBP",
    cr5ab_submissiondeadline: "2026-03-30T17:00:00.000Z",
    cr5ab_description: "Replacement of end-of-life MFD estate across HMRC offices.",
    cr5ab_submittedby: SYSTEM_USER,
    cr5ab_assignedto: USERS.tom,
    cr5ab_qualificationoutcome: QualificationOutcome.SalesLed,
    cr5ab_qualificationrationale: "Sales team have existing relationship. Light-touch bid support only.",
    cr5ab_qualifiedby: USERS.helen,
    cr5ab_qualifiedon: "2026-02-22T10:00:00.000Z",
    cr5ab_bidworkspaceid: { id: "ws-002", cr5ab_title: "Central Government MFD Refresh — Workspace" },
    createdOn: "2026-02-20T11:30:00.000Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "br-003",
    cr5ab_bidreferencenumber: "BID-2026-00003",
    cr5ab_title: "University Print Supplier PQQ",
    cr5ab_bidtypeid: { id: "bt-001", cr5ab_name: "Supplier Questionnaire", cr5ab_code: BidTypeCode.SupplierQuestionnaire },
    cr5ab_status: BidStatus.InReview,
    cr5ab_opportunitystage: OpportunityStage.Allocated,
    cr5ab_source: BidSource.Portal,
    cr5ab_sourceportalname: "Contracts Finder",
    cr5ab_customername: "University of Manchester",
    cr5ab_customerindustry: "Education",
    cr5ab_estimatedvalue: 350000,
    cr5ab_currency: "GBP",
    cr5ab_submissiondeadline: "2026-03-20T17:00:00.000Z",
    cr5ab_description: "Pre-qualification questionnaire for university print supplier list.",
    cr5ab_submittedby: SYSTEM_USER,
    cr5ab_assignedto: USERS.james,
    cr5ab_qualificationoutcome: QualificationOutcome.BidManagement,
    cr5ab_qualifiedby: USERS.helen,
    cr5ab_qualifiedon: "2026-03-03T09:00:00.000Z",
    cr5ab_bidworkspaceid: { id: "ws-003", cr5ab_title: "University Print Supplier PQQ — Workspace" },
    createdOn: "2026-03-01T08:00:00.000Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "br-004",
    cr5ab_bidreferencenumber: "BID-2026-00004",
    cr5ab_title: "Retail Chain Document Management",
    cr5ab_bidtypeid: { id: "bt-003", cr5ab_name: "SME to Qualify", cr5ab_code: BidTypeCode.SMEToQualify },
    cr5ab_status: BidStatus.InReview,
    cr5ab_opportunitystage: OpportunityStage.UnderQualification,
    cr5ab_source: BidSource.SalesSubmitted,
    cr5ab_customername: "Tesco PLC",
    cr5ab_customerindustry: "Retail",
    cr5ab_estimatedvalue: 780000,
    cr5ab_currency: "GBP",
    cr5ab_submissiondeadline: "2026-04-30T17:00:00.000Z",
    cr5ab_description: "Document management and workflow automation for retail operations.",
    cr5ab_submittedby: SYSTEM_USER,
    createdOn: "2026-03-10T14:00:00.000Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "br-005",
    cr5ab_bidreferencenumber: "BID-2025-00041",
    cr5ab_title: "Local Authority Print Contract",
    cr5ab_bidtypeid: { id: "bt-004", cr5ab_name: "Bid Management", cr5ab_code: BidTypeCode.BidManagement },
    cr5ab_status: BidStatus.Won,
    cr5ab_opportunitystage: OpportunityStage.Allocated,
    cr5ab_source: BidSource.Portal,
    cr5ab_sourceportalname: "Contracts Finder",
    cr5ab_customername: "Birmingham City Council",
    cr5ab_customerindustry: "Government",
    cr5ab_estimatedvalue: 2100000,
    cr5ab_currency: "GBP",
    cr5ab_submissiondeadline: "2025-11-30T17:00:00.000Z",
    cr5ab_description: "5-year managed print services contract for Birmingham City Council.",
    cr5ab_submittedby: SYSTEM_USER,
    cr5ab_assignedto: USERS.louise,
    cr5ab_qualificationoutcome: QualificationOutcome.BidManagement,
    cr5ab_qualifiedby: USERS.helen,
    cr5ab_qualifiedon: "2025-09-20T10:00:00.000Z",
    cr5ab_bidworkspaceid: { id: "ws-005", cr5ab_title: "Local Authority Print Contract — Workspace" },
    createdOn: "2025-09-15T09:00:00.000Z", modifiedOn: "2025-12-10T09:00:00.000Z", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "br-006",
    cr5ab_bidreferencenumber: "BID-2026-00005",
    cr5ab_title: "County Council Managed Print Services",
    cr5ab_bidtypeid: { id: "bt-004", cr5ab_name: "Bid Management", cr5ab_code: BidTypeCode.BidManagement },
    cr5ab_status: BidStatus.Qualified,
    cr5ab_opportunitystage: OpportunityStage.Allocated,
    cr5ab_source: BidSource.Portal,
    cr5ab_sourceportalname: "Find a Tender",
    cr5ab_customername: "Surrey County Council",
    cr5ab_customerindustry: "Government",
    cr5ab_estimatedvalue: 980000,
    cr5ab_currency: "GBP",
    cr5ab_submissiondeadline: "2026-05-10T17:00:00.000Z",
    cr5ab_description: "Managed print services across council offices and libraries.",
    cr5ab_submittedby: SYSTEM_USER,
    cr5ab_assignedto: USERS.helen,
    cr5ab_qualificationoutcome: QualificationOutcome.BidManagement,
    cr5ab_qualificationrationale: "Good fit with our framework experience. Needs full bid management.",
    cr5ab_qualifiedby: USERS.helen,
    cr5ab_qualifiedon: "2026-03-08T14:00:00.000Z",
    cr5ab_oafdata: JSON.stringify({
      servicesInScope: "Managed print, fleet management, consumables supply, on-site support.",
      keyRisks: "Council procurement process may extend timeline. Social value scoring weighted heavily.",
      hasIncumbent: false, incumbentName: "",
      relationshipsInPlace: false,
      estimatedBidEffortDays: "10",
      recommendedResource: "Helen Cross (Bid Manager)",
      additionalNotes: "Social value section will need strong local examples.",
    }),
    cr5ab_bidworkspaceid: { id: "ws-004", cr5ab_title: "County Council Managed Print Services — Workspace" },
    createdOn: "2026-03-05T09:00:00.000Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "br-007",
    cr5ab_bidreferencenumber: "BID-2026-00006",
    cr5ab_title: "Ireland MFD Supply & Maintenance",
    cr5ab_bidtypeid: { id: "bt-004", cr5ab_name: "Bid Management", cr5ab_code: BidTypeCode.BidManagement },
    cr5ab_status: BidStatus.InReview,
    cr5ab_opportunitystage: OpportunityStage.PendingOAF,
    cr5ab_source: BidSource.SalesSubmitted,
    cr5ab_customername: "An Post (Ireland)",
    cr5ab_customerindustry: "Government",
    cr5ab_estimatedvalue: 450000,
    cr5ab_currency: "EUR",
    cr5ab_submissiondeadline: "2026-03-30T17:00:00.000Z",
    cr5ab_description: "Supply and maintenance of MFD fleet for An Post offices across Ireland.",
    cr5ab_submittedby: SYSTEM_USER,
    createdOn: "2026-03-12T09:00:00.000Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
];

// Mock team members per workspace
const MOCK_TEAM_WS001: BidRoleAssignment[] = [
  { id: "tm-001-1", cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" }, cr5ab_userid: USERS.sarah,  cr5ab_role: TeamRole.BidManager,     cr5ab_isactive: true, cr5ab_assigneddate: "2026-02-12T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
  { id: "tm-001-2", cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" }, cr5ab_userid: USERS.james,  cr5ab_role: TeamRole.SME,             cr5ab_isactive: true, cr5ab_assigneddate: "2026-02-13T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
  { id: "tm-001-3", cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" }, cr5ab_userid: USERS.priya,  cr5ab_role: TeamRole.Approver,        cr5ab_isactive: true, cr5ab_assigneddate: "2026-02-14T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
  { id: "tm-001-4", cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" }, cr5ab_userid: USERS.tom,    cr5ab_role: TeamRole.BidCoordinator,  cr5ab_isactive: true, cr5ab_assigneddate: "2026-02-15T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
];

const MOCK_APPROVALS_WS001: BidApproval[] = [
  { id: "ap-001-1", cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" }, cr5ab_title: "Technical Review",   cr5ab_approverstage: 1, cr5ab_approverid: USERS.priya,  cr5ab_status: ApprovalStatus.Approved, cr5ab_requesteddate: "2026-02-20T09:00:00Z", cr5ab_respondeddate: "2026-02-22T14:00:00Z", cr5ab_comments: "Technically sound. Proceed.", createdOn: "", modifiedOn: "", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
  { id: "ap-001-2", cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" }, cr5ab_title: "Commercial Review", cr5ab_approverstage: 2, cr5ab_approverid: USERS.tom,    cr5ab_status: ApprovalStatus.Pending,  cr5ab_requesteddate: "2026-02-23T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
  { id: "ap-001-3", cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" }, cr5ab_title: "Executive Sign-off",cr5ab_approverstage: 3, cr5ab_approverid: USERS.helen,  cr5ab_status: ApprovalStatus.Pending,  cr5ab_requesteddate: "2026-02-23T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
];

const MOCK_DOCS_WS001: BidDocumentRecord[] = [
  { id: "doc-001-1", cr5ab_title: "ITT Document",            cr5ab_filename: "ITT_NHS_2026.pdf",       cr5ab_documenttype: "Tender Document",  cr5ab_category: "originals",  cr5ab_sharepointurl: "https://ricoh.sharepoint.com/sites/bids/BID-2026-00001/ITT.pdf",     cr5ab_version: "1.0", cr5ab_uploadedby: USERS.sarah,  cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" }, createdOn: "2026-02-12T09:00:00Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
  { id: "doc-001-2", cr5ab_title: "Technical Response Draft",cr5ab_filename: "Tech_Response_v2.docx",  cr5ab_documenttype: "Response",          cr5ab_category: "working",    cr5ab_sharepointurl: "https://ricoh.sharepoint.com/sites/bids/BID-2026-00001/TechResp.docx",cr5ab_version: "2.1", cr5ab_uploadedby: USERS.james,  cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" }, createdOn: "2026-02-20T09:00:00Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
  { id: "doc-001-3", cr5ab_title: "Executive Summary",       cr5ab_filename: "Exec_Summary.docx",      cr5ab_documenttype: "Summary",           cr5ab_category: "submission", cr5ab_sharepointurl: "https://ricoh.sharepoint.com/sites/bids/BID-2026-00001/ExecSum.docx", cr5ab_version: "1.0", cr5ab_uploadedby: USERS.sarah,  cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" }, createdOn: "2026-03-01T09:00:00Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
];

export const MOCK_WORKSPACES: BidWorkspace[] = [
  {
    id: "ws-001",
    cr5ab_title: "NHS Digital Printing Framework — Workspace",
    cr5ab_bidrequestid: { id: "br-001", cr5ab_title: "NHS Digital Printing Framework", cr5ab_bidreferencenumber: "BID-2026-00001" },
    cr5ab_status: BidStatus.InProgress,
    cr5ab_bidmanagerid: USERS.sarah,
    cr5ab_completionpercentage: 45,
    cr5ab_sharepointfolderurl: "https://ricoh.sharepoint.com/sites/bids/BID-2026-00001",
    cr5ab_teamschannelurl: "https://teams.microsoft.com/l/channel/NHS-Digital",
    teamMembers: MOCK_TEAM_WS001,
    approvals: MOCK_APPROVALS_WS001,
    documents: MOCK_DOCS_WS001,
    clarifications: [], // populated after MOCK_CLARIFICATIONS is declared below
    createdOn: "2026-02-12T09:00:00.000Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "ws-002",
    cr5ab_title: "Central Government MFD Refresh — Workspace",
    cr5ab_bidrequestid: { id: "br-002", cr5ab_title: "Central Government MFD Refresh", cr5ab_bidreferencenumber: "BID-2026-00002" },
    cr5ab_status: BidStatus.Submitted,
    cr5ab_bidmanagerid: USERS.tom,
    cr5ab_completionpercentage: 20,
    cr5ab_sharepointfolderurl: "https://ricoh.sharepoint.com/sites/bids/BID-2026-00002",
    teamMembers: [
      { id: "tm-002-1", cr5ab_bidworkspaceid: { id: "ws-002", cr5ab_title: "" }, cr5ab_userid: USERS.tom,    cr5ab_role: TeamRole.BidManager,    cr5ab_isactive: true, cr5ab_assigneddate: "2026-02-21T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
      { id: "tm-002-2", cr5ab_bidworkspaceid: { id: "ws-002", cr5ab_title: "" }, cr5ab_userid: USERS.louise, cr5ab_role: TeamRole.BidCoordinator, cr5ab_isactive: true, cr5ab_assigneddate: "2026-02-21T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
    ],
    approvals: [],
    documents: [],
    createdOn: "2026-02-21T09:00:00.000Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "ws-003",
    cr5ab_title: "University Print Supplier PQQ — Workspace",
    cr5ab_bidrequestid: { id: "br-003", cr5ab_title: "University Print Supplier PQQ", cr5ab_bidreferencenumber: "BID-2026-00003" },
    cr5ab_status: BidStatus.InReview,
    cr5ab_bidmanagerid: USERS.james,
    cr5ab_completionpercentage: 65,
    cr5ab_sharepointfolderurl: "https://ricoh.sharepoint.com/sites/bids/BID-2026-00003",
    teamMembers: [
      { id: "tm-003-1", cr5ab_bidworkspaceid: { id: "ws-003", cr5ab_title: "" }, cr5ab_userid: USERS.james,  cr5ab_role: TeamRole.BidManager, cr5ab_isactive: true, cr5ab_assigneddate: "2026-03-02T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
      { id: "tm-003-2", cr5ab_bidworkspaceid: { id: "ws-003", cr5ab_title: "" }, cr5ab_userid: USERS.priya,  cr5ab_role: TeamRole.Reviewer,   cr5ab_isactive: true, cr5ab_assigneddate: "2026-03-02T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
    ],
    approvals: [
      { id: "ap-003-1", cr5ab_bidworkspaceid: { id: "ws-003", cr5ab_title: "" }, cr5ab_title: "Technical Review", cr5ab_approverstage: 1, cr5ab_approverid: USERS.priya, cr5ab_status: ApprovalStatus.Pending, cr5ab_requesteddate: "2026-03-10T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
    ],
    documents: [],
    createdOn: "2026-03-02T09:00:00.000Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "ws-004",
    cr5ab_title: "County Council Managed Print Services — Workspace",
    cr5ab_bidrequestid: { id: "br-006", cr5ab_title: "County Council Managed Print Services", cr5ab_bidreferencenumber: "BID-2026-00005" },
    cr5ab_status: BidStatus.Qualified,
    cr5ab_bidmanagerid: USERS.helen,
    cr5ab_completionpercentage: 30,
    cr5ab_sharepointfolderurl: "https://ricoh.sharepoint.com/sites/bids/BID-2026-00005",
    teamMembers: [
      { id: "tm-004-1", cr5ab_bidworkspaceid: { id: "ws-004", cr5ab_title: "" }, cr5ab_userid: USERS.helen, cr5ab_role: TeamRole.BidManager, cr5ab_isactive: true, cr5ab_assigneddate: "2026-03-06T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
    ],
    approvals: [],
    documents: [],
    createdOn: "2026-03-06T09:00:00.000Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "ws-005",
    cr5ab_title: "Local Authority Print Contract — Workspace",
    cr5ab_bidrequestid: { id: "br-005", cr5ab_title: "Local Authority Print Contract", cr5ab_bidreferencenumber: "BID-2025-00041" },
    cr5ab_status: BidStatus.Won,
    cr5ab_bidmanagerid: USERS.louise,
    cr5ab_completionpercentage: 100,
    cr5ab_sharepointfolderurl: "https://ricoh.sharepoint.com/sites/bids/BID-2025-00041",
    teamMembers: [
      { id: "tm-005-1", cr5ab_bidworkspaceid: { id: "ws-005", cr5ab_title: "" }, cr5ab_userid: USERS.louise, cr5ab_role: TeamRole.BidManager, cr5ab_isactive: true, cr5ab_assigneddate: "2025-09-16T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER },
    ],
    approvals: [],
    documents: [],
    createdOn: "2025-09-16T09:00:00.000Z", modifiedOn: "2025-12-10T09:00:00.000Z", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
];

// Mock TOR items — attached to ws-001
export let MOCK_TOR_ITEMS: TorItem[] = [
  {
    id: "tor-001-01", cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" },
    cr5ab_section: "Section 1 — Company Information",
    cr5ab_questionnumber: "1.1",
    cr5ab_questiondetail: "Provide a brief company overview including history, ownership structure, and key statistics (turnover, headcount).",
    cr5ab_allocatedto: USERS.sarah, cr5ab_department: "Bid Management",
    cr5ab_firstdraftdeadline: "2026-03-20T17:00:00Z", cr5ab_finaldraftdeadline: "2026-03-28T17:00:00Z", cr5ab_actualdeadline: "2026-04-10T17:00:00Z",
    cr5ab_answeredstatus: TorAnsweredStatus.Yes,
    cr5ab_comments: "Completed using standard company overview template.",
    cr5ab_scoreweightingthreshold: "Pass/Fail",
    createdOn: "2026-02-15T09:00:00Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "tor-001-02", cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" },
    cr5ab_section: "Section 1 — Company Information",
    cr5ab_questionnumber: "1.2",
    cr5ab_questiondetail: "Provide evidence of financial stability including audited accounts for the last 3 years.",
    cr5ab_allocatedto: USERS.tom, cr5ab_department: "Commercial",
    cr5ab_firstdraftdeadline: "2026-03-20T17:00:00Z", cr5ab_finaldraftdeadline: "2026-03-28T17:00:00Z", cr5ab_actualdeadline: "2026-04-10T17:00:00Z",
    cr5ab_answeredstatus: TorAnsweredStatus.Yes,
    cr5ab_scoreweightingthreshold: "Pass/Fail",
    createdOn: "2026-02-15T09:00:00Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "tor-001-03", cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" },
    cr5ab_section: "Section 2 — Technical Capability",
    cr5ab_questionnumber: "2.1",
    cr5ab_questiondetail: "Describe your managed print service offering, including device management, consumables, and service desk capabilities.",
    cr5ab_allocatedto: USERS.james, cr5ab_department: "Technical",
    cr5ab_firstdraftdeadline: "2026-03-22T17:00:00Z", cr5ab_finaldraftdeadline: "2026-04-01T17:00:00Z", cr5ab_actualdeadline: "2026-04-10T17:00:00Z",
    cr5ab_answeredstatus: TorAnsweredStatus.Partial,
    cr5ab_comments: "Draft in progress — waiting on service desk SLA data from James.",
    cr5ab_scoreweightingthreshold: "20%",
    cr5ab_specialinstructions: "Max 1,000 words. Include case studies.",
    createdOn: "2026-02-15T09:00:00Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "tor-001-04", cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" },
    cr5ab_section: "Section 2 — Technical Capability",
    cr5ab_questionnumber: "2.2",
    cr5ab_questiondetail: "Provide details of your security accreditations and data handling policies (ISO 27001, Cyber Essentials Plus).",
    cr5ab_allocatedto: USERS.james, cr5ab_department: "Technical",
    cr5ab_firstdraftdeadline: "2026-03-22T17:00:00Z", cr5ab_finaldraftdeadline: "2026-04-01T17:00:00Z", cr5ab_actualdeadline: "2026-04-10T17:00:00Z",
    cr5ab_answeredstatus: TorAnsweredStatus.No,
    cr5ab_scoreweightingthreshold: "15%",
    createdOn: "2026-02-15T09:00:00Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "tor-001-05", cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" },
    cr5ab_section: "Section 3 — Pricing & Commercial",
    cr5ab_questionnumber: "3.1",
    cr5ab_questiondetail: "Complete the attached pricing schedule for all device types and service components.",
    cr5ab_allocatedto: USERS.priya, cr5ab_department: "Commercial",
    cr5ab_firstdraftdeadline: "2026-04-01T17:00:00Z", cr5ab_finaldraftdeadline: "2026-04-08T17:00:00Z", cr5ab_actualdeadline: "2026-04-12T17:00:00Z",
    cr5ab_answeredstatus: TorAnsweredStatus.No,
    cr5ab_scoreweightingthreshold: "40%",
    cr5ab_specialinstructions: "Use NHS pricing schedule template v3. Do not deviate from format.",
    createdOn: "2026-02-15T09:00:00Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "tor-001-06", cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" },
    cr5ab_section: "Section 3 — Pricing & Commercial",
    cr5ab_questionnumber: "3.2",
    cr5ab_questiondetail: "Describe your approach to cost transparency and reporting for framework call-offs.",
    cr5ab_allocatedto: USERS.priya, cr5ab_department: "Commercial",
    cr5ab_firstdraftdeadline: "2026-04-01T17:00:00Z", cr5ab_finaldraftdeadline: "2026-04-08T17:00:00Z", cr5ab_actualdeadline: "2026-04-12T17:00:00Z",
    cr5ab_answeredstatus: TorAnsweredStatus.No,
    cr5ab_scoreweightingthreshold: "25%",
    createdOn: "2026-02-15T09:00:00Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "tor-001-07", cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" },
    cr5ab_section: "Section 4 — Social Value",
    cr5ab_questionnumber: "4.1",
    cr5ab_questiondetail: "Describe your commitments to social value under PPN 06/20, with specific local initiatives for NHS regions.",
    cr5ab_allocatedto: USERS.sarah, cr5ab_department: "Bid Management",
    cr5ab_firstdraftdeadline: "2026-03-25T17:00:00Z", cr5ab_finaldraftdeadline: "2026-04-03T17:00:00Z", cr5ab_actualdeadline: "2026-04-10T17:00:00Z",
    cr5ab_answeredstatus: TorAnsweredStatus.Partial,
    cr5ab_comments: "Using standard social value narrative — needs NHS-specific examples added.",
    cr5ab_scoreweightingthreshold: "Pass/Fail",
    createdOn: "2026-02-15T09:00:00Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
];

// Mock clarifications — attached to ws-001
export let MOCK_CLARIFICATIONS: BidClarification[] = [
  {
    id: "cq-001-01",
    cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" },
    cr5ab_questionnumber: "CQ-01",
    cr5ab_questiontext: "Can you confirm whether Lot 2 (Nationwide Framework) requires a separate accreditation to Lot 1, or whether a single ISO 9001 certificate covers both lots?",
    cr5ab_raisedby: USERS.sarah,
    cr5ab_raiseddate: "2026-02-28T10:00:00Z",
    cr5ab_deadline: "2026-03-14T17:00:00Z",
    cr5ab_responsetext: "A single ISO 9001 certificate will cover both Lot 1 and Lot 2 provided it explicitly references managed print services.",
    cr5ab_respondeddate: "2026-03-05T14:30:00Z",
    cr5ab_status: ClarificationStatus.Closed,
    cr5ab_iscustomerraised: false,
    createdOn: "2026-02-28T10:00:00Z", modifiedOn: "2026-03-05T14:30:00Z", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "cq-001-02",
    cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" },
    cr5ab_questionnumber: "CQ-02",
    cr5ab_questiontext: "Please clarify the expected format for the pricing schedule — specifically whether day-rate adjustments should be expressed as percentage uplifts or absolute values in columns F–H.",
    cr5ab_raisedby: USERS.priya,
    cr5ab_raiseddate: "2026-03-10T09:30:00Z",
    cr5ab_deadline: "2026-03-24T17:00:00Z",
    cr5ab_status: ClarificationStatus.Open,
    cr5ab_iscustomerraised: false,
    createdOn: "2026-03-10T09:30:00Z", modifiedOn: "2026-03-10T09:30:00Z", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "cq-001-03",
    cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" },
    cr5ab_questionnumber: "CQ-03",
    cr5ab_questiontext: "The ITT references Annex 7 (Environmental Policy Template) but this document was not included in the procurement pack. Can you confirm when this will be issued?",
    cr5ab_raisedby: USERS.james,
    cr5ab_raiseddate: "2026-03-14T11:00:00Z",
    cr5ab_deadline: "2026-03-28T17:00:00Z",
    cr5ab_responsetext: "Annex 7 has now been uploaded to the portal. Please re-download the procurement pack.",
    cr5ab_respondeddate: "2026-03-15T09:00:00Z",
    cr5ab_status: ClarificationStatus.AnswerReceived,
    cr5ab_iscustomerraised: false,
    createdOn: "2026-03-14T11:00:00Z", modifiedOn: "2026-03-15T09:00:00Z", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "cq-001-04",
    cr5ab_bidworkspaceid: { id: "ws-001", cr5ab_title: "" },
    cr5ab_questionnumber: "CQ-04",
    cr5ab_questiontext: "Please confirm whether your company holds a current NHS DSP Toolkit submission and provide the submission reference number.",
    cr5ab_raisedby: SYSTEM_USER,
    cr5ab_raiseddate: "2026-03-16T14:00:00Z",
    cr5ab_deadline: "2026-03-22T17:00:00Z",
    cr5ab_status: ClarificationStatus.Open,
    cr5ab_iscustomerraised: true,
    createdOn: "2026-03-16T14:00:00Z", modifiedOn: "2026-03-16T14:00:00Z", createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
];

// Back-fill workspace ws-001 clarifications now that the array is declared
MOCK_WORKSPACES[0].clarifications = MOCK_CLARIFICATIONS.filter((c) => c.cr5ab_bidworkspaceid.id === "ws-001");

// Mock TOR for ws-002
export let MOCK_TOR_ITEMS_WS002: TorItem[] = [
  {
    id: "tor-002-01", cr5ab_bidworkspaceid: { id: "ws-002", cr5ab_title: "" },
    cr5ab_section: "Section 1 — Supplier Information",
    cr5ab_questionnumber: "1.1",
    cr5ab_questiondetail: "Provide company registration and VAT details.",
    cr5ab_allocatedto: USERS.tom, cr5ab_department: "Bid Management",
    cr5ab_actualdeadline: "2026-03-25T17:00:00Z",
    cr5ab_answeredstatus: TorAnsweredStatus.Yes,
    createdOn: "2026-02-22T09:00:00Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "tor-002-02", cr5ab_bidworkspaceid: { id: "ws-002", cr5ab_title: "" },
    cr5ab_section: "Section 2 — Technical",
    cr5ab_questionnumber: "2.1",
    cr5ab_questiondetail: "Detail your MFD product range and replacement cycle management.",
    cr5ab_allocatedto: USERS.james, cr5ab_department: "Technical",
    cr5ab_actualdeadline: "2026-03-25T17:00:00Z",
    cr5ab_answeredstatus: TorAnsweredStatus.No,
    createdOn: "2026-02-22T09:00:00Z", modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
];

// ---------------------------------------------------------------------------
// Mappers
// ---------------------------------------------------------------------------

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function mapBidRequest(row: any): BidRequest {
  return {
    id: row.cr5ab_bidrequestid ?? row.id ?? "",
    cr5ab_bidreferencenumber: row.cr5ab_bidreferencenumber ?? "",
    cr5ab_title: row.cr5ab_title ?? row.cr5ab_name ?? "",
    cr5ab_bidtypeid: row._cr5ab_bidtypeid_value
      ? { id: row._cr5ab_bidtypeid_value, cr5ab_name: row["_cr5ab_bidtypeid_value@OData.Community.Display.V1.FormattedValue"] ?? "", cr5ab_code: row.cr5ab_bidtypeid_code ?? BidTypeCode.BidManagement }
      : { id: "", cr5ab_name: "", cr5ab_code: BidTypeCode.BidManagement },
    cr5ab_status: row.cr5ab_status ?? BidStatus.Draft,
    cr5ab_opportunitystage: row.cr5ab_opportunitystage,
    cr5ab_source: row.cr5ab_source,
    cr5ab_sourceportalname: row.cr5ab_sourceportalname,
    cr5ab_customername: row.cr5ab_customername ?? "",
    cr5ab_customerindustry: row.cr5ab_customerindustry,
    cr5ab_estimatedvalue: row.cr5ab_estimatedvalue,
    cr5ab_currency: row.cr5ab_currency ?? "GBP",
    cr5ab_submissiondeadline: row.cr5ab_submissiondeadline ?? "",
    cr5ab_expectedawarddate: row.cr5ab_expectedawarddate,
    cr5ab_contractstartdate: row.cr5ab_contractstartdate,
    cr5ab_contractduration: row.cr5ab_contractduration,
    cr5ab_description: row.cr5ab_description ?? "",
    cr5ab_scope: row.cr5ab_scope,
    cr5ab_specialrequirements: row.cr5ab_specialrequirements,
    cr5ab_incumbentvendor: row.cr5ab_incumbentvendor,
    cr5ab_oafdata: row.cr5ab_oafdata,
    cr5ab_qualificationoutcome: row.cr5ab_qualificationoutcome,
    cr5ab_qualificationrationale: row.cr5ab_qualificationrationale,
    cr5ab_qualifiedby: row._cr5ab_qualifiedby_value
      ? { id: row._cr5ab_qualifiedby_value, fullName: row["_cr5ab_qualifiedby_value@OData.Community.Display.V1.FormattedValue"] ?? "", email: "" }
      : undefined,
    cr5ab_qualifiedon: row.cr5ab_qualifiedon,
    cr5ab_submittedby: {
      id: row._cr5ab_submittedby_value ?? "",
      fullName: row["_cr5ab_submittedby_value@OData.Community.Display.V1.FormattedValue"] ?? "",
      email: "",
    },
    cr5ab_assignedto: row._cr5ab_assignedto_value
      ? { id: row._cr5ab_assignedto_value, fullName: row["_cr5ab_assignedto_value@OData.Community.Display.V1.FormattedValue"] ?? "", email: "" }
      : undefined,
    cr5ab_routedto: row.cr5ab_routedto,
    createdOn: row.createdon ?? nowIso(),
    modifiedOn: row.modifiedon ?? nowIso(),
    createdBy: { id: row._createdby_value ?? "", fullName: row["_createdby_value@OData.Community.Display.V1.FormattedValue"] ?? "", email: "" },
    modifiedBy: { id: row._modifiedby_value ?? "", fullName: row["_modifiedby_value@OData.Community.Display.V1.FormattedValue"] ?? "", email: "" },
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function mapClarification(row: any): BidClarification {
  return {
    id: row.cr5ab_clarificationid ?? row.id ?? "",
    cr5ab_bidworkspaceid: { id: row._cr5ab_bidworkspaceid_value ?? "", cr5ab_title: "" },
    cr5ab_questionnumber: row.cr5ab_questionnumber ?? "",
    cr5ab_questiontext: row.cr5ab_questiontext ?? "",
    cr5ab_raisedby: row._cr5ab_raisedby_value
      ? { id: row._cr5ab_raisedby_value, fullName: row["_cr5ab_raisedby_value@OData.Community.Display.V1.FormattedValue"] ?? "", email: "" }
      : SYSTEM_USER,
    cr5ab_raiseddate: row.cr5ab_raiseddate ?? nowIso(),
    cr5ab_deadline: row.cr5ab_deadline,
    cr5ab_responsetext: row.cr5ab_responsetext,
    cr5ab_respondeddate: row.cr5ab_respondeddate,
    cr5ab_status: row.cr5ab_status ?? ClarificationStatus.Open,
    cr5ab_iscustomerraised: row.cr5ab_iscustomerraised ?? false,
    createdOn: row.createdon ?? nowIso(),
    modifiedOn: row.modifiedon ?? nowIso(),
    createdBy: { id: "", fullName: "", email: "" },
    modifiedBy: { id: "", fullName: "", email: "" },
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function mapBidWorkspace(row: any): BidWorkspace {
  return {
    id: row.cr5ab_bidworkspaceid ?? row.id ?? "",
    cr5ab_title: row.cr5ab_title ?? row.cr5ab_name ?? "",
    cr5ab_bidrequestid: row._cr5ab_bidrequestid_value
      ? { id: row._cr5ab_bidrequestid_value, cr5ab_title: row["_cr5ab_bidrequestid_value@OData.Community.Display.V1.FormattedValue"] ?? "", cr5ab_bidreferencenumber: "" }
      : { id: "", cr5ab_title: "", cr5ab_bidreferencenumber: "" },
    cr5ab_status: row.cr5ab_status ?? BidStatus.Draft,
    cr5ab_bidmanagerid: {
      id: row._cr5ab_bidmanagerid_value ?? "",
      fullName: row["_cr5ab_bidmanagerid_value@OData.Community.Display.V1.FormattedValue"] ?? "",
      email: "",
    },
    cr5ab_completionpercentage: row.cr5ab_completionpercentage ?? 0,
    cr5ab_sharepointfolderurl: row.cr5ab_sharepointfolderurl,
    cr5ab_teamschannelurl: row.cr5ab_teamschannelurl,
    cr5ab_wonlostdate: row.cr5ab_wonlostdate,
    cr5ab_wonlostreason: row.cr5ab_wonlostreason,
    createdOn: row.createdon ?? nowIso(),
    modifiedOn: row.modifiedon ?? nowIso(),
    createdBy: { id: row._createdby_value ?? "", fullName: "", email: "" },
    modifiedBy: { id: row._modifiedby_value ?? "", fullName: "", email: "" },
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function mapBidType(row: any): BidType {
  return {
    id: row.cr5ab_bidtypeid ?? row.id ?? "",
    cr5ab_code: row.cr5ab_code ?? BidTypeCode.BidManagement,
    cr5ab_name: row.cr5ab_name ?? "",
    cr5ab_description: row.cr5ab_description ?? "",
    cr5ab_routingteam: row.cr5ab_routingteam ?? "",
    cr5ab_routingemail: row.cr5ab_routingemail ?? "",
    cr5ab_slaresponsedays: row.cr5ab_slaresponsedays ?? 5,
    createdOn: row.createdon ?? nowIso(),
    modifiedOn: row.modifiedon ?? nowIso(),
    createdBy: { id: "", fullName: "", email: "" },
    modifiedBy: { id: "", fullName: "", email: "" },
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function mapTorItem(row: any): TorItem {
  return {
    id: row.cr5ab_toritemid ?? row.id ?? "",
    cr5ab_section: row.cr5ab_section ?? "",
    cr5ab_questionnumber: row.cr5ab_questionnumber ?? "",
    cr5ab_questiondetail: row.cr5ab_questiondetail ?? "",
    cr5ab_allocatedto: row._cr5ab_allocatedto_value
      ? { id: row._cr5ab_allocatedto_value, fullName: row["_cr5ab_allocatedto_value@OData.Community.Display.V1.FormattedValue"] ?? "", email: "" }
      : undefined,
    cr5ab_department: row.cr5ab_department,
    cr5ab_supportcontributors: row.cr5ab_supportcontributors,
    cr5ab_scoreweightingthreshold: row.cr5ab_scoreweightingthreshold,
    cr5ab_specialinstructions: row.cr5ab_specialinstructions,
    cr5ab_firstdraftdeadline: row.cr5ab_firstdraftdeadline,
    cr5ab_reviewperiod: row.cr5ab_reviewperiod,
    cr5ab_finaldraftdeadline: row.cr5ab_finaldraftdeadline,
    cr5ab_actualdeadline: row.cr5ab_actualdeadline,
    cr5ab_comments: row.cr5ab_comments,
    cr5ab_unittime: row.cr5ab_unittime,
    cr5ab_answeredstatus: row.cr5ab_answeredstatus ?? TorAnsweredStatus.No,
    cr5ab_bidworkspaceid: { id: row._cr5ab_bidworkspaceid_value ?? "", cr5ab_title: "" },
    createdOn: row.createdon ?? nowIso(),
    modifiedOn: row.modifiedon ?? nowIso(),
    createdBy: { id: "", fullName: "", email: "" },
    modifiedBy: { id: "", fullName: "", email: "" },
  };
}

// ---------------------------------------------------------------------------
// Client
// ---------------------------------------------------------------------------

class DataverseClient {

  // ------------------------------------------------------------------
  // BidType
  // ------------------------------------------------------------------

  async getBidTypes(): Promise<BidType[]> {
    if (IS_DEV) { await delay(); return MOCK_BID_TYPES; }

    const client = getDataClient();
    const result = await client.retrieveMultipleRecordsAsync<Record<string, unknown>>(
      "cr5ab_bidtype",
      { select: ["cr5ab_bidtypeid","cr5ab_name","cr5ab_code","cr5ab_description","cr5ab_routingteam","cr5ab_routingemail","cr5ab_slaresponsedays"], orderBy: ["cr5ab_code asc"] }
    );
    if (!result.success) throw new Error(result.error?.message ?? "Failed to load bid types");
    return result.data.map(mapBidType);
  }

  // ------------------------------------------------------------------
  // BidStatus
  // ------------------------------------------------------------------

  async getBidStatuses(): Promise<BidStatusRecord[]> {
    if (IS_DEV) {
      await delay();
      return []; // statuses are enum-driven; this is for admin display
    }
    const client = getDataClient();
    const result = await client.retrieveMultipleRecordsAsync<Record<string, unknown>>(
      "cr5ab_bidstatus",
      { select: ["cr5ab_bidstatusid","cr5ab_name","cr5ab_statuscode","cr5ab_description"], orderBy: ["cr5ab_statuscode asc"] }
    );
    if (!result.success) throw new Error(result.error?.message ?? "Failed to load bid statuses");
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return result.data.map((row: any) => ({
      id: row.cr5ab_bidstatusid ?? row.id ?? "",
      cr5ab_name: row.cr5ab_name ?? "",
      cr5ab_statuscode: row.cr5ab_statuscode,
      cr5ab_description: row.cr5ab_description,
      createdOn: row.createdon ?? nowIso(), modifiedOn: row.modifiedon ?? nowIso(),
      createdBy: { id: "", fullName: "", email: "" }, modifiedBy: { id: "", fullName: "", email: "" },
    }));
  }

  // ------------------------------------------------------------------
  // BidRequest
  // ------------------------------------------------------------------

  async getBidRequests(options?: DataverseQueryOptions): Promise<PagedResult<BidRequest>> {
    if (IS_DEV) { await delay(); return { value: MOCK_BID_REQUESTS, totalCount: MOCK_BID_REQUESTS.length }; }

    const client = getDataClient();
    const result = await client.retrieveMultipleRecordsAsync<Record<string, unknown>>(
      "cr5ab_bidrequest",
      {
        select: ["cr5ab_bidrequestid","cr5ab_bidreferencenumber","cr5ab_title","cr5ab_status","cr5ab_customername","cr5ab_customerindustry","cr5ab_estimatedvalue","cr5ab_currency","cr5ab_submissiondeadline","cr5ab_expectedawarddate","cr5ab_description","cr5ab_routedto","createdon","modifiedon","_cr5ab_bidtypeid_value","_cr5ab_submittedby_value","_cr5ab_assignedto_value"],
        filter: options?.filter,
        orderBy: options?.orderBy ? [options.orderBy] : ["createdon desc"],
        top: options?.top, skip: options?.skip, count: true,
      }
    );
    if (!result.success) throw new Error(result.error?.message ?? "Failed to load bid requests");
    return { value: result.data.map(mapBidRequest), totalCount: result.count };
  }

  async getBidRequest(id: string): Promise<BidRequest | null> {
    if (IS_DEV) { await delay(); return MOCK_BID_REQUESTS.find((b) => b.id === id) ?? null; }

    const client = getDataClient();
    const result = await client.retrieveRecordAsync<Record<string, unknown>>("cr5ab_bidrequest", id);
    if (!result.success) return null;
    return mapBidRequest(result.data);
  }

  async createBidRequest(
    payload: Omit<BidRequest, "id" | "createdOn" | "modifiedOn" | "createdBy" | "modifiedBy">
  ): Promise<BidRequest> {
    if (IS_DEV) {
      await delay();
      const newRecord: BidRequest = { ...payload, id: `br-${Date.now()}`, createdOn: nowIso(), modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER };
      MOCK_BID_REQUESTS.push(newRecord);
      return newRecord;
    }

    const client = getDataClient();
    const body = {
      cr5ab_title: payload.cr5ab_title, cr5ab_bidreferencenumber: payload.cr5ab_bidreferencenumber,
      cr5ab_status: payload.cr5ab_status, cr5ab_customername: payload.cr5ab_customername,
      cr5ab_customerindustry: payload.cr5ab_customerindustry, cr5ab_estimatedvalue: payload.cr5ab_estimatedvalue,
      cr5ab_currency: payload.cr5ab_currency, cr5ab_submissiondeadline: payload.cr5ab_submissiondeadline,
      cr5ab_expectedawarddate: payload.cr5ab_expectedawarddate, cr5ab_contractduration: payload.cr5ab_contractduration,
      cr5ab_description: payload.cr5ab_description, cr5ab_scope: payload.cr5ab_scope,
      cr5ab_specialrequirements: payload.cr5ab_specialrequirements, cr5ab_incumbentvendor: payload.cr5ab_incumbentvendor,
      cr5ab_routedto: payload.cr5ab_routedto,
      cr5ab_opportunitystage: payload.cr5ab_opportunitystage,
      cr5ab_source: payload.cr5ab_source,
      cr5ab_sourceportalname: payload.cr5ab_sourceportalname,
      cr5ab_oafdata: payload.cr5ab_oafdata,
      cr5ab_qualificationoutcome: payload.cr5ab_qualificationoutcome,
      cr5ab_qualificationrationale: payload.cr5ab_qualificationrationale,
      cr5ab_qualifiedon: payload.cr5ab_qualifiedon,
      "cr5ab_BidTypeId@odata.bind": payload.cr5ab_bidtypeid?.id ? `/cr5ab_bidtypes(${payload.cr5ab_bidtypeid.id})` : undefined,
    };
    const result = await client.createRecordAsync<typeof body, Record<string, unknown>>("cr5ab_bidrequest", body);
    if (!result.success) throw new Error(result.error?.message ?? "Failed to create bid request");
    return mapBidRequest(result.data);
  }

  async updateBidRequest(id: string, changes: Partial<BidRequest>): Promise<BidRequest> {
    if (IS_DEV) {
      await delay();
      const idx = MOCK_BID_REQUESTS.findIndex((b) => b.id === id);
      if (idx === -1) throw new Error(`BidRequest ${id} not found`);
      MOCK_BID_REQUESTS[idx] = { ...MOCK_BID_REQUESTS[idx], ...changes, modifiedOn: nowIso(), modifiedBy: SYSTEM_USER };
      return MOCK_BID_REQUESTS[idx];
    }
    const client = getDataClient();
    const result = await client.updateRecordAsync<Partial<BidRequest>, Record<string, unknown>>("cr5ab_bidrequest", id, changes);
    if (!result.success) throw new Error(result.error?.message ?? "Failed to update bid request");
    return mapBidRequest(result.data);
  }

  // ------------------------------------------------------------------
  // BidWorkspace
  // ------------------------------------------------------------------

  async getBidWorkspaces(options?: DataverseQueryOptions): Promise<PagedResult<BidWorkspace>> {
    if (IS_DEV) { await delay(); return { value: MOCK_WORKSPACES, totalCount: MOCK_WORKSPACES.length }; }

    const client = getDataClient();
    const result = await client.retrieveMultipleRecordsAsync<Record<string, unknown>>(
      "cr5ab_bidworkspace",
      { select: ["cr5ab_bidworkspaceid","cr5ab_title","cr5ab_status","cr5ab_completionpercentage","cr5ab_sharepointfolderurl","cr5ab_teamschannelurl","createdon","modifiedon","_cr5ab_bidrequestid_value","_cr5ab_bidmanagerid_value"], filter: options?.filter, orderBy: options?.orderBy ? [options.orderBy] : ["createdon desc"], top: options?.top, count: true }
    );
    if (!result.success) throw new Error(result.error?.message ?? "Failed to load workspaces");
    return { value: result.data.map(mapBidWorkspace), totalCount: result.count };
  }

  async getBidWorkspace(id: string): Promise<BidWorkspace | null> {
    if (IS_DEV) { await delay(); return MOCK_WORKSPACES.find((w) => w.id === id) ?? null; }
    const client = getDataClient();
    const result = await client.retrieveRecordAsync<Record<string, unknown>>("cr5ab_bidworkspace", id);
    if (!result.success) return null;
    return mapBidWorkspace(result.data);
  }

  /** Find the workspace associated with a given bid request id */
  async getWorkspaceForBid(bidRequestId: string): Promise<BidWorkspace | null> {
    if (IS_DEV) {
      await delay();
      return MOCK_WORKSPACES.find((w) => w.cr5ab_bidrequestid.id === bidRequestId) ?? null;
    }
    const client = getDataClient();
    const result = await client.retrieveMultipleRecordsAsync<Record<string, unknown>>(
      "cr5ab_bidworkspace",
      { filter: `_cr5ab_bidrequestid_value eq '${bidRequestId}'`, top: 1 }
    );
    if (!result.success || result.data.length === 0) return null;
    return mapBidWorkspace(result.data[0]);
  }

  async createBidWorkspace(
    payload: Omit<BidWorkspace, "id" | "createdOn" | "modifiedOn" | "createdBy" | "modifiedBy" | "teamMembers" | "approvals" | "documents">
  ): Promise<BidWorkspace> {
    if (IS_DEV) {
      await delay();
      const newWs: BidWorkspace = {
        ...payload, id: `ws-${Date.now()}`,
        teamMembers: [], approvals: [], documents: [],
        createdOn: nowIso(), modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
      };
      MOCK_WORKSPACES.push(newWs);
      return newWs;
    }
    const client = getDataClient();
    const body = {
      cr5ab_title: payload.cr5ab_title, cr5ab_status: payload.cr5ab_status,
      cr5ab_completionpercentage: payload.cr5ab_completionpercentage ?? 0,
      cr5ab_sharepointfolderurl: payload.cr5ab_sharepointfolderurl,
      cr5ab_teamschannelurl: payload.cr5ab_teamschannelurl,
      "cr5ab_BidRequestId@odata.bind": `/cr5ab_bidrequests(${payload.cr5ab_bidrequestid.id})`,
      "cr5ab_BidManagerId@odata.bind": `/systemusers(${payload.cr5ab_bidmanagerid.id})`,
    };
    const result = await client.createRecordAsync<typeof body, Record<string, unknown>>("cr5ab_bidworkspace", body);
    if (!result.success) throw new Error(result.error?.message ?? "Failed to create workspace");
    return mapBidWorkspace(result.data);
  }

  async updateBidWorkspace(id: string, changes: Partial<BidWorkspace>): Promise<void> {
    if (IS_DEV) {
      await delay();
      const idx = MOCK_WORKSPACES.findIndex((w) => w.id === id);
      if (idx !== -1) MOCK_WORKSPACES[idx] = { ...MOCK_WORKSPACES[idx], ...changes, modifiedOn: nowIso() };
      return;
    }
    const client = getDataClient();
    await client.updateRecordAsync("cr5ab_bidworkspace", id, changes);
  }

  // ------------------------------------------------------------------
  // BidRoleAssignment
  // ------------------------------------------------------------------

  async createRoleAssignment(
    payload: Omit<BidRoleAssignment, "id" | "createdOn" | "modifiedOn" | "createdBy" | "modifiedBy">
  ): Promise<BidRoleAssignment> {
    if (IS_DEV) {
      await delay();
      const newRa: BidRoleAssignment = { ...payload, id: `tm-${Date.now()}`, createdOn: nowIso(), modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER };
      const ws = MOCK_WORKSPACES.find((w) => w.id === payload.cr5ab_bidworkspaceid.id);
      if (ws) { ws.teamMembers = [...(ws.teamMembers ?? []), newRa]; }
      return newRa;
    }
    const client = getDataClient();
    const body = {
      cr5ab_role: payload.cr5ab_role, cr5ab_isactive: payload.cr5ab_isactive, cr5ab_assigneddate: payload.cr5ab_assigneddate,
      "cr5ab_BidWorkspaceId@odata.bind": `/cr5ab_bidworkspaces(${payload.cr5ab_bidworkspaceid.id})`,
      "cr5ab_UserId@odata.bind": `/systemusers(${payload.cr5ab_userid.id})`,
    };
    const result = await client.createRecordAsync<typeof body, Record<string, unknown>>("cr5ab_bidroleassignment", body);
    if (!result.success) throw new Error(result.error?.message ?? "Failed to assign role");
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const row = result.data as any;
    return { ...payload, id: row.cr5ab_bidroleassignmentid ?? row.id ?? "", createdOn: nowIso(), modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER };
  }

  // ------------------------------------------------------------------
  // BidApproval
  // ------------------------------------------------------------------

  async createApproval(
    payload: Omit<BidApproval, "id" | "createdOn" | "modifiedOn" | "createdBy" | "modifiedBy">
  ): Promise<BidApproval> {
    if (IS_DEV) {
      await delay();
      const newAp: BidApproval = { ...payload, id: `ap-${Date.now()}`, createdOn: nowIso(), modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER };
      const ws = MOCK_WORKSPACES.find((w) => w.id === payload.cr5ab_bidworkspaceid.id);
      if (ws) { ws.approvals = [...(ws.approvals ?? []), newAp]; }
      return newAp;
    }
    const client = getDataClient();
    const body = {
      cr5ab_title: payload.cr5ab_title, cr5ab_approverstage: payload.cr5ab_approverstage,
      cr5ab_status: payload.cr5ab_status, cr5ab_requesteddate: payload.cr5ab_requesteddate,
      "cr5ab_BidWorkspaceId@odata.bind": `/cr5ab_bidworkspaces(${payload.cr5ab_bidworkspaceid.id})`,
      "cr5ab_ApproverId@odata.bind": `/systemusers(${payload.cr5ab_approverid.id})`,
    };
    const result = await client.createRecordAsync<typeof body, Record<string, unknown>>("cr5ab_approval", body);
    if (!result.success) throw new Error(result.error?.message ?? "Failed to create approval");
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const row = result.data as any;
    return { ...payload, id: row.cr5ab_approvalid ?? row.id ?? "", createdOn: nowIso(), modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER };
  }

  async updateApproval(id: string, changes: Partial<BidApproval>): Promise<void> {
    if (IS_DEV) {
      await delay();
      for (const ws of MOCK_WORKSPACES) {
        if (!ws.approvals) continue;
        const idx = ws.approvals.findIndex((a) => a.id === id);
        if (idx !== -1) { ws.approvals[idx] = { ...ws.approvals[idx], ...changes, modifiedOn: nowIso() }; break; }
      }
      return;
    }
    const client = getDataClient();
    await client.updateRecordAsync("cr5ab_approval", id, changes);
  }

  // ------------------------------------------------------------------
  // TOR Items
  // ------------------------------------------------------------------

  async getTorItems(workspaceId: string): Promise<TorItem[]> {
    if (IS_DEV) {
      await delay();
      if (workspaceId === "ws-002") return MOCK_TOR_ITEMS_WS002;
      return MOCK_TOR_ITEMS.filter((t) => t.cr5ab_bidworkspaceid.id === workspaceId);
    }
    const client = getDataClient();
    const result = await client.retrieveMultipleRecordsAsync<Record<string, unknown>>(
      "cr5ab_toritem",
      {
        select: ["cr5ab_toritemid","cr5ab_section","cr5ab_questionnumber","cr5ab_questiondetail","cr5ab_department","cr5ab_supportcontributors","cr5ab_scoreweightingthreshold","cr5ab_specialinstructions","cr5ab_firstdraftdeadline","cr5ab_reviewperiod","cr5ab_finaldraftdeadline","cr5ab_actualdeadline","cr5ab_comments","cr5ab_unittime","cr5ab_answeredstatus","createdon","modifiedon","_cr5ab_allocatedto_value","_cr5ab_bidworkspaceid_value"],
        filter: `_cr5ab_bidworkspaceid_value eq '${workspaceId}'`,
        orderBy: ["cr5ab_section asc", "cr5ab_questionnumber asc"],
      }
    );
    if (!result.success) throw new Error(result.error?.message ?? "Failed to load TOR items");
    return result.data.map(mapTorItem);
  }

  async getTorItemsForUser(userId: string): Promise<TorItem[]> {
    if (IS_DEV) {
      await delay();
      const allItems = [...MOCK_TOR_ITEMS, ...MOCK_TOR_ITEMS_WS002];
      return allItems.filter((t) => t.cr5ab_allocatedto?.id === userId);
    }
    const client = getDataClient();
    const result = await client.retrieveMultipleRecordsAsync<Record<string, unknown>>(
      "cr5ab_toritem",
      {
        select: ["cr5ab_toritemid","cr5ab_section","cr5ab_questionnumber","cr5ab_questiondetail","cr5ab_actualdeadline","cr5ab_answeredstatus","_cr5ab_bidworkspaceid_value"],
        filter: `_cr5ab_allocatedto_value eq '${userId}' and cr5ab_answeredstatus ne 'Y'`,
        orderBy: ["cr5ab_actualdeadline asc"],
      }
    );
    if (!result.success) throw new Error(result.error?.message ?? "Failed to load user TOR items");
    return result.data.map(mapTorItem);
  }

  async createTorItem(
    payload: Omit<TorItem, "id" | "createdOn" | "modifiedOn" | "createdBy" | "modifiedBy">
  ): Promise<TorItem> {
    if (IS_DEV) {
      await delay();
      const newItem: TorItem = { ...payload, id: `tor-${Date.now()}`, createdOn: nowIso(), modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER };
      if (payload.cr5ab_bidworkspaceid.id === "ws-002") {
        MOCK_TOR_ITEMS_WS002.push(newItem);
      } else {
        MOCK_TOR_ITEMS.push(newItem);
      }
      return newItem;
    }
    const client = getDataClient();
    const body = {
      cr5ab_section: payload.cr5ab_section, cr5ab_questionnumber: payload.cr5ab_questionnumber,
      cr5ab_questiondetail: payload.cr5ab_questiondetail, cr5ab_department: payload.cr5ab_department,
      cr5ab_supportcontributors: payload.cr5ab_supportcontributors, cr5ab_scoreweightingthreshold: payload.cr5ab_scoreweightingthreshold,
      cr5ab_specialinstructions: payload.cr5ab_specialinstructions, cr5ab_firstdraftdeadline: payload.cr5ab_firstdraftdeadline,
      cr5ab_reviewperiod: payload.cr5ab_reviewperiod, cr5ab_finaldraftdeadline: payload.cr5ab_finaldraftdeadline,
      cr5ab_actualdeadline: payload.cr5ab_actualdeadline, cr5ab_comments: payload.cr5ab_comments,
      cr5ab_unittime: payload.cr5ab_unittime, cr5ab_answeredstatus: payload.cr5ab_answeredstatus,
      "cr5ab_BidWorkspaceId@odata.bind": `/cr5ab_bidworkspaces(${payload.cr5ab_bidworkspaceid.id})`,
      ...(payload.cr5ab_allocatedto ? { "cr5ab_AllocatedTo@odata.bind": `/systemusers(${payload.cr5ab_allocatedto.id})` } : {}),
    };
    const result = await client.createRecordAsync<typeof body, Record<string, unknown>>("cr5ab_toritem", body);
    if (!result.success) throw new Error(result.error?.message ?? "Failed to create TOR item");
    return mapTorItem(result.data);
  }

  async updateTorItem(id: string, changes: Partial<TorItem>): Promise<TorItem> {
    if (IS_DEV) {
      await delay();
      for (const arr of [MOCK_TOR_ITEMS, MOCK_TOR_ITEMS_WS002]) {
        const idx = arr.findIndex((t) => t.id === id);
        if (idx !== -1) { arr[idx] = { ...arr[idx], ...changes, modifiedOn: nowIso() }; return arr[idx]; }
      }
      throw new Error(`TorItem ${id} not found`);
    }
    const client = getDataClient();
    const result = await client.updateRecordAsync<Partial<TorItem>, Record<string, unknown>>("cr5ab_toritem", id, changes);
    if (!result.success) throw new Error(result.error?.message ?? "Failed to update TOR item");
    return mapTorItem(result.data);
  }

  async deleteTorItem(id: string): Promise<void> {
    if (IS_DEV) {
      await delay();
      const idx1 = MOCK_TOR_ITEMS.findIndex((t) => t.id === id);
      if (idx1 !== -1) { MOCK_TOR_ITEMS.splice(idx1, 1); return; }
      const idx2 = MOCK_TOR_ITEMS_WS002.findIndex((t) => t.id === id);
      if (idx2 !== -1) { MOCK_TOR_ITEMS_WS002.splice(idx2, 1); return; }
      return;
    }
    const client = getDataClient();
    await client.deleteRecordAsync("cr5ab_toritem", id);
  }

  // ------------------------------------------------------------------
  // Clarifications
  // ------------------------------------------------------------------

  async getClarifications(workspaceId: string): Promise<BidClarification[]> {
    if (IS_DEV) {
      await delay();
      return MOCK_CLARIFICATIONS.filter((c) => c.cr5ab_bidworkspaceid.id === workspaceId);
    }
    const client = getDataClient();
    const result = await client.retrieveMultipleRecordsAsync<Record<string, unknown>>(
      "cr5ab_clarification",
      {
        select: ["cr5ab_clarificationid","cr5ab_questionnumber","cr5ab_questiontext","cr5ab_raiseddate","cr5ab_deadline","cr5ab_responsetext","cr5ab_respondeddate","cr5ab_status","cr5ab_iscustomerraised","createdon","modifiedon","_cr5ab_bidworkspaceid_value","_cr5ab_raisedby_value"],
        filter: `_cr5ab_bidworkspaceid_value eq '${workspaceId}'`,
        orderBy: ["cr5ab_raiseddate asc"],
      }
    );
    if (!result.success) throw new Error(result.error?.message ?? "Failed to load clarifications");
    return result.data.map(mapClarification);
  }

  async getClarificationsForUser(userId: string): Promise<BidClarification[]> {
    if (IS_DEV) {
      await delay();
      if (userId === "all") {
        // Return all non-closed clarifications (for calendar population)
        return MOCK_CLARIFICATIONS.filter((c) => c.cr5ab_status !== ClarificationStatus.Closed);
      }
      return MOCK_CLARIFICATIONS.filter(
        (c) => c.cr5ab_status !== ClarificationStatus.Closed &&
               (c.cr5ab_raisedby.id === userId || c.cr5ab_iscustomerraised)
      );
    }
    const client = getDataClient();
    const result = await client.retrieveMultipleRecordsAsync<Record<string, unknown>>(
      "cr5ab_clarification",
      {
        select: ["cr5ab_clarificationid","cr5ab_questionnumber","cr5ab_questiontext","cr5ab_deadline","cr5ab_status","cr5ab_iscustomerraised","createdon","_cr5ab_bidworkspaceid_value","_cr5ab_raisedby_value"],
        filter: `cr5ab_status ne ${ClarificationStatus.Closed} and _cr5ab_raisedby_value eq '${userId}'`,
        orderBy: ["cr5ab_deadline asc"],
      }
    );
    if (!result.success) throw new Error(result.error?.message ?? "Failed to load user clarifications");
    return result.data.map(mapClarification);
  }

  async createClarification(
    payload: Omit<BidClarification, "id" | "createdOn" | "modifiedOn" | "createdBy" | "modifiedBy">
  ): Promise<BidClarification> {
    if (IS_DEV) {
      await delay();
      const newCq: BidClarification = {
        ...payload,
        id: `cq-${Date.now()}`,
        createdOn: nowIso(), modifiedOn: nowIso(), createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
      };
      MOCK_CLARIFICATIONS.push(newCq);
      // Also keep the workspace clarifications array in sync
      const ws = MOCK_WORKSPACES.find((w) => w.id === payload.cr5ab_bidworkspaceid.id);
      if (ws) { ws.clarifications = [...(ws.clarifications ?? []), newCq]; }
      return newCq;
    }
    const client = getDataClient();
    const body = {
      cr5ab_questionnumber: payload.cr5ab_questionnumber,
      cr5ab_questiontext: payload.cr5ab_questiontext,
      cr5ab_raiseddate: payload.cr5ab_raiseddate,
      cr5ab_deadline: payload.cr5ab_deadline,
      cr5ab_status: payload.cr5ab_status,
      cr5ab_iscustomerraised: payload.cr5ab_iscustomerraised,
      "cr5ab_BidWorkspaceId@odata.bind": `/cr5ab_bidworkspaces(${payload.cr5ab_bidworkspaceid.id})`,
      "cr5ab_RaisedBy@odata.bind": `/systemusers(${payload.cr5ab_raisedby.id})`,
    };
    const result = await client.createRecordAsync<typeof body, Record<string, unknown>>("cr5ab_clarification", body);
    if (!result.success) throw new Error(result.error?.message ?? "Failed to create clarification");
    return mapClarification(result.data);
  }

  async updateClarification(id: string, changes: Partial<BidClarification>): Promise<void> {
    if (IS_DEV) {
      await delay();
      const idx = MOCK_CLARIFICATIONS.findIndex((c) => c.id === id);
      if (idx !== -1) {
        MOCK_CLARIFICATIONS[idx] = { ...MOCK_CLARIFICATIONS[idx], ...changes, modifiedOn: nowIso() };
        // Sync to workspace
        for (const ws of MOCK_WORKSPACES) {
          if (!ws.clarifications) continue;
          const wi = ws.clarifications.findIndex((c) => c.id === id);
          if (wi !== -1) { ws.clarifications[wi] = { ...ws.clarifications[wi], ...changes, modifiedOn: nowIso() }; break; }
        }
      }
      return;
    }
    const client = getDataClient();
    await client.updateRecordAsync("cr5ab_clarification", id, changes);
  }

  // ------------------------------------------------------------------
  // Power Automate flows
  // ------------------------------------------------------------------

  async triggerRouteBidFlow(payload: RouteBidFlowPayload): Promise<void> {
    if (IS_DEV) {
      await delay(500);
      console.log("[DEV] triggerRouteBidFlow", payload);
      return;
    }
    // Power Automate HTTP trigger URL — set via environment variable or app config
    const url = (import.meta.env.VITE_FLOW_ROUTE_BID_URL as string | undefined) ?? "";
    if (!url) { console.warn("VITE_FLOW_ROUTE_BID_URL not configured — skipping flow trigger"); return; }
    await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
  }

  async triggerNotifyTeamFlow(payload: NotifyTeamFlowPayload): Promise<void> {
    if (IS_DEV) {
      await delay(300);
      console.log("[DEV] triggerNotifyTeamFlow", payload);
      return;
    }
    const url = (import.meta.env.VITE_FLOW_NOTIFY_TEAM_URL as string | undefined) ?? "";
    if (!url) { console.warn("VITE_FLOW_NOTIFY_TEAM_URL not configured — skipping flow trigger"); return; }
    await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
  }

  /**
   * Upload a file to SharePoint via Power Automate HTTP trigger.
   * Returns the SharePoint URL of the saved file.
   */
  async triggerUploadDocumentFlow(payload: {
    workspaceId: string;
    fileName: string;
    fileBase64: string;
    documentType: string;
    category: string;
  }): Promise<{ sharepointUrl: string }> {
    if (IS_DEV) {
      await delay(800);
      console.log("[DEV] triggerUploadDocumentFlow", payload);
      // Simulate a returned SharePoint URL
      const fakeUrl = `https://ricoh.sharepoint.com/sites/bids/${payload.workspaceId}/${encodeURIComponent(payload.fileName)}`;
      return { sharepointUrl: fakeUrl };
    }
    const url = (import.meta.env.VITE_FLOW_UPLOAD_DOCUMENT_URL as string | undefined) ?? "";
    if (!url) throw new Error("VITE_FLOW_UPLOAD_DOCUMENT_URL not configured");
    const response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (!response.ok) throw new Error(`Upload flow failed: ${response.status}`);
    return response.json() as Promise<{ sharepointUrl: string }>;
  }
}

function delay(ms = 200) {
  return new Promise((r) => setTimeout(r, ms));
}

export const dataverseClient = new DataverseClient();
