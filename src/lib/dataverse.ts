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
  PagedResult,
  DataverseQueryOptions,
} from "../types/dataverse";
import { BidStatus, BidTypeCode } from "../types/dataverse";

// ---------------------------------------------------------------------------
// DataSourcesInfo — maps our table aliases to the Dataverse logical names
// registered in power.config.json under databaseReferences["default.cds"]
// ---------------------------------------------------------------------------

const dataSourcesInfo = {
  cr5ab_bidrequest: {
    tableId: "cr5ab_bidrequests",
    apis: {
      getItems: {
        path: "/cr5ab_bidrequests",
        method: "GET",
        parameters: [],
      },
      getItem: {
        path: "/cr5ab_bidrequests/{id}",
        method: "GET",
        parameters: [{ name: "id", in: "path", required: true, type: "string" }],
      },
      createItem: {
        path: "/cr5ab_bidrequests",
        method: "POST",
        parameters: [],
      },
      updateItem: {
        path: "/cr5ab_bidrequests/{id}",
        method: "PATCH",
        parameters: [{ name: "id", in: "path", required: true, type: "string" }],
      },
    },
  },
  cr5ab_bidworkspace: {
    tableId: "cr5ab_bidworkspaces",
    apis: {
      getItems: {
        path: "/cr5ab_bidworkspaces",
        method: "GET",
        parameters: [],
      },
      getItem: {
        path: "/cr5ab_bidworkspaces/{id}",
        method: "GET",
        parameters: [{ name: "id", in: "path", required: true, type: "string" }],
      },
      createItem: {
        path: "/cr5ab_bidworkspaces",
        method: "POST",
        parameters: [],
      },
      updateItem: {
        path: "/cr5ab_bidworkspaces/{id}",
        method: "PATCH",
        parameters: [{ name: "id", in: "path", required: true, type: "string" }],
      },
    },
  },
  cr5ab_bidtype: {
    tableId: "cr5ab_bidtypes",
    apis: {
      getItems: {
        path: "/cr5ab_bidtypes",
        method: "GET",
        parameters: [],
      },
    },
  },
  cr5ab_approval: {
    tableId: "cr5ab_approvals",
    apis: {
      getItems: {
        path: "/cr5ab_approvals",
        method: "GET",
        parameters: [],
      },
      createItem: {
        path: "/cr5ab_approvals",
        method: "POST",
        parameters: [],
      },
      updateItem: {
        path: "/cr5ab_approvals/{id}",
        method: "PATCH",
        parameters: [{ name: "id", in: "path", required: true, type: "string" }],
      },
    },
  },
  cr5ab_bidroleassignment: {
    tableId: "cr5ab_bidroleassignments",
    apis: {
      getItems: {
        path: "/cr5ab_bidroleassignments",
        method: "GET",
        parameters: [],
      },
      createItem: {
        path: "/cr5ab_bidroleassignments",
        method: "POST",
        parameters: [],
      },
    },
  },
  cr5ab_bidstatus: {
    tableId: "cr5ab_bidstatuses",
    apis: {
      getItems: {
        path: "/cr5ab_bidstatuses",
        method: "GET",
        parameters: [],
      },
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
// Dev-mode mock data (used when IS_DEV and not running via Local Play)
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

const MOCK_BID_TYPES: BidType[] = [
  {
    id: "bt-001",
    cr5ab_code: BidTypeCode.SupplierQuestionnaire,
    cr5ab_name: "Supplier Questionnaire",
    cr5ab_description: "Pre-qualification questionnaire submitted by potential suppliers.",
    cr5ab_routingteam: "Procurement",
    cr5ab_routingemail: "procurement@ricoh.co.uk",
    cr5ab_slaresponsedays: 5,
    createdOn: nowIso(), modifiedOn: nowIso(),
    createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "bt-002",
    cr5ab_code: BidTypeCode.SalesLed,
    cr5ab_name: "Sales Led",
    cr5ab_description: "Bids driven by the Sales team in response to an RFP.",
    cr5ab_routingteam: "Sales Bids",
    cr5ab_routingemail: "salesbids@ricoh.co.uk",
    cr5ab_slaresponsedays: 3,
    createdOn: nowIso(), modifiedOn: nowIso(),
    createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "bt-003",
    cr5ab_code: BidTypeCode.SMEToQualify,
    cr5ab_name: "SME to Qualify",
    cr5ab_description: "Opportunities routed to an SME for qualification.",
    cr5ab_routingteam: "SME Pool",
    cr5ab_routingemail: "sme@ricoh.co.uk",
    cr5ab_slaresponsedays: 7,
    createdOn: nowIso(), modifiedOn: nowIso(),
    createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "bt-004",
    cr5ab_code: BidTypeCode.BidManagement,
    cr5ab_name: "Bid Management",
    cr5ab_description: "Full bid management lifecycle handled by the Bid Management team.",
    cr5ab_routingteam: "Bid Management",
    cr5ab_routingemail: "bidmanagement@ricoh.co.uk",
    cr5ab_slaresponsedays: 2,
    createdOn: nowIso(), modifiedOn: nowIso(),
    createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
];

const MOCK_BID_REQUESTS: BidRequest[] = [
  {
    id: "br-001",
    cr5ab_bidreferencenumber: "BID-2026-00001",
    cr5ab_title: "NHS Digital Printing Framework",
    cr5ab_bidtypeid: { id: "bt-004", cr5ab_name: "Bid Management", cr5ab_code: BidTypeCode.BidManagement },
    cr5ab_status: BidStatus.InProgress,
    cr5ab_customername: "NHS England",
    cr5ab_customerindustry: "Healthcare",
    cr5ab_estimatedvalue: 4500000,
    cr5ab_currency: "GBP",
    cr5ab_submissiondeadline: "2026-04-15T17:00:00.000Z",
    cr5ab_expectedawarddate: "2026-05-30T00:00:00.000Z",
    cr5ab_description: "National framework for managed print services across NHS trusts.",
    cr5ab_submittedby: SYSTEM_USER,
    cr5ab_assignedto: SYSTEM_USER,
    cr5ab_routedto: "Bid Management",
    createdOn: "2026-02-10T09:00:00.000Z", modifiedOn: nowIso(),
    createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "br-002",
    cr5ab_bidreferencenumber: "BID-2026-00002",
    cr5ab_title: "Central Government MFD Refresh",
    cr5ab_bidtypeid: { id: "bt-002", cr5ab_name: "Sales Led", cr5ab_code: BidTypeCode.SalesLed },
    cr5ab_status: BidStatus.Submitted,
    cr5ab_customername: "HMRC",
    cr5ab_customerindustry: "Government",
    cr5ab_estimatedvalue: 1200000,
    cr5ab_currency: "GBP",
    cr5ab_submissiondeadline: "2026-03-30T17:00:00.000Z",
    cr5ab_description: "Replacement of end-of-life MFD estate across HMRC offices.",
    cr5ab_submittedby: SYSTEM_USER,
    createdOn: "2026-02-20T11:30:00.000Z", modifiedOn: nowIso(),
    createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "br-003",
    cr5ab_bidreferencenumber: "BID-2026-00003",
    cr5ab_title: "University Print Supplier PQQ",
    cr5ab_bidtypeid: { id: "bt-001", cr5ab_name: "Supplier Questionnaire", cr5ab_code: BidTypeCode.SupplierQuestionnaire },
    cr5ab_status: BidStatus.InReview,
    cr5ab_customername: "University of Manchester",
    cr5ab_customerindustry: "Education",
    cr5ab_estimatedvalue: 350000,
    cr5ab_currency: "GBP",
    cr5ab_submissiondeadline: "2026-03-20T17:00:00.000Z",
    cr5ab_description: "Pre-qualification questionnaire for university print supplier list.",
    cr5ab_submittedby: SYSTEM_USER,
    createdOn: "2026-03-01T08:00:00.000Z", modifiedOn: nowIso(),
    createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "br-004",
    cr5ab_bidreferencenumber: "BID-2026-00004",
    cr5ab_title: "Retail Chain Document Management",
    cr5ab_bidtypeid: { id: "bt-003", cr5ab_name: "SME to Qualify", cr5ab_code: BidTypeCode.SMEToQualify },
    cr5ab_status: BidStatus.Draft,
    cr5ab_customername: "Tesco PLC",
    cr5ab_customerindustry: "Retail",
    cr5ab_estimatedvalue: 780000,
    cr5ab_currency: "GBP",
    cr5ab_submissiondeadline: "2026-04-30T17:00:00.000Z",
    cr5ab_description: "Document management and workflow automation for retail operations.",
    cr5ab_submittedby: SYSTEM_USER,
    createdOn: "2026-03-10T14:00:00.000Z", modifiedOn: nowIso(),
    createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
  {
    id: "br-005",
    cr5ab_bidreferencenumber: "BID-2025-00041",
    cr5ab_title: "Local Authority Print Contract",
    cr5ab_bidtypeid: { id: "bt-004", cr5ab_name: "Bid Management", cr5ab_code: BidTypeCode.BidManagement },
    cr5ab_status: BidStatus.Won,
    cr5ab_customername: "Birmingham City Council",
    cr5ab_customerindustry: "Government",
    cr5ab_estimatedvalue: 2100000,
    cr5ab_currency: "GBP",
    cr5ab_submissiondeadline: "2025-11-30T17:00:00.000Z",
    cr5ab_description: "5-year managed print services contract for Birmingham City Council.",
    cr5ab_submittedby: SYSTEM_USER,
    cr5ab_assignedto: SYSTEM_USER,
    createdOn: "2025-09-15T09:00:00.000Z", modifiedOn: "2025-12-10T09:00:00.000Z",
    createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
];

const MOCK_WORKSPACES: BidWorkspace[] = [
  {
    id: "ws-001",
    cr5ab_title: "NHS Digital Printing Framework — Workspace",
    cr5ab_bidrequestid: {
      id: "br-001",
      cr5ab_title: "NHS Digital Printing Framework",
      cr5ab_bidreferencenumber: "BID-2026-00001",
    },
    cr5ab_status: BidStatus.InProgress,
    cr5ab_bidmanagerid: SYSTEM_USER,
    cr5ab_completionpercentage: 45,
    cr5ab_sharepointfolderurl: "https://ricoh.sharepoint.com/sites/bids/BID-2026-00001",
    createdOn: "2026-02-12T09:00:00.000Z", modifiedOn: nowIso(),
    createdBy: SYSTEM_USER, modifiedBy: SYSTEM_USER,
  },
];

// ---------------------------------------------------------------------------
// Helper: map a raw Dataverse row to our BidRequest type
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

// ---------------------------------------------------------------------------
// Client
// ---------------------------------------------------------------------------

class DataverseClient {

  // ------------------------------------------------------------------
  // BidType
  // ------------------------------------------------------------------

  async getBidTypes(): Promise<BidType[]> {
    if (IS_DEV) {
      await delay();
      return MOCK_BID_TYPES;
    }

    const client = getDataClient();
    const result = await client.retrieveMultipleRecordsAsync<Record<string, unknown>>(
      "cr5ab_bidtype",
      {
        select: [
          "cr5ab_bidtypeid",
          "cr5ab_name",
          "cr5ab_code",
          "cr5ab_description",
          "cr5ab_routingteam",
          "cr5ab_routingemail",
          "cr5ab_slaresponsedays",
        ],
        orderBy: ["cr5ab_code asc"],
      }
    );

    if (!result.success) throw new Error(result.error?.message ?? "Failed to load bid types");
    return result.data.map(mapBidType);
  }

  // ------------------------------------------------------------------
  // BidRequest
  // ------------------------------------------------------------------

  async getBidRequests(options?: DataverseQueryOptions): Promise<PagedResult<BidRequest>> {
    if (IS_DEV) {
      await delay();
      return { value: MOCK_BID_REQUESTS, totalCount: MOCK_BID_REQUESTS.length };
    }

    const client = getDataClient();
    const result = await client.retrieveMultipleRecordsAsync<Record<string, unknown>>(
      "cr5ab_bidrequest",
      {
        select: [
          "cr5ab_bidrequestid",
          "cr5ab_bidreferencenumber",
          "cr5ab_title",
          "cr5ab_status",
          "cr5ab_customername",
          "cr5ab_customerindustry",
          "cr5ab_estimatedvalue",
          "cr5ab_currency",
          "cr5ab_submissiondeadline",
          "cr5ab_expectedawarddate",
          "cr5ab_description",
          "cr5ab_routedto",
          "createdon",
          "modifiedon",
          "_cr5ab_bidtypeid_value",
          "_cr5ab_submittedby_value",
          "_cr5ab_assignedto_value",
        ],
        filter: options?.filter,
        orderBy: options?.orderBy ? [options.orderBy] : ["createdon desc"],
        top: options?.top,
        skip: options?.skip,
        count: true,
      }
    );

    if (!result.success) throw new Error(result.error?.message ?? "Failed to load bid requests");
    return {
      value: result.data.map(mapBidRequest),
      totalCount: result.count,
    };
  }

  async getBidRequest(id: string): Promise<BidRequest | null> {
    if (IS_DEV) {
      await delay();
      return MOCK_BID_REQUESTS.find((b) => b.id === id) ?? null;
    }

    const client = getDataClient();
    const result = await client.retrieveRecordAsync<Record<string, unknown>>(
      "cr5ab_bidrequest",
      id
    );

    if (!result.success) return null;
    return mapBidRequest(result.data);
  }

  async createBidRequest(
    payload: Omit<BidRequest, "id" | "createdOn" | "modifiedOn" | "createdBy" | "modifiedBy">
  ): Promise<BidRequest> {
    if (IS_DEV) {
      await delay();
      const newRecord: BidRequest = {
        ...payload,
        id: `br-${Date.now()}`,
        createdOn: nowIso(),
        modifiedOn: nowIso(),
        createdBy: SYSTEM_USER,
        modifiedBy: SYSTEM_USER,
      };
      MOCK_BID_REQUESTS.push(newRecord);
      return newRecord;
    }

    const client = getDataClient();
    const body = {
      cr5ab_title: payload.cr5ab_title,
      cr5ab_bidreferencenumber: payload.cr5ab_bidreferencenumber,
      cr5ab_status: payload.cr5ab_status,
      cr5ab_customername: payload.cr5ab_customername,
      cr5ab_customerindustry: payload.cr5ab_customerindustry,
      cr5ab_estimatedvalue: payload.cr5ab_estimatedvalue,
      cr5ab_currency: payload.cr5ab_currency,
      cr5ab_submissiondeadline: payload.cr5ab_submissiondeadline,
      cr5ab_expectedawarddate: payload.cr5ab_expectedawarddate,
      cr5ab_contractduration: payload.cr5ab_contractduration,
      cr5ab_description: payload.cr5ab_description,
      cr5ab_scope: payload.cr5ab_scope,
      cr5ab_specialrequirements: payload.cr5ab_specialrequirements,
      cr5ab_incumbentvendor: payload.cr5ab_incumbentvendor,
      cr5ab_routedto: payload.cr5ab_routedto,
      // Lookup fields use OData bind syntax
      "cr5ab_BidTypeId@odata.bind": payload.cr5ab_bidtypeid?.id
        ? `/cr5ab_bidtypes(${payload.cr5ab_bidtypeid.id})`
        : undefined,
    };

    const result = await client.createRecordAsync<typeof body, Record<string, unknown>>(
      "cr5ab_bidrequest",
      body
    );

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
    const result = await client.updateRecordAsync<Partial<BidRequest>, Record<string, unknown>>(
      "cr5ab_bidrequest",
      id,
      changes
    );

    if (!result.success) throw new Error(result.error?.message ?? "Failed to update bid request");
    return mapBidRequest(result.data);
  }

  // ------------------------------------------------------------------
  // BidWorkspace
  // ------------------------------------------------------------------

  async getBidWorkspaces(options?: DataverseQueryOptions): Promise<PagedResult<BidWorkspace>> {
    if (IS_DEV) {
      await delay();
      return { value: MOCK_WORKSPACES, totalCount: MOCK_WORKSPACES.length };
    }

    const client = getDataClient();
    const result = await client.retrieveMultipleRecordsAsync<Record<string, unknown>>(
      "cr5ab_bidworkspace",
      {
        select: [
          "cr5ab_bidworkspaceid",
          "cr5ab_title",
          "cr5ab_status",
          "cr5ab_completionpercentage",
          "cr5ab_sharepointfolderurl",
          "cr5ab_teamschannelurl",
          "createdon",
          "modifiedon",
          "_cr5ab_bidrequestid_value",
          "_cr5ab_bidmanagerid_value",
        ],
        filter: options?.filter,
        orderBy: options?.orderBy ? [options.orderBy] : ["createdon desc"],
        top: options?.top,
        count: true,
      }
    );

    if (!result.success) throw new Error(result.error?.message ?? "Failed to load workspaces");
    return {
      value: result.data.map(mapBidWorkspace),
      totalCount: result.count,
    };
  }

  async getBidWorkspace(id: string): Promise<BidWorkspace | null> {
    if (IS_DEV) {
      await delay();
      return MOCK_WORKSPACES.find((w) => w.id === id) ?? null;
    }

    const client = getDataClient();
    const result = await client.retrieveRecordAsync<Record<string, unknown>>(
      "cr5ab_bidworkspace",
      id
    );

    if (!result.success) return null;
    return mapBidWorkspace(result.data);
  }
}

function delay() {
  return new Promise((r) => setTimeout(r, 200));
}

export const dataverseClient = new DataverseClient();
