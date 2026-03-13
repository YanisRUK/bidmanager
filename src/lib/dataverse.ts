/**
 * Dataverse client abstraction
 *
 * In production this wraps the Power Apps SDK connector.
 * In local dev it returns typed mock/stub data so every page renders.
 *
 * Replace the stubbed implementations with real SDK calls as each feature
 * is built out.
 */

import type {
  BidRequest,
  BidWorkspace,
  BidType,
  PagedResult,
  DataverseQueryOptions,
} from "../types/dataverse";
import { BidStatus, BidTypeCode } from "../types/dataverse";

// ---------------------------------------------------------------------------
// Helpers
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

// ---------------------------------------------------------------------------
// Mock data
// ---------------------------------------------------------------------------

const MOCK_BID_TYPES: BidType[] = [
  {
    id: "bt-001",
    ricoh_code: BidTypeCode.SupplierQuestionnaire,
    ricoh_name: "Supplier Questionnaire",
    ricoh_description:
      "Pre-qualification questionnaire submitted by potential suppliers.",
    ricoh_routingTeam: "Procurement",
    ricoh_routingEmail: "procurement@ricoh.co.uk",
    ricoh_slaResponseDays: 5,
    createdOn: nowIso(),
    modifiedOn: nowIso(),
    createdBy: SYSTEM_USER,
    modifiedBy: SYSTEM_USER,
  },
  {
    id: "bt-002",
    ricoh_code: BidTypeCode.SalesLed,
    ricoh_name: "Sales Led",
    ricoh_description: "Bids driven by the Sales team in response to an RFP.",
    ricoh_routingTeam: "Sales Bids",
    ricoh_routingEmail: "salesbids@ricoh.co.uk",
    ricoh_slaResponseDays: 3,
    createdOn: nowIso(),
    modifiedOn: nowIso(),
    createdBy: SYSTEM_USER,
    modifiedBy: SYSTEM_USER,
  },
  {
    id: "bt-003",
    ricoh_code: BidTypeCode.SMEToQualify,
    ricoh_name: "SME to Qualify",
    ricoh_description: "Opportunities routed to an SME for qualification.",
    ricoh_routingTeam: "SME Pool",
    ricoh_routingEmail: "sme@ricoh.co.uk",
    ricoh_slaResponseDays: 7,
    createdOn: nowIso(),
    modifiedOn: nowIso(),
    createdBy: SYSTEM_USER,
    modifiedBy: SYSTEM_USER,
  },
  {
    id: "bt-004",
    ricoh_code: BidTypeCode.BidManagement,
    ricoh_name: "Bid Management",
    ricoh_description:
      "Full bid management lifecycle handled by the Bid Management team.",
    ricoh_routingTeam: "Bid Management",
    ricoh_routingEmail: "bidmanagement@ricoh.co.uk",
    ricoh_slaResponseDays: 2,
    createdOn: nowIso(),
    modifiedOn: nowIso(),
    createdBy: SYSTEM_USER,
    modifiedBy: SYSTEM_USER,
  },
];

const MOCK_BID_REQUESTS: BidRequest[] = [
  {
    id: "br-001",
    ricoh_bidreferencenumber: "BID-2026-00001",
    ricoh_title: "NHS Digital Printing Framework",
    ricoh_bidtypeid: { id: "bt-004", ricoh_name: "Bid Management", ricoh_code: BidTypeCode.BidManagement },
    ricoh_status: BidStatus.InProgress,
    ricoh_customername: "NHS England",
    ricoh_customerindustry: "Healthcare",
    ricoh_estimatedvalue: 4500000,
    ricoh_currency: "GBP",
    ricoh_submissiondeadline: "2026-04-15T17:00:00.000Z",
    ricoh_expectedawarddate: "2026-05-30T00:00:00.000Z",
    ricoh_description: "National framework for managed print services across NHS trusts.",
    ricoh_submittedby: SYSTEM_USER,
    ricoh_assignedto: SYSTEM_USER,
    ricoh_routedto: "Bid Management",
    createdOn: "2026-02-10T09:00:00.000Z",
    modifiedOn: nowIso(),
    createdBy: SYSTEM_USER,
    modifiedBy: SYSTEM_USER,
  },
  {
    id: "br-002",
    ricoh_bidreferencenumber: "BID-2026-00002",
    ricoh_title: "Central Government MFD Refresh",
    ricoh_bidtypeid: { id: "bt-002", ricoh_name: "Sales Led", ricoh_code: BidTypeCode.SalesLed },
    ricoh_status: BidStatus.Submitted,
    ricoh_customername: "HMRC",
    ricoh_customerindustry: "Government",
    ricoh_estimatedvalue: 1200000,
    ricoh_currency: "GBP",
    ricoh_submissiondeadline: "2026-03-30T17:00:00.000Z",
    ricoh_description: "Replacement of end-of-life MFD estate across HMRC offices.",
    ricoh_submittedby: SYSTEM_USER,
    createdOn: "2026-02-20T11:30:00.000Z",
    modifiedOn: nowIso(),
    createdBy: SYSTEM_USER,
    modifiedBy: SYSTEM_USER,
  },
  {
    id: "br-003",
    ricoh_bidreferencenumber: "BID-2026-00003",
    ricoh_title: "University Print Supplier PQQ",
    ricoh_bidtypeid: { id: "bt-001", ricoh_name: "Supplier Questionnaire", ricoh_code: BidTypeCode.SupplierQuestionnaire },
    ricoh_status: BidStatus.InReview,
    ricoh_customername: "University of Manchester",
    ricoh_customerindustry: "Education",
    ricoh_estimatedvalue: 350000,
    ricoh_currency: "GBP",
    ricoh_submissiondeadline: "2026-03-20T17:00:00.000Z",
    ricoh_description: "Pre-qualification questionnaire for university print supplier list.",
    ricoh_submittedby: SYSTEM_USER,
    createdOn: "2026-03-01T08:00:00.000Z",
    modifiedOn: nowIso(),
    createdBy: SYSTEM_USER,
    modifiedBy: SYSTEM_USER,
  },
  {
    id: "br-004",
    ricoh_bidreferencenumber: "BID-2026-00004",
    ricoh_title: "Retail Chain Document Management",
    ricoh_bidtypeid: { id: "bt-003", ricoh_name: "SME to Qualify", ricoh_code: BidTypeCode.SMEToQualify },
    ricoh_status: BidStatus.Draft,
    ricoh_customername: "Tesco PLC",
    ricoh_customerindustry: "Retail",
    ricoh_estimatedvalue: 780000,
    ricoh_currency: "GBP",
    ricoh_submissiondeadline: "2026-04-30T17:00:00.000Z",
    ricoh_description: "Document management and workflow automation for retail operations.",
    ricoh_submittedby: SYSTEM_USER,
    createdOn: "2026-03-10T14:00:00.000Z",
    modifiedOn: nowIso(),
    createdBy: SYSTEM_USER,
    modifiedBy: SYSTEM_USER,
  },
  {
    id: "br-005",
    ricoh_bidreferencenumber: "BID-2025-00041",
    ricoh_title: "Local Authority Print Contract",
    ricoh_bidtypeid: { id: "bt-004", ricoh_name: "Bid Management", ricoh_code: BidTypeCode.BidManagement },
    ricoh_status: BidStatus.Won,
    ricoh_customername: "Birmingham City Council",
    ricoh_customerindustry: "Government",
    ricoh_estimatedvalue: 2100000,
    ricoh_currency: "GBP",
    ricoh_submissiondeadline: "2025-11-30T17:00:00.000Z",
    ricoh_description: "5-year managed print services contract for Birmingham City Council.",
    ricoh_submittedby: SYSTEM_USER,
    ricoh_assignedto: SYSTEM_USER,
    createdOn: "2025-09-15T09:00:00.000Z",
    modifiedOn: "2025-12-10T09:00:00.000Z",
    createdBy: SYSTEM_USER,
    modifiedBy: SYSTEM_USER,
  },
];

const MOCK_WORKSPACES: BidWorkspace[] = [
  {
    id: "ws-001",
    ricoh_title: "NHS Digital Printing Framework — Workspace",
    ricoh_bidrequestid: {
      id: "br-001",
      ricoh_title: "NHS Digital Printing Framework",
      ricoh_bidreferencenumber: "BID-2026-00001",
    },
    ricoh_status: BidStatus.InProgress,
    ricoh_bidmanagerid: SYSTEM_USER,
    ricoh_completionpercentage: 45,
    ricoh_sharepointfolderurl:
      "https://ricoh.sharepoint.com/sites/bids/BID-2026-00001",
    createdOn: "2026-02-12T09:00:00.000Z",
    modifiedOn: nowIso(),
    createdBy: SYSTEM_USER,
    modifiedBy: SYSTEM_USER,
  },
];

// ---------------------------------------------------------------------------
// Client class
// ---------------------------------------------------------------------------

class DataverseClient {
  private async callSdk<T>(
    _operation: string,
    devFallback: () => T
  ): Promise<T> {
    if (IS_DEV) {
      // Simulate async latency in dev
      await new Promise((r) => setTimeout(r, 200));
      return devFallback();
    }

    // TODO: replace with real SDK connector calls, e.g.:
    // const sdk = await import("@microsoft/powerapps-code-apps");
    // return sdk.dataverse.query(...);
    throw new Error("SDK not implemented for this operation: " + _operation);
  }

  // ------------------------------------------------------------------
  // BidType
  // ------------------------------------------------------------------

  async getBidTypes(): Promise<BidType[]> {
    return this.callSdk("getBidTypes", () => MOCK_BID_TYPES);
  }

  // ------------------------------------------------------------------
  // BidRequest
  // ------------------------------------------------------------------

  async getBidRequests(
    options?: DataverseQueryOptions
  ): Promise<PagedResult<BidRequest>> {
    return this.callSdk("getBidRequests", () => {
      let results = [...MOCK_BID_REQUESTS];

      // Very basic client-side filtering for dev mock
      if (options?.filter) {
        // Not implemented in mock — return all
      }
      if (options?.top) {
        results = results.slice(options.skip ?? 0, (options.skip ?? 0) + options.top);
      }

      return { value: results, totalCount: MOCK_BID_REQUESTS.length };
    });
  }

  async getBidRequest(id: string): Promise<BidRequest | null> {
    return this.callSdk("getBidRequest", () => {
      return MOCK_BID_REQUESTS.find((b) => b.id === id) ?? null;
    });
  }

  async createBidRequest(
    data: Omit<BidRequest, "id" | "createdOn" | "modifiedOn" | "createdBy" | "modifiedBy">
  ): Promise<BidRequest> {
    return this.callSdk("createBidRequest", () => {
      const newRecord: BidRequest = {
        ...data,
        id: `br-${Date.now()}`,
        createdOn: nowIso(),
        modifiedOn: nowIso(),
        createdBy: SYSTEM_USER,
        modifiedBy: SYSTEM_USER,
      };
      MOCK_BID_REQUESTS.push(newRecord);
      return newRecord;
    });
  }

  async updateBidRequest(
    id: string,
    data: Partial<BidRequest>
  ): Promise<BidRequest> {
    return this.callSdk("updateBidRequest", () => {
      const idx = MOCK_BID_REQUESTS.findIndex((b) => b.id === id);
      if (idx === -1) throw new Error(`BidRequest ${id} not found`);
      MOCK_BID_REQUESTS[idx] = {
        ...MOCK_BID_REQUESTS[idx],
        ...data,
        modifiedOn: nowIso(),
        modifiedBy: SYSTEM_USER,
      };
      return MOCK_BID_REQUESTS[idx];
    });
  }

  // ------------------------------------------------------------------
  // BidWorkspace
  // ------------------------------------------------------------------

  async getBidWorkspaces(
    options?: DataverseQueryOptions
  ): Promise<PagedResult<BidWorkspace>> {
    return this.callSdk("getBidWorkspaces", () => {
      const results = [...MOCK_WORKSPACES];
      void options;
      return { value: results, totalCount: MOCK_WORKSPACES.length };
    });
  }

  async getBidWorkspace(id: string): Promise<BidWorkspace | null> {
    return this.callSdk("getBidWorkspace", () => {
      return MOCK_WORKSPACES.find((w) => w.id === id) ?? null;
    });
  }
}

// Singleton export
export const dataverseClient = new DataverseClient();
