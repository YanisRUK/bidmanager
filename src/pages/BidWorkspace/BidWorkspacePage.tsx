/**
 * Bid Workspace Page
 *
 * Individual workspace for an active bid showing:
 * - Bid summary header
 * - Progress tracker
 * - Team members & roles
 * - Approvals chain
 * - Document library (SharePoint links)
 * - Activity feed placeholder
 */

import React from "react";
import { useNavigate, useSearchParams } from "react-router-dom";
import {
  makeStyles,
  shorthands,
  tokens,
  Text,
  Button,
  Card,
  CardHeader,
  ProgressBar,
  Badge,
  Table,
  TableHeader,
  TableHeaderCell,
  TableBody,
  TableRow,
  TableCell,
  TableCellLayout,
  Tab,
  TabList,
} from "@fluentui/react-components";
import {
  ArrowLeftRegular,
  PeopleRegular,
  DocumentRegular,
  CheckmarkCircleRegular,
  ClockRegular,
  LinkRegular,
  AddRegular,
} from "@fluentui/react-icons";
import { PageHeader } from "../../components/common/PageHeader";
import { StatusBadge } from "../../components/common/StatusBadge";
import { LoadingState } from "../../components/common/LoadingState";
import { ErrorState } from "../../components/common/ErrorState";
import { EmptyState } from "../../components/common/EmptyState";
import { useDataverse } from "../../hooks/useDataverse";
import { dataverseClient } from "../../lib/dataverse";
import { ApprovalStatus, TeamRoleLabel } from "../../types/dataverse";
import type { BidWorkspace, BidTeamMember, BidApproval } from "../../types/dataverse";

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

const useStyles = makeStyles({
  headerGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr 1fr",
    gap: tokens.spacingVerticalM,
    marginBottom: tokens.spacingVerticalXL,
    "@media (max-width: 800px)": {
      gridTemplateColumns: "1fr",
    },
  },
  metaItem: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  tabContent: {
    paddingTop: tokens.spacingVerticalL,
  },
  progressCard: {
    marginBottom: tokens.spacingVerticalL,
  },
  progressRow: {
    display: "flex",
    justifyContent: "space-between",
    marginBottom: tokens.spacingVerticalS,
  },
  teamTable: {
    width: "100%",
  },
  approvalRow: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    ...shorthands.padding(tokens.spacingVerticalM, "0"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ":last-child": { borderBottom: "none" },
  },
  stageNumber: {
    width: "32px",
    height: "32px",
    borderRadius: "50%",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    backgroundColor: tokens.colorNeutralBackground3,
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    flexShrink: 0,
  },
  docRow: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    ...shorthands.padding(tokens.spacingVerticalS, "0"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ":last-child": { borderBottom: "none" },
  },
});

// ---------------------------------------------------------------------------
// Approval badge helper
// ---------------------------------------------------------------------------

function ApprovalBadge({ status }: { status: ApprovalStatus }) {
  const map: Record<ApprovalStatus, { color: "success" | "warning" | "danger" | "subtle"; label: string }> = {
    [ApprovalStatus.Pending]:  { color: "warning", label: "Pending" },
    [ApprovalStatus.Approved]: { color: "success", label: "Approved" },
    [ApprovalStatus.Rejected]: { color: "danger",  label: "Rejected" },
  };
  const { color, label } = map[status];
  return <Badge appearance="filled" color={color}>{label}</Badge>;
}

// ---------------------------------------------------------------------------
// Mock team / approvals when no workspace is loaded
// ---------------------------------------------------------------------------

const MOCK_TEAM: BidTeamMember[] = [
  { id: "tm-1", ricoh_bidworkspaceid: { id: "ws-001", ricoh_title: "" }, ricoh_userid: { id: "u1", fullName: "Sarah Mitchell", email: "s.mitchell@ricoh.co.uk" }, ricoh_role: 100000000, ricoh_isactive: true, ricoh_assigneddate: "2026-02-12T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: { id: "", fullName: "", email: "" }, modifiedBy: { id: "", fullName: "", email: "" } },
  { id: "tm-2", ricoh_bidworkspaceid: { id: "ws-001", ricoh_title: "" }, ricoh_userid: { id: "u2", fullName: "James O'Brien", email: "j.obrien@ricoh.co.uk" }, ricoh_role: 100000002, ricoh_isactive: true, ricoh_assigneddate: "2026-02-13T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: { id: "", fullName: "", email: "" }, modifiedBy: { id: "", fullName: "", email: "" } },
  { id: "tm-3", ricoh_bidworkspaceid: { id: "ws-001", ricoh_title: "" }, ricoh_userid: { id: "u3", fullName: "Priya Sharma", email: "p.sharma@ricoh.co.uk" }, ricoh_role: 100000003, ricoh_isactive: true, ricoh_assigneddate: "2026-02-14T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: { id: "", fullName: "", email: "" }, modifiedBy: { id: "", fullName: "", email: "" } },
];

const MOCK_APPROVALS: BidApproval[] = [
  { id: "ap-1", ricoh_bidworkspaceid: { id: "ws-001", ricoh_title: "" }, ricoh_title: "Technical Review", ricoh_approverstage: 1, ricoh_approverid: { id: "u3", fullName: "Priya Sharma", email: "p.sharma@ricoh.co.uk" }, ricoh_status: ApprovalStatus.Approved, ricoh_requesteddate: "2026-02-20T09:00:00Z", ricoh_respondeddate: "2026-02-22T14:00:00Z", ricoh_comments: "Technically sound. Proceed.", createdOn: "", modifiedOn: "", createdBy: { id: "", fullName: "", email: "" }, modifiedBy: { id: "", fullName: "", email: "" } },
  { id: "ap-2", ricoh_bidworkspaceid: { id: "ws-001", ricoh_title: "" }, ricoh_title: "Commercial Review", ricoh_approverstage: 2, ricoh_approverid: { id: "u4", fullName: "Tom Watkins", email: "t.watkins@ricoh.co.uk" }, ricoh_status: ApprovalStatus.Pending, ricoh_requesteddate: "2026-02-23T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: { id: "", fullName: "", email: "" }, modifiedBy: { id: "", fullName: "", email: "" } },
  { id: "ap-3", ricoh_bidworkspaceid: { id: "ws-001", ricoh_title: "" }, ricoh_title: "Executive Sign-off", ricoh_approverstage: 3, ricoh_approverid: { id: "u5", fullName: "Helen Cross", email: "h.cross@ricoh.co.uk" }, ricoh_status: ApprovalStatus.Pending, ricoh_requesteddate: "2026-02-23T09:00:00Z", createdOn: "", modifiedOn: "", createdBy: { id: "", fullName: "", email: "" }, modifiedBy: { id: "", fullName: "", email: "" } },
];

// ---------------------------------------------------------------------------
// Page
// ---------------------------------------------------------------------------

type WorkspaceTab = "overview" | "team" | "approvals" | "documents";

export function BidWorkspacePage() {
  const styles = useStyles();
  const navigate = useNavigate();
  const [searchParams] = useSearchParams();
  const [activeTab, setActiveTab] = React.useState<WorkspaceTab>("overview");

  // If a bidId query param is present, try to find the workspace for that bid
  const bidId = searchParams.get("bidId");
  const workspaceId = searchParams.get("workspaceId") ?? "ws-001";

  const { data: workspace, isLoading, error, refresh } = useDataverse(
    () => dataverseClient.getBidWorkspace(workspaceId),
    [workspaceId]
  );

  const { data: bidRequest } = useDataverse(
    () => (bidId ? dataverseClient.getBidRequest(bidId) : Promise.resolve(null)),
    [bidId]
  );

  if (isLoading) return <LoadingState label="Loading workspace..." />;
  if (error) return <ErrorState message={error} onRetry={refresh} />;

  // Use workspace data or fall back to bid request details
  const title = workspace?.ricoh_title ?? bidRequest?.ricoh_title ?? "Bid Workspace";
  const status = workspace?.ricoh_status ?? bidRequest?.ricoh_status;
  const progress = workspace?.ricoh_completionpercentage ?? 0;
  const teamMembers = workspace?.teamMembers ?? MOCK_TEAM;
  const approvals = workspace?.approvals ?? MOCK_APPROVALS;
  const documents = workspace?.documents ?? [];

  return (
    <div>
      <PageHeader
        title={title}
        subtitle={
          workspace?.ricoh_bidrequestid?.ricoh_bidreferencenumber ??
          bidRequest?.ricoh_bidreferencenumber ??
          "Workspace"
        }
        actions={
          <>
            <Button
              appearance="subtle"
              icon={<ArrowLeftRegular />}
              onClick={() => navigate("/bid-register")}
            >
              Back to Register
            </Button>
            {workspace?.ricoh_sharepointfolderurl && (
              <Button
                appearance="secondary"
                icon={<LinkRegular />}
                as="a"
                href={workspace.ricoh_sharepointfolderurl}
                target="_blank"
                rel="noopener noreferrer"
              >
                SharePoint
              </Button>
            )}
          </>
        }
      />

      {/* Status + meta strip */}
      <div className={styles.headerGrid}>
        <div className={styles.metaItem}>
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Status</Text>
          {status !== undefined ? <StatusBadge status={status} /> : <Text>—</Text>}
        </div>
        <div className={styles.metaItem}>
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Bid Manager</Text>
          <Text size={300} weight="semibold">
            {workspace?.ricoh_bidmanagerid?.fullName ?? "—"}
          </Text>
        </div>
        <div className={styles.metaItem}>
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Completion</Text>
          <Text size={300} weight="semibold">{progress}%</Text>
        </div>
      </div>

      {/* Progress bar */}
      <Card className={styles.progressCard}>
        <div className={styles.progressRow}>
          <Text size={300} weight="semibold">Bid Progress</Text>
          <Text size={300}>{progress}%</Text>
        </div>
        <ProgressBar
          value={progress / 100}
          color={progress >= 75 ? "success" : progress >= 40 ? "warning" : "error"}
        />
        <Text size={200} style={{ color: tokens.colorNeutralForeground3, marginTop: tokens.spacingVerticalXS }}>
          {progress < 25 && "Just getting started"}
          {progress >= 25 && progress < 50 && "In progress — key sections underway"}
          {progress >= 50 && progress < 75 && "Good progress — final sections remaining"}
          {progress >= 75 && progress < 100 && "Nearly there — review stage"}
          {progress === 100 && "Complete — ready for submission"}
        </Text>
      </Card>

      {/* Tab navigation */}
      <TabList
        selectedValue={activeTab}
        onTabSelect={(_, d) => setActiveTab(d.value as WorkspaceTab)}
      >
        <Tab value="overview" icon={<CheckmarkCircleRegular />}>Overview</Tab>
        <Tab value="team" icon={<PeopleRegular />}>Team</Tab>
        <Tab value="approvals" icon={<ClockRegular />}>Approvals</Tab>
        <Tab value="documents" icon={<DocumentRegular />}>Documents</Tab>
      </TabList>

      <div className={styles.tabContent}>
        {activeTab === "overview" && (
          <OverviewTab workspace={workspace ?? null} bidRequest={bidRequest} />
        )}
        {activeTab === "team" && (
          <TeamTab members={teamMembers} />
        )}
        {activeTab === "approvals" && (
          <ApprovalsTab approvals={approvals} />
        )}
        {activeTab === "documents" && (
          <DocumentsTab documents={documents} />
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Tab: Overview
// ---------------------------------------------------------------------------

function OverviewTab({ workspace, bidRequest }: { workspace: BidWorkspace | null; bidRequest: any }) {
  const ref = workspace?.ricoh_bidrequestid ?? bidRequest;

  const rows = [
    ["Reference", ref?.ricoh_bidreferencenumber ?? "—"],
    ["Title", ref?.ricoh_title ?? workspace?.ricoh_title ?? "—"],
    ["Created", workspace ? new Date(workspace.createdOn).toLocaleDateString("en-GB") : "—"],
    ["SharePoint", workspace?.ricoh_sharepointfolderurl ?? "Not yet configured"],
    ["Teams Channel", workspace?.ricoh_teamschannelurl ?? "Not yet configured"],
  ];

  return (
    <Card>
      {rows.map(([label, value]) => (
        <div key={label} style={{ display: "flex", gap: tokens.spacingHorizontalM, padding: `${tokens.spacingVerticalS} 0`, borderBottom: `1px solid ${tokens.colorNeutralStroke2}` }}>
          <Text style={{ minWidth: "180px", color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>{label}</Text>
          <Text size={300}>{value}</Text>
        </div>
      ))}
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Tab: Team
// ---------------------------------------------------------------------------

function TeamTab({ members }: { members: BidTeamMember[] }) {
  const styles = useStyles();

  return (
    <Card>
      <CardHeader
        header={<Text weight="semibold">Team Members</Text>}
        action={
          <Button appearance="primary" size="small" icon={<AddRegular />}>
            Add Member
          </Button>
        }
      />
      {members.length === 0 ? (
        <EmptyState title="No team members" description="Add team members to get started." />
      ) : (
        <Table className={styles.teamTable} aria-label="Team members">
          <TableHeader>
            <TableRow>
              <TableHeaderCell>Name</TableHeaderCell>
              <TableHeaderCell>Email</TableHeaderCell>
              <TableHeaderCell>Role</TableHeaderCell>
              <TableHeaderCell>Status</TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {members.map((m) => (
              <TableRow key={m.id}>
                <TableCell>
                  <TableCellLayout>
                    <Text weight="semibold" size={300}>{m.ricoh_userid.fullName}</Text>
                  </TableCellLayout>
                </TableCell>
                <TableCell>
                  <TableCellLayout>
                    <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                      {m.ricoh_userid.email}
                    </Text>
                  </TableCellLayout>
                </TableCell>
                <TableCell>
                  <TableCellLayout>
                    <Text size={300}>{TeamRoleLabel[m.ricoh_role as keyof typeof TeamRoleLabel]}</Text>
                  </TableCellLayout>
                </TableCell>
                <TableCell>
                  <TableCellLayout>
                    <Badge appearance="filled" color={m.ricoh_isactive ? "success" : "subtle"} size="small">
                      {m.ricoh_isactive ? "Active" : "Inactive"}
                    </Badge>
                  </TableCellLayout>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      )}
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Tab: Approvals
// ---------------------------------------------------------------------------

function ApprovalsTab({ approvals }: { approvals: BidApproval[] }) {
  const styles = useStyles();

  return (
    <Card>
      <CardHeader
        header={<Text weight="semibold">Approval Chain</Text>}
        action={
          <Button appearance="primary" size="small" icon={<AddRegular />}>
            Request Approval
          </Button>
        }
      />
      {approvals.length === 0 ? (
        <EmptyState title="No approvals" description="No approvals have been requested yet." />
      ) : (
        approvals.map((ap) => (
          <div key={ap.id} className={styles.approvalRow}>
            <div className={styles.stageNumber}>{ap.ricoh_approverstage}</div>
            <div style={{ flexGrow: 1 }}>
              <Text weight="semibold" size={300}>{ap.ricoh_title}</Text>
              <Text size={200} style={{ display: "block", color: tokens.colorNeutralForeground3 }}>
                {ap.ricoh_approverid.fullName} — {ap.ricoh_approverid.email}
              </Text>
              {ap.ricoh_comments && (
                <Text size={200} style={{ color: tokens.colorNeutralForeground2, fontStyle: "italic" }}>
                  "{ap.ricoh_comments}"
                </Text>
              )}
            </div>
            <div style={{ minWidth: "90px", textAlign: "right" }}>
              <ApprovalBadge status={ap.ricoh_status} />
              {ap.ricoh_respondeddate && (
                <Text size={100} style={{ display: "block", color: tokens.colorNeutralForeground3, marginTop: "4px" }}>
                  {new Date(ap.ricoh_respondeddate).toLocaleDateString("en-GB")}
                </Text>
              )}
            </div>
          </div>
        ))
      )}
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Tab: Documents
// ---------------------------------------------------------------------------

function DocumentsTab({ documents }: { documents: any[] }) {
  const styles = useStyles();

  return (
    <Card>
      <CardHeader
        header={<Text weight="semibold">Documents</Text>}
        action={
          <Button appearance="primary" size="small" icon={<AddRegular />}>
            Upload Document
          </Button>
        }
      />
      {documents.length === 0 ? (
        <EmptyState
          icon={<DocumentRegular />}
          title="No documents yet"
          description="Upload documents or link SharePoint files to this workspace."
        />
      ) : (
        documents.map((doc) => (
          <div key={doc.id} className={styles.docRow}>
            <DocumentRegular style={{ fontSize: "20px", color: tokens.colorNeutralForeground3, flexShrink: 0 }} />
            <div style={{ flexGrow: 1 }}>
              <Text weight="semibold" size={300}>{doc.ricoh_title}</Text>
              <Text size={200} style={{ display: "block", color: tokens.colorNeutralForeground3 }}>
                {doc.ricoh_documenttype} · v{doc.ricoh_version}
              </Text>
            </div>
            <Button
              appearance="subtle"
              size="small"
              icon={<LinkRegular />}
              as="a"
              href={doc.ricoh_sharepointurl}
              target="_blank"
              rel="noopener noreferrer"
            >
              Open
            </Button>
          </div>
        ))
      )}
    </Card>
  );
}
