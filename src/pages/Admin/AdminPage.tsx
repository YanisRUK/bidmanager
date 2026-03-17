/**
 * Admin Page
 *
 * Configuration for bid types, routing rules, user roles, and system settings.
 */

import { useState } from "react";
import {
  makeStyles,
  shorthands,
  tokens,
  Text,
  Card,
  CardHeader,
  Button,
  Table,
  TableHeader,
  TableHeaderCell,
  TableBody,
  TableRow,
  TableCell,
  TableCellLayout,
  Badge,
  Tab,
  TabList,
  Input,
  Field,
  Divider,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import {
  EditRegular,
  AddRegular,
  SettingsRegular,
  PeopleRegular,
  ArrowRoutingRegular,
  InfoRegular,
} from "@fluentui/react-icons";
import { PageHeader } from "../../components/common/PageHeader";
import { useDataverse } from "../../hooks/useDataverse";
import { dataverseClient } from "../../lib/dataverse";
import { BidTypeCode, BidTypeLabel } from "../../types/dataverse";

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

const useStyles = makeStyles({
  tabContent: {
    paddingTop: tokens.spacingVerticalL,
  },
  sectionCard: {
    marginBottom: tokens.spacingVerticalL,
  },
  settingRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    ...shorthands.padding(tokens.spacingVerticalM, "0"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ":last-child": { borderBottom: "none" },
  },
  settingInfo: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
    maxWidth: "500px",
  },
});

type AdminTab = "bid-types" | "routing" | "users" | "system";

// ---------------------------------------------------------------------------
// Page
// ---------------------------------------------------------------------------

export function AdminPage() {
  const styles = useStyles();
  const [activeTab, setActiveTab] = useState<AdminTab>("bid-types");

  return (
    <div>
      <PageHeader
        title="Admin"
        subtitle="System configuration and management"
      />

      <MessageBar intent="info" style={{ marginBottom: tokens.spacingVerticalL }}>
        <MessageBarBody>
          Admin settings are applied globally. Changes to routing rules will affect new bid submissions immediately.
        </MessageBarBody>
      </MessageBar>

      <TabList
        selectedValue={activeTab}
        onTabSelect={(_, d) => setActiveTab(d.value as AdminTab)}
      >
        <Tab value="bid-types" icon={<SettingsRegular />}>Bid Types</Tab>
        <Tab value="routing" icon={<ArrowRoutingRegular />}>Routing Rules</Tab>
        <Tab value="users" icon={<PeopleRegular />}>User Roles</Tab>
        <Tab value="system" icon={<InfoRegular />}>System</Tab>
      </TabList>

      <div className={styles.tabContent}>
        {activeTab === "bid-types" && <BidTypesTab />}
        {activeTab === "routing" && <RoutingTab />}
        {activeTab === "users" && <UsersTab />}
        {activeTab === "system" && <SystemTab />}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Tab: Bid Types
// ---------------------------------------------------------------------------

function BidTypesTab() {
  const { data: bidTypes, isLoading } = useDataverse(
    () => dataverseClient.getBidTypes(),
    []
  );

  return (
    <Card>
      <CardHeader
        header={<Text weight="semibold">Bid Type Configuration</Text>}
        action={
          <Button appearance="primary" size="small" icon={<AddRegular />} disabled>
            Add Type
          </Button>
        }
      />
      {isLoading ? (
        <Text>Loading...</Text>
      ) : (
        <Table aria-label="Bid types">
          <TableHeader>
            <TableRow>
              <TableHeaderCell>Name</TableHeaderCell>
              <TableHeaderCell>Routing Team</TableHeaderCell>
              <TableHeaderCell>Routing Email</TableHeaderCell>
              <TableHeaderCell>SLA (days)</TableHeaderCell>
              <TableHeaderCell />
            </TableRow>
          </TableHeader>
          <TableBody>
            {(bidTypes ?? []).map((bt) => (
              <TableRow key={bt.id}>
                <TableCell>
                  <TableCellLayout>
                    <Text weight="semibold" size={300}>{bt.cr5ab_name}</Text>
                    <Text size={200} style={{ display: "block", color: tokens.colorNeutralForeground3 }}>
                      {bt.cr5ab_description}
                    </Text>
                  </TableCellLayout>
                </TableCell>
                <TableCell>
                  <TableCellLayout>
                    <Badge appearance="outline" color="informative" size="medium">
                      {bt.cr5ab_routingteam}
                    </Badge>
                  </TableCellLayout>
                </TableCell>
                <TableCell>
                  <TableCellLayout>
                    <Text size={200}>{bt.cr5ab_routingemail}</Text>
                  </TableCellLayout>
                </TableCell>
                <TableCell>
                  <TableCellLayout>
                    <Text size={300}>{bt.cr5ab_slaresponsedays}</Text>
                  </TableCellLayout>
                </TableCell>
                <TableCell>
                  <TableCellLayout>
                    <Button appearance="subtle" size="small" icon={<EditRegular />} disabled>
                      Edit
                    </Button>
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
// Tab: Routing Rules
// ---------------------------------------------------------------------------

const ROUTING_RULES = [
  { id: 1, bidType: BidTypeCode.SupplierQuestionnaire, condition: "On submission", action: "Route to Procurement queue + email", active: true },
  { id: 2, bidType: BidTypeCode.SalesLed,              condition: "On submission", action: "Route to Sales Bids queue + Teams notification", active: true },
  { id: 3, bidType: BidTypeCode.SMEToQualify,          condition: "On submission", action: "Route to SME Pool queue + email", active: true },
  { id: 4, bidType: BidTypeCode.BidManagement,         condition: "On submission", action: "Create workspace + notify Bid Management", active: true },
  { id: 5, bidType: null,                              condition: "Deadline within 7 days", action: "Send reminder to Bid Manager + team", active: true },
  { id: 6, bidType: null,                              condition: "Status changed to Won/Lost", action: "Notify submitter + update pipeline report", active: false },
];

function RoutingTab() {
  return (
    <Card>
      <CardHeader
        header={<Text weight="semibold">Power Automate Routing Rules</Text>}
        description={
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
            These rules trigger Power Automate flows. Edit flows directly in Power Automate.
          </Text>
        }
      />
      <Table aria-label="Routing rules">
        <TableHeader>
          <TableRow>
            <TableHeaderCell>Bid Type</TableHeaderCell>
            <TableHeaderCell>Condition</TableHeaderCell>
            <TableHeaderCell>Action</TableHeaderCell>
            <TableHeaderCell>Status</TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {ROUTING_RULES.map((rule) => (
            <TableRow key={rule.id}>
              <TableCell>
                <TableCellLayout>
                  <Text size={300}>
                    {rule.bidType !== null
                      ? BidTypeLabel[rule.bidType]
                      : "All types"}
                  </Text>
                </TableCellLayout>
              </TableCell>
              <TableCell>
                <TableCellLayout>
                  <Text size={200}>{rule.condition}</Text>
                </TableCellLayout>
              </TableCell>
              <TableCell>
                <TableCellLayout>
                  <Text size={200}>{rule.action}</Text>
                </TableCellLayout>
              </TableCell>
              <TableCell>
                <TableCellLayout>
                  <Badge appearance="filled" color={rule.active ? "success" : "subtle"} size="small">
                    {rule.active ? "Active" : "Inactive"}
                  </Badge>
                </TableCellLayout>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Tab: Users
// ---------------------------------------------------------------------------

const MOCK_USERS = [
  { id: "u1", name: "Sarah Mitchell", email: "s.mitchell@ricoh.co.uk", role: "Bid Manager", department: "Bid Management" },
  { id: "u2", name: "James O'Brien",  email: "j.obrien@ricoh.co.uk",  role: "Subject Matter Expert", department: "Technical" },
  { id: "u3", name: "Priya Sharma",   email: "p.sharma@ricoh.co.uk",  role: "Approver", department: "Commercial" },
  { id: "u4", name: "Tom Watkins",    email: "t.watkins@ricoh.co.uk", role: "Bid Coordinator", department: "Bid Management" },
  { id: "u5", name: "Helen Cross",    email: "h.cross@ricoh.co.uk",   role: "Approver", department: "Executive" },
];

function UsersTab() {
  return (
    <Card>
      <CardHeader
        header={<Text weight="semibold">User Roles & Permissions</Text>}
        description={<Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Managed via Azure AD groups. Sync runs nightly.</Text>}
        action={
          <Button appearance="primary" size="small" icon={<AddRegular />} disabled>
            Add User
          </Button>
        }
      />
      <Table aria-label="Users">
        <TableHeader>
          <TableRow>
            <TableHeaderCell>Name</TableHeaderCell>
            <TableHeaderCell>Email</TableHeaderCell>
            <TableHeaderCell>Role</TableHeaderCell>
            <TableHeaderCell>Department</TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {MOCK_USERS.map((u) => (
            <TableRow key={u.id}>
              <TableCell><TableCellLayout><Text weight="semibold" size={300}>{u.name}</Text></TableCellLayout></TableCell>
              <TableCell><TableCellLayout><Text size={200}>{u.email}</Text></TableCellLayout></TableCell>
              <TableCell><TableCellLayout><Badge appearance="outline" color="informative" size="medium">{u.role}</Badge></TableCellLayout></TableCell>
              <TableCell><TableCellLayout><Text size={200}>{u.department}</Text></TableCellLayout></TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Tab: System
// ---------------------------------------------------------------------------

function SystemTab() {
  const styles = useStyles();
  const [orgName, setOrgName] = useState("Ricoh UK");
  const [dvEnv, setDvEnv] = useState("ricoh-prod.crm11.dynamics.com");
  const [spSite, setSpSite] = useState("https://ricoh.sharepoint.com/sites/bids");
  const [saved, setSaved] = useState(false);

  function handleSave() {
    // TODO: persist to app config / environment variables
    setSaved(true);
    setTimeout(() => setSaved(false), 3000);
  }

  return (
    <div>
      <Card className={styles.sectionCard}>
        <CardHeader header={<Text weight="semibold">General Settings</Text>} />
        <div style={{ display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM }}>
          <Field label="Organisation Name">
            <Input value={orgName} onChange={(_, d) => setOrgName(d.value)} />
          </Field>
          <Field label="Dataverse Environment URL">
            <Input value={dvEnv} onChange={(_, d) => setDvEnv(d.value)} />
          </Field>
          <Field label="SharePoint Site URL">
            <Input value={spSite} onChange={(_, d) => setSpSite(d.value)} />
          </Field>
        </div>
        <Divider style={{ margin: `${tokens.spacingVerticalL} 0` }} />
        <div style={{ display: "flex", justifyContent: "flex-end", gap: tokens.spacingHorizontalS }}>
          {saved && (
            <MessageBar intent="success" style={{ flexGrow: 1 }}>
              <MessageBarBody>Settings saved.</MessageBarBody>
            </MessageBar>
          )}
          <Button appearance="primary" onClick={handleSave}>Save Settings</Button>
        </div>
      </Card>

      <Card>
        <CardHeader header={<Text weight="semibold">System Information</Text>} />
        {[
          ["App Version", "0.1.0"],
          ["Environment", import.meta.env.DEV ? "Development (mock data)" : "Production"],
          ["Power Apps SDK", import.meta.env.DEV ? "Not loaded (dev mode)" : "Loaded"],
          ["Dataverse Publisher Prefix", "cr5ab_"],
        ].map(([label, value]) => (
          <div key={label} className={styles.settingRow}>
            <div className={styles.settingInfo}>
              <Text size={300} weight="semibold">{label}</Text>
            </div>
            <Text size={300} style={{ color: tokens.colorNeutralForeground2 }}>{value}</Text>
          </div>
        ))}
      </Card>
    </div>
  );
}
