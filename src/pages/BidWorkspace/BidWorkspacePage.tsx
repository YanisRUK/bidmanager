/**
 * Bid Workspace Page
 *
 * Individual workspace for an active bid with full feature set:
 *  - Status transition controls
 *  - Overview tab
 *  - TOR (Table of Responsibility) editable grid — grouped by section, collapsible
 *  - Team tab with Add Member dialog
 *  - Approvals tab with Request Approval dialog
 *  - Documents tab with Link + Upload actions
 */

import React, { useState, useCallback, useRef, useEffect } from "react";
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
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Field,
  Input,
  Textarea,
  Select,
  Spinner,
  MessageBar,
  MessageBarBody,
  Tooltip,
} from "@fluentui/react-components";
import {
  ArrowLeftRegular,
  PeopleRegular,
  DocumentRegular,
  CheckmarkCircleRegular,
  ClockRegular,
  LinkRegular,
  AddRegular,
  EditRegular,
  DeleteRegular,
  TableRegular,
  CheckmarkRegular,
  DismissRegular,
  ArrowUpRegular,
  ArrowDownRegular,
  ChevronDownRegular,
  ChevronRightRegular,
  ArrowUploadRegular,
  LockClosedRegular,
  PersonSearchRegular,
  QuestionCircleRegular,
  ChatRegular,
  ArrowReplyRegular,
  CheckmarkStarburstRegular,
} from "@fluentui/react-icons";
import { PageHeader } from "../../components/common/PageHeader";
import { StatusBadge } from "../../components/common/StatusBadge";
import { LoadingState } from "../../components/common/LoadingState";
import { ErrorState } from "../../components/common/ErrorState";
import { EmptyState } from "../../components/common/EmptyState";
import { useDataverse } from "../../hooks/useDataverse";
import { dataverseClient } from "../../lib/dataverse";
import {
  ApprovalStatus,
  TeamRoleLabel,
  TeamRole,
  BidStatus,
  BidStatusLabel,
  TorAnsweredStatus,
  TorAnsweredStatusLabel,
  TorAnsweredStatusColor,
  DocumentCategoryLabel,
  OpportunityStage,
  OpportunityStageLabel,
  OpportunityStageColor,
  QualificationOutcome,
  QualificationOutcomeLabel,
  QualificationOutcomeColor,
  ClarificationStatus,
  ClarificationStatusLabel,
  ClarificationStatusColor,
} from "../../types/dataverse";
import type {
  BidRequest,
  BidWorkspace,
  BidRoleAssignment,
  BidApproval,
  TorItem,
  BidDocumentRecord,
  BidClarification,
  DocumentCategory,
  DataverseUser,
  OafData,
} from "../../types/dataverse";
import { EMPTY_OAF } from "../../types/dataverse";

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

const useStyles = makeStyles({
  headerGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr 1fr 1fr",
    gap: tokens.spacingVerticalM,
    marginBottom: tokens.spacingVerticalL,
    "@media (max-width: 900px)": { gridTemplateColumns: "1fr 1fr" },
  },
  metaItem: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS },
  tabContent: { paddingTop: tokens.spacingVerticalL },
  progressCard: { marginBottom: tokens.spacingVerticalL },
  progressRow: { display: "flex", justifyContent: "space-between", marginBottom: tokens.spacingVerticalS },
  teamTable: { width: "100%" },
  approvalRow: {
    display: "flex", alignItems: "flex-start", gap: tokens.spacingHorizontalM,
    ...shorthands.padding(tokens.spacingVerticalM, "0"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ":last-child": { borderBottom: "none" },
  },
  stageNumber: {
    width: "32px", height: "32px", borderRadius: "50%",
    display: "flex", alignItems: "center", justifyContent: "center",
    backgroundColor: tokens.colorNeutralBackground3,
    fontWeight: tokens.fontWeightSemibold, fontSize: tokens.fontSizeBase200, flexShrink: 0,
  },
  stageNumberApproved: { backgroundColor: tokens.colorPaletteGreenBackground2, color: tokens.colorPaletteGreenForeground1 },
  stageNumberPending:  { backgroundColor: tokens.colorPaletteMarigoldBackground2, color: tokens.colorPaletteMarigoldForeground2 },

  // ── TOR styles ──────────────────────────────────────────────────────────
  torToolbar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-end",
    marginBottom: tokens.spacingVerticalM,
    gap: tokens.spacingHorizontalS,
  },
  torSection: { marginBottom: tokens.spacingVerticalL },
  torSectionHeader: {
    display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS,
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM),
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    marginBottom: tokens.spacingVerticalXS,
    cursor: "pointer",
    border: "none",
    width: "100%",
    textAlign: "left",
    ":hover": { backgroundColor: tokens.colorNeutralBackground3 },
  },
  torSectionProgress: {
    marginLeft: "auto",
    display: "flex", alignItems: "center", gap: tokens.spacingHorizontalM,
  },
  torRow: {
    display: "grid",
    gridTemplateColumns: "55px 1fr 110px 120px 90px 80px 68px",
    gap: tokens.spacingHorizontalS,
    alignItems: "center",
    ...shorthands.padding("6px", tokens.spacingHorizontalS),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ":hover": { backgroundColor: tokens.colorNeutralBackground2 },
  },
  torRowAnswered:  { backgroundColor: tokens.colorPaletteGreenBackground1 },
  torRowPartial:   { backgroundColor: tokens.colorPaletteMarigoldBackground1 },
  torRowNo:        { backgroundColor: tokens.colorPaletteRedBackground1 },
  torHeaderRow: {
    display: "grid",
    gridTemplateColumns: "55px 1fr 110px 120px 90px 80px 68px",
    gap: tokens.spacingHorizontalS,
    ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalS),
    borderBottom: `2px solid ${tokens.colorNeutralStroke1}`,
    marginBottom: tokens.spacingVerticalXS,
  },

  // Status picker inline
  statusPill: {
    display: "inline-flex", alignItems: "center",
    ...shorthands.padding("2px", tokens.spacingHorizontalS),
    borderRadius: tokens.borderRadiusCircular,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    cursor: "pointer",
    border: "none",
    gap: "4px",
  },
  statusPillAnswered: { backgroundColor: tokens.colorPaletteGreenBackground2, color: tokens.colorPaletteGreenForeground1 },
  statusPillPartial:  { backgroundColor: tokens.colorPaletteMarigoldBackground2, color: tokens.colorPaletteMarigoldForeground2 },
  statusPillNo:       { backgroundColor: tokens.colorPaletteRedBackground2, color: tokens.colorPaletteRedForeground3 },

  // People picker
  peoplePicker: {
    position: "relative",
  },
  peoplePickerDropdown: {
    position: "absolute",
    top: "100%",
    left: 0,
    right: 0,
    zIndex: 100,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow8,
    maxHeight: "200px",
    overflowY: "auto",
    marginTop: "2px",
  },
  peoplePickerOption: {
    display: "flex",
    flexDirection: "column",
    ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalM),
    cursor: "pointer",
    ":hover": { backgroundColor: tokens.colorNeutralBackground2 },
  },
  peoplePickerSelected: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
    ...shorthands.padding("4px", tokens.spacingHorizontalS),
    backgroundColor: tokens.colorBrandBackground2,
    borderRadius: tokens.borderRadiusCircular,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorBrandForeground1,
    border: `1px solid ${tokens.colorBrandStroke1}`,
    marginTop: "2px",
    cursor: "pointer",
  },

  // Section picker (combobox-style)
  sectionPicker: { position: "relative" },
  sectionPickerDropdown: {
    position: "absolute",
    top: "100%",
    left: 0,
    right: 0,
    zIndex: 100,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow8,
    marginTop: "2px",
    overflow: "hidden",
  },
  sectionPickerOption: {
    display: "block",
    width: "100%",
    ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalM),
    cursor: "pointer",
    textAlign: "left",
    background: "none",
    border: "none",
    fontSize: tokens.fontSizeBase300,
    ":hover": { backgroundColor: tokens.colorNeutralBackground2 },
  },
  sectionPickerNewOption: {
    display: "block",
    width: "100%",
    ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalM),
    cursor: "pointer",
    textAlign: "left",
    background: "none",
    border: "none",
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorBrandForeground1,
    fontWeight: tokens.fontWeightSemibold,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    ":hover": { backgroundColor: tokens.colorBrandBackground2 },
  },

  // Weighting toggle
  weightingToggle: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalS,
  },
  weightingBtn: {
    flex: 1,
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM),
    border: `1.5px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    cursor: "pointer",
    background: "none",
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightRegular,
    color: tokens.colorNeutralForeground2,
    transition: "all 0.1s ease",
    ":hover": { ...shorthands.borderColor(tokens.colorBrandStroke1), color: tokens.colorBrandForeground1 },
  },
  weightingBtnActive: {
    ...shorthands.borderColor(tokens.colorBrandBackground),
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    fontWeight: tokens.fontWeightSemibold,
  },

  // Inline edit comment toggle
  commentToggle: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    cursor: "pointer",
    background: "none",
    border: "none",
    padding: 0,
    marginTop: "2px",
    textDecoration: "underline",
    ":hover": { color: tokens.colorBrandForeground1 },
  },

  // Status transition bar
  statusRow: {
    display: "flex", alignItems: "center", gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
    ...shorthands.padding(tokens.spacingVerticalM),
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    flexWrap: "wrap",
  },

  // Doc tabs
  docTabRow: {
    display: "flex", gap: tokens.spacingHorizontalM, marginBottom: tokens.spacingVerticalM,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ...shorthands.padding("0", "0", tokens.spacingVerticalS, "0"),
  },
  docTab: {
    cursor: "pointer",
    ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalM),
    borderRadius: tokens.borderRadiusSmall,
    border: "none",
    backgroundColor: "transparent",
    color: tokens.colorNeutralForeground2,
    fontWeight: tokens.fontWeightRegular,
    fontSize: tokens.fontSizeBase300,
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalXS,
    ":hover": { backgroundColor: tokens.colorNeutralBackground3 },
  },
  docTabActive: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    fontWeight: tokens.fontWeightSemibold,
  },
  docRow: {
    display: "flex", alignItems: "center", gap: tokens.spacingHorizontalM,
    ...shorthands.padding(tokens.spacingVerticalS, "0"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ":last-child": { borderBottom: "none" },
  },
  uploadProgress: {
    display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS,
    ...shorthands.padding(tokens.spacingVerticalM),
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    marginTop: tokens.spacingVerticalS,
  },

  // ── Qualification panel ─────────────────────────────────────────────────
  qualPanel: {
    ...shorthands.padding(tokens.spacingVerticalL),
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusLarge,
    border: `1.5px solid ${tokens.colorBrandStroke1}`,
    marginBottom: tokens.spacingVerticalL,
  },
  qualPanelHeader: {
    display: "flex", alignItems: "center", gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalM,
    cursor: "pointer",
    background: "none",
    border: "none",
    width: "100%",
    textAlign: "left",
    ...shorthands.padding("0"),
  },
  qualDecisionGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))",
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalM,
  },
  qualDecisionBtn: {
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalM),
    border: `2px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    cursor: "pointer",
    background: "none",
    textAlign: "center" as const,
    transition: "all 0.1s ease",
    ":hover": { ...shorthands.borderColor(tokens.colorBrandStroke1) },
  },
  qualDecisionBtnSelected: {
    ...shorthands.borderColor(tokens.colorBrandBackground),
    backgroundColor: tokens.colorBrandBackground2,
  },
  qualSummaryBanner: {
    display: "flex", alignItems: "center", gap: tokens.spacingHorizontalM,
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM),
    backgroundColor: tokens.colorPaletteGreenBackground2,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorPaletteGreenBorder1}`,
    marginBottom: tokens.spacingVerticalL,
    cursor: "pointer",
  },
  oafSection: {
    marginBottom: tokens.spacingVerticalM,
  },
  oafGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalM,
    "@media (max-width: 700px)": { gridTemplateColumns: "1fr" },
  },
  oafFullWidth: { gridColumn: "1 / -1" },

  // ── Clarifications tab ──────────────────────────────────────────────────
  cqRow: {
    display: "flex", alignItems: "flex-start", gap: tokens.spacingHorizontalM,
    ...shorthands.padding(tokens.spacingVerticalM, "0"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ":last-child": { borderBottom: "none" },
  },
  cqDirectionBadge: {
    flexShrink: 0,
    width: "24px", height: "24px",
    borderRadius: "50%",
    display: "flex", alignItems: "center", justifyContent: "center",
    fontSize: "12px",
  },
  cqDirectionUs:       { backgroundColor: tokens.colorBrandBackground2, color: tokens.colorBrandForeground1 },
  cqDirectionCustomer: { backgroundColor: tokens.colorPaletteMarigoldBackground2, color: tokens.colorPaletteMarigoldForeground2 },
});

// ---------------------------------------------------------------------------
// Helpers
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

function daysUntil(iso: string): number {
  return Math.ceil((new Date(iso).getTime() - Date.now()) / 86400000);
}

function fmtDate(iso?: string): string {
  if (!iso) return "—";
  return new Date(iso).toLocaleDateString("en-GB", { day: "numeric", month: "short" });
}

/** Derive a section prefix number, e.g. "Section 2 — Technical" → "2" */
function sectionPrefix(sectionName: string): string {
  const m = sectionName.match(/^(?:section\s*)?(\d+)/i);
  return m ? m[1] : "";
}

/** Given existing items in a section, produce the next question number, e.g. "2.3" */
function nextQuestionNumber(sectionName: string, sectionItems: TorItem[]): string {
  const prefix = sectionPrefix(sectionName);
  if (!prefix) return "";
  const nums = sectionItems
    .map((i) => {
      const m = i.cr5ab_questionnumber?.match(/^(\d+)\.(\d+)$/);
      return m && m[1] === prefix ? parseInt(m[2], 10) : NaN;
    })
    .filter((n) => !isNaN(n));
  const next = nums.length > 0 ? Math.max(...nums) + 1 : 1;
  return `${prefix}.${next}`;
}

// ---------------------------------------------------------------------------
// Shared user list (mock — in production replace with Office 365 Users search)
// ---------------------------------------------------------------------------

const AVAILABLE_USERS: DataverseUser[] = [
  { id: "u1", fullName: "Sarah Mitchell",  email: "s.mitchell@ricoh.co.uk" },
  { id: "u2", fullName: "James O'Brien",   email: "j.obrien@ricoh.co.uk" },
  { id: "u3", fullName: "Priya Sharma",    email: "p.sharma@ricoh.co.uk" },
  { id: "u4", fullName: "Tom Watkins",     email: "t.watkins@ricoh.co.uk" },
  { id: "u5", fullName: "Helen Cross",     email: "h.cross@ricoh.co.uk" },
  { id: "u6", fullName: "Louise Brennan",  email: "l.brennan@ricoh.co.uk" },
];

// ---------------------------------------------------------------------------
// PeoplePicker — search-filterable user selector
// ---------------------------------------------------------------------------

interface PeoplePickerProps {
  value: DataverseUser | null;
  onChange: (user: DataverseUser | null) => void;
  placeholder?: string;
}

function PeoplePicker({ value, onChange, placeholder = "Search by name..." }: PeoplePickerProps) {
  const styles = useStyles();
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  // Close on outside click
  useEffect(() => {
    function handle(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener("mousedown", handle);
    return () => document.removeEventListener("mousedown", handle);
  }, []);

  const filtered = AVAILABLE_USERS.filter(
    (u) =>
      u.fullName.toLowerCase().includes(query.toLowerCase()) ||
      u.email.toLowerCase().includes(query.toLowerCase())
  );

  if (value) {
    return (
      <div>
        <button
          className={styles.peoplePickerSelected}
          onClick={() => { onChange(null); setQuery(""); }}
          type="button"
          title="Click to change"
        >
          <PersonSearchRegular style={{ fontSize: "12px" }} />
          <span>{value.fullName}</span>
          <DismissRegular style={{ fontSize: "10px", opacity: 0.7 }} />
        </button>
      </div>
    );
  }

  return (
    <div className={styles.peoplePicker} ref={ref}>
      <Input
        placeholder={placeholder}
        value={query}
        contentBefore={<PersonSearchRegular />}
        onChange={(_, d) => { setQuery(d.value); setOpen(true); }}
        onFocus={() => setOpen(true)}
      />
      {open && filtered.length > 0 && (
        <div className={styles.peoplePickerDropdown}>
          {filtered.map((u) => (
            <div
              key={u.id}
              className={styles.peoplePickerOption}
              onMouseDown={() => { onChange(u); setQuery(""); setOpen(false); }}
            >
              <Text size={300} weight="semibold">{u.fullName}</Text>
              <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>{u.email}</Text>
            </div>
          ))}
        </div>
      )}
      {open && filtered.length === 0 && query.length > 0 && (
        <div className={styles.peoplePickerDropdown}>
          <div style={{ padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}` }}>
            <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>No users found</Text>
          </div>
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// SectionPicker — combobox: pick existing section or type a new one
// ---------------------------------------------------------------------------

interface SectionPickerProps {
  value: string;
  onChange: (val: string) => void;
  existingSections: string[];
}

function SectionPicker({ value, onChange, existingSections }: SectionPickerProps) {
  const styles = useStyles();
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    function handle(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener("mousedown", handle);
    return () => document.removeEventListener("mousedown", handle);
  }, []);

  const filtered = existingSections.filter((s) =>
    s.toLowerCase().includes(value.toLowerCase())
  );
  const showNew = value.trim().length > 0 && !existingSections.includes(value.trim());

  return (
    <div className={styles.sectionPicker} ref={ref}>
      <Input
        placeholder="Pick existing section or type a new one..."
        value={value}
        onChange={(_, d) => { onChange(d.value); setOpen(true); }}
        onFocus={() => setOpen(true)}
      />
      {open && (filtered.length > 0 || showNew) && (
        <div className={styles.sectionPickerDropdown}>
          {filtered.map((s) => (
            <button
              key={s}
              className={styles.sectionPickerOption}
              type="button"
              onMouseDown={() => { onChange(s); setOpen(false); }}
            >
              {s}
            </button>
          ))}
          {showNew && (
            <button
              className={styles.sectionPickerNewOption}
              type="button"
              onMouseDown={() => { onChange(value.trim()); setOpen(false); }}
            >
              + Create "{value.trim()}"
            </button>
          )}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// WeightingField — toggle between Percentage and Pass/Fail
// ---------------------------------------------------------------------------

type WeightingMode = "percentage" | "passfail";

interface WeightingFieldProps {
  value: string;
  onChange: (val: string) => void;
}

function WeightingField({ value, onChange }: WeightingFieldProps) {
  const styles = useStyles();

  // Detect current mode from stored value
  const isPassFail = value === "Pass/Fail";
  const mode: WeightingMode = isPassFail ? "passfail" : "percentage";
  const numericVal = (!isPassFail && value.endsWith("%")) ? value.slice(0, -1) : (!isPassFail ? value : "");

  function setMode(m: WeightingMode) {
    if (m === "passfail") onChange("Pass/Fail");
    else onChange("");
  }

  return (
    <div>
      <div className={styles.weightingToggle}>
        <button
          type="button"
          className={`${styles.weightingBtn} ${mode === "percentage" ? styles.weightingBtnActive : ""}`}
          onClick={() => setMode("percentage")}
        >
          %&nbsp;Percentage
        </button>
        <button
          type="button"
          className={`${styles.weightingBtn} ${mode === "passfail" ? styles.weightingBtnActive : ""}`}
          onClick={() => setMode("passfail")}
        >
          Pass / Fail
        </button>
      </div>
      {mode === "percentage" && (
        <Input
          type="number"
          min={0}
          max={100}
          placeholder="0–100"
          value={numericVal}
          onChange={(_, d) => {
            const n = Math.min(100, Math.max(0, Number(d.value)));
            onChange(d.value === "" ? "" : `${n}%`);
          }}
          contentAfter={<Text size={200}>%</Text>}
        />
      )}
      {mode === "passfail" && (
        <div style={{
          padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
          backgroundColor: tokens.colorNeutralBackground2,
          borderRadius: tokens.borderRadiusMedium,
          border: `1px solid ${tokens.colorNeutralStroke2}`,
        }}>
          <Text size={200} style={{ color: tokens.colorNeutralForeground2 }}>Pass / Fail — no numeric weighting</Text>
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// StatusTransitionBar
// ---------------------------------------------------------------------------

const STATUS_FLOW: BidStatus[] = [
  BidStatus.Draft, BidStatus.Submitted, BidStatus.InReview,
  BidStatus.Qualified, BidStatus.InProgress, BidStatus.Won,
];

// ---------------------------------------------------------------------------
// QualificationPanel — OAF form + decision (shown when bid needs qualifying)
// ---------------------------------------------------------------------------

interface QualificationPanelProps {
  bidRequest: BidRequest;
  onDecisionRecorded: (updated: Partial<BidRequest>) => void;
}

function QualificationPanel({ bidRequest, onDecisionRecorded }: QualificationPanelProps) {
  const styles = useStyles();

  // If already decided, show a collapsed read-only summary
  const isDecided = bidRequest.cr5ab_qualificationoutcome !== undefined && bidRequest.cr5ab_qualificationoutcome !== null;
  const [expanded, setExpanded] = useState(!isDecided);
  const [oafExpanded, setOafExpanded] = useState(false);

  // OAF form state
  const parsedOaf: OafData = React.useMemo(() => {
    try { return bidRequest.cr5ab_oafdata ? JSON.parse(bidRequest.cr5ab_oafdata) : { ...EMPTY_OAF }; }
    catch { return { ...EMPTY_OAF }; }
  }, [bidRequest.cr5ab_oafdata]);

  const [oaf, setOaf] = useState<OafData>(parsedOaf);
  function updateOaf<K extends keyof OafData>(key: K, value: OafData[K]) {
    setOaf((prev) => ({ ...prev, [key]: value }));
  }

  // Decision state
  const [outcome, setOutcome] = useState<QualificationOutcome | null>(bidRequest.cr5ab_qualificationoutcome ?? null);
  const [allocatedBidManager, setAllocatedBidManager] = useState<DataverseUser | null>(bidRequest.cr5ab_assignedto ?? null);
  const [rationale, setRationale] = useState(bidRequest.cr5ab_qualificationrationale ?? "");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const needsBidManager = outcome === QualificationOutcome.BidManagement || outcome === QualificationOutcome.LightSupport;

  async function handleRecordDecision() {
    if (!outcome) return;
    setBusy(true);
    setError(null);
    try {
      const changes: Partial<BidRequest> = {
        cr5ab_qualificationoutcome: outcome,
        cr5ab_qualificationrationale: rationale || undefined,
        cr5ab_qualifiedon: new Date().toISOString(),
        cr5ab_oafdata: JSON.stringify(oaf),
        cr5ab_opportunitystage: outcome === QualificationOutcome.NoBid
          ? OpportunityStage.NoBid
          : OpportunityStage.Allocated,
        cr5ab_status: outcome === QualificationOutcome.NoBid ? BidStatus.NoBid : BidStatus.Qualified,
        ...(allocatedBidManager ? { cr5ab_assignedto: allocatedBidManager } : {}),
      };
      await dataverseClient.updateBidRequest(bidRequest.id, changes);

      // Fire notification if allocating a bid manager
      if (allocatedBidManager && (outcome === QualificationOutcome.BidManagement || outcome === QualificationOutcome.LightSupport)) {
        await dataverseClient.triggerNotifyTeamFlow({
          bidWorkspaceId: bidRequest.cr5ab_bidworkspaceid?.id ?? bidRequest.id,
          recipientIds: [allocatedBidManager.id, bidRequest.cr5ab_submittedby.id],
          messageType: "qualification_decision",
          additionalContext: {
            outcome: QualificationOutcomeLabel[outcome],
            rationale: rationale,
            bidTitle: bidRequest.cr5ab_title,
          },
        });
      }

      onDecisionRecorded(changes);
      setExpanded(false);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to record decision");
    } finally { setBusy(false); }
  }

  // ── Read-only summary banner (post-decision) ───────────────────────────
  if (isDecided && !expanded) {
    const oc = bidRequest.cr5ab_qualificationoutcome!;
    return (
      <button className={styles.qualSummaryBanner} onClick={() => setExpanded(true)} type="button">
        <CheckmarkStarburstRegular style={{ fontSize: "20px", color: tokens.colorPaletteGreenForeground1, flexShrink: 0 }} />
        <div>
          <Text weight="semibold" size={300} style={{ color: tokens.colorPaletteGreenForeground1 }}>
            Qualified — {QualificationOutcomeLabel[oc]}
          </Text>
          {bidRequest.cr5ab_assignedto && (
            <Text size={200} style={{ display: "block", color: tokens.colorNeutralForeground2 }}>
              Allocated to {bidRequest.cr5ab_assignedto.fullName}
            </Text>
          )}
          <Text size={100} style={{ display: "block", color: tokens.colorNeutralForeground3 }}>
            {bidRequest.cr5ab_qualifiedby ? `by ${bidRequest.cr5ab_qualifiedby.fullName} · ` : ""}
            {bidRequest.cr5ab_qualifiedon ? fmtDate(bidRequest.cr5ab_qualifiedon) : ""}
          </Text>
        </div>
        <Text size={100} style={{ marginLeft: "auto", color: tokens.colorNeutralForeground3 }}>Click to review</Text>
      </button>
    );
  }

  // ── Full panel ─────────────────────────────────────────────────────────
  return (
    <div className={styles.qualPanel}>
      {/* Panel header */}
      <button className={styles.qualPanelHeader} onClick={() => isDecided && setExpanded(false)} type="button">
        <CheckmarkStarburstRegular style={{ fontSize: "22px", color: tokens.colorBrandForeground1, flexShrink: 0 }} />
        <div>
          <Text weight="semibold" size={400} style={{ color: tokens.colorBrandForeground1 }}>Qualification — Opportunity Assessment</Text>
          <Text size={200} style={{ display: "block", color: tokens.colorNeutralForeground2 }}>
            Complete the OAF and record the qualification decision before allocating to a bid manager.
          </Text>
        </div>
        {isDecided && (
          <DismissRegular style={{ marginLeft: "auto", fontSize: "16px", color: tokens.colorNeutralForeground3 }} />
        )}
      </button>

      {/* OAF section — collapsible */}
      <div className={styles.oafSection}>
        <button
          type="button"
          style={{ display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS, background: "none", border: "none", cursor: "pointer", marginBottom: tokens.spacingVerticalS, padding: 0 }}
          onClick={() => setOafExpanded((v) => !v)}
        >
          {oafExpanded ? <ChevronDownRegular style={{ fontSize: "14px" }} /> : <ChevronRightRegular style={{ fontSize: "14px" }} />}
          <Text weight="semibold" size={300}>Opportunity Assessment Form (OAF)</Text>
          <Badge appearance="outline" size="small" color={bidRequest.cr5ab_oafdata ? "success" : "warning"}>
            {bidRequest.cr5ab_oafdata ? "Completed" : "Not completed"}
          </Badge>
        </button>

        {oafExpanded && (
          <div className={styles.oafGrid}>
            <Field label="Services in Scope" className={styles.oafFullWidth}>
              <Textarea rows={3} placeholder="Describe all services and products in scope..." value={oaf.servicesInScope} onChange={(_, d) => updateOaf("servicesInScope", d.value)} />
            </Field>
            <Field label="Key Risks" className={styles.oafFullWidth}>
              <Textarea rows={3} placeholder="e.g. Complex framework, tight timeline, pricing format..." value={oaf.keyRisks} onChange={(_, d) => updateOaf("keyRisks", d.value)} />
            </Field>
            <Field label="Incumbent Supplier?">
              <Select value={oaf.hasIncumbent ? "yes" : "no"} onChange={(_, d) => updateOaf("hasIncumbent", d.value === "yes")}>
                <option value="no">No incumbent</option>
                <option value="yes">Yes — incumbent in place</option>
              </Select>
            </Field>
            {oaf.hasIncumbent && (
              <Field label="Incumbent Name">
                <Input placeholder="e.g. Xerox, Konica Minolta" value={oaf.incumbentName ?? ""} onChange={(_, d) => updateOaf("incumbentName", d.value)} />
              </Field>
            )}
            <Field label="Existing Relationships?">
              <Select value={oaf.relationshipsInPlace ? "yes" : "no"} onChange={(_, d) => updateOaf("relationshipsInPlace", d.value === "yes")}>
                <option value="no">No existing relationships</option>
                <option value="yes">Yes — relationships in place</option>
              </Select>
            </Field>
            <Field label="Estimated Bid Effort (days)">
              <Input type="number" placeholder="e.g. 15" value={oaf.estimatedBidEffortDays} onChange={(_, d) => updateOaf("estimatedBidEffortDays", d.value)} />
            </Field>
            <Field label="Recommended Resource / Bid Manager" className={styles.oafFullWidth}>
              <Input placeholder="e.g. Sarah Mitchell (Bid Manager), James O'Brien (Technical SME)" value={oaf.recommendedResource} onChange={(_, d) => updateOaf("recommendedResource", d.value)} />
            </Field>
            <Field label="Additional Notes" className={styles.oafFullWidth}>
              <Textarea rows={2} placeholder="Any other context for the qualification decision..." value={oaf.additionalNotes} onChange={(_, d) => updateOaf("additionalNotes", d.value)} />
            </Field>
          </div>
        )}
      </div>

      <Tooltip content="Select the qualification outcome" relationship="label">
        <Text weight="semibold" size={300} style={{ display: "block", marginBottom: tokens.spacingVerticalS }}>Qualification Decision</Text>
      </Tooltip>

      {/* Decision buttons */}
      <div className={styles.qualDecisionGrid}>
        {([
          QualificationOutcome.BidManagement,
          QualificationOutcome.SalesLed,
          QualificationOutcome.LightSupport,
          QualificationOutcome.NoBid,
        ] as QualificationOutcome[]).map((opt) => (
          <button
            key={opt}
            type="button"
            className={`${styles.qualDecisionBtn} ${outcome === opt ? styles.qualDecisionBtnSelected : ""}`}
            onClick={() => setOutcome(opt)}
          >
            <Text weight="semibold" size={300} style={{ display: "block", color: outcome === opt ? tokens.colorBrandForeground1 : undefined }}>
              {QualificationOutcomeLabel[opt]}
            </Text>
            <Badge appearance="filled" color={QualificationOutcomeColor[opt]} size="small" style={{ marginTop: "4px" }}>
              {opt === QualificationOutcome.NoBid ? "Decline" : opt === QualificationOutcome.LightSupport ? "Partial" : "Proceed"}
            </Badge>
          </button>
        ))}
      </div>

      {/* Allocate bid manager (if BidManagement or LightSupport) */}
      {needsBidManager && (
        <Field label="Allocate Bid Manager" style={{ marginBottom: tokens.spacingVerticalM }}>
          <PeoplePicker value={allocatedBidManager} onChange={setAllocatedBidManager} placeholder="Search for bid manager..." />
        </Field>
      )}

      {/* Rationale */}
      <Field label="Rationale / Notes" style={{ marginBottom: tokens.spacingVerticalM }}>
        <Textarea rows={3} placeholder="Explain the decision — key factors, risks, relationships..." value={rationale} onChange={(_, d) => setRationale(d.value)} />
      </Field>

      {error && <MessageBar intent="error" style={{ marginBottom: tokens.spacingVerticalS }}><MessageBarBody>{error}</MessageBarBody></MessageBar>}

      <div style={{ display: "flex", justifyContent: "flex-end", gap: tokens.spacingHorizontalS }}>
        {isDecided && (
          <Button appearance="subtle" onClick={() => setExpanded(false)}>Cancel</Button>
        )}
        <Button
          appearance="primary"
          icon={busy ? <Spinner size="tiny" /> : <CheckmarkStarburstRegular />}
          disabled={busy || !outcome || (needsBidManager && !allocatedBidManager)}
          onClick={handleRecordDecision}
        >
          {busy ? "Saving..." : "Record Decision"}
        </Button>
      </div>
    </div>
  );
}

function StatusTransitionBar({ workspace, onStatusChange }: { workspace: BidWorkspace; onStatusChange: (s: BidStatus) => void }) {
  const styles = useStyles();
  const [busy, setBusy] = useState(false);
  const currentIdx = STATUS_FLOW.indexOf(workspace.cr5ab_status);

  async function advance(newStatus: BidStatus) {
    setBusy(true);
    try {
      await dataverseClient.updateBidRequest(workspace.cr5ab_bidrequestid.id, { cr5ab_status: newStatus });
      await dataverseClient.updateBidWorkspace(workspace.id, { cr5ab_status: newStatus });
      onStatusChange(newStatus);
    } finally { setBusy(false); }
  }

  const canGoForward = currentIdx < STATUS_FLOW.length - 1 && workspace.cr5ab_status !== BidStatus.Won && workspace.cr5ab_status !== BidStatus.Lost && workspace.cr5ab_status !== BidStatus.Withdrawn;
  const canGoBack = currentIdx > 0 && workspace.cr5ab_status !== BidStatus.Won && workspace.cr5ab_status !== BidStatus.Lost;

  return (
    <div className={styles.statusRow}>
      <Text size={200} weight="semibold" style={{ color: tokens.colorNeutralForeground3 }}>STATUS</Text>
      <StatusBadge status={workspace.cr5ab_status} />
      <div style={{ display: "flex", gap: tokens.spacingHorizontalS, marginLeft: "auto" }}>
        {canGoBack && (
          <Button size="small" appearance="subtle" icon={<ArrowDownRegular />} disabled={busy} onClick={() => advance(STATUS_FLOW[currentIdx - 1])}>
            {busy ? <Spinner size="tiny" /> : `← ${BidStatusLabel[STATUS_FLOW[currentIdx - 1]]}`}
          </Button>
        )}
        {canGoForward && (
          <Button size="small" appearance="primary" icon={<ArrowUpRegular />} disabled={busy} onClick={() => advance(STATUS_FLOW[currentIdx + 1])}>
            {busy ? <Spinner size="tiny" /> : `Move to ${BidStatusLabel[STATUS_FLOW[currentIdx + 1]]}`}
          </Button>
        )}
        {!canGoForward && !canGoBack && (
          <Text size={200} style={{ color: tokens.colorNeutralForeground3, fontStyle: "italic" }}>
            {workspace.cr5ab_status === BidStatus.Won ? "Bid won — no further transitions" : "Bid closed"}
          </Text>
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Tab: Overview
// ---------------------------------------------------------------------------

function OverviewTab({ workspace, bidRequest }: {
  workspace: BidWorkspace | null;
  bidRequest: { cr5ab_bidreferencenumber?: string; cr5ab_title?: string; cr5ab_submissiondeadline?: string; cr5ab_expectedawarddate?: string; cr5ab_customername?: string; cr5ab_estimatedvalue?: number } | null;
}) {
  const ref = bidRequest;
  const rows: [string, React.ReactNode][] = [
    ["Reference",          ref?.cr5ab_bidreferencenumber ?? "—"],
    ["Customer",           ref?.cr5ab_customername ?? "—"],
    ["Submission Deadline",fmtDate(ref?.cr5ab_submissiondeadline)],
    ["Expected Award",     fmtDate(ref?.cr5ab_expectedawarddate)],
    ["Estimated Value",    ref?.cr5ab_estimatedvalue ? `£${ref.cr5ab_estimatedvalue.toLocaleString()}` : "—"],
    ["Created",            workspace ? new Date(workspace.createdOn).toLocaleDateString("en-GB") : "—"],
    ["SharePoint", workspace?.cr5ab_sharepointfolderurl
      ? <a href={workspace.cr5ab_sharepointfolderurl} target="_blank" rel="noopener noreferrer" style={{ color: tokens.colorBrandForeground1 }}>{workspace.cr5ab_sharepointfolderurl}</a>
      : "Not yet configured"],
    ["Teams Channel", workspace?.cr5ab_teamschannelurl
      ? <a href={workspace.cr5ab_teamschannelurl} target="_blank" rel="noopener noreferrer" style={{ color: tokens.colorBrandForeground1 }}>Open Teams channel</a>
      : "Not yet configured"],
  ];
  return (
    <Card>
      {rows.map(([label, value]) => (
        <div key={label as string} style={{ display: "flex", gap: tokens.spacingHorizontalM, padding: `${tokens.spacingVerticalS} 0`, borderBottom: `1px solid ${tokens.colorNeutralStroke2}` }}>
          <Text style={{ minWidth: "180px", color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>{label}</Text>
          <Text size={300}>{value}</Text>
        </div>
      ))}
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Tab: TOR
// ---------------------------------------------------------------------------

function TorTab({ workspaceId }: { workspaceId: string }) {
  const styles = useStyles();

  // ── fetch ──────────────────────────────────────────────────────────────
  const { data: torItems, isLoading, error, refresh } = useDataverse(
    () => dataverseClient.getTorItems(workspaceId),
    [workspaceId]
  );

  // ── local optimistic state for status changes ─────────────────────────
  const [localStatuses, setLocalStatuses] = useState<Record<string, TorAnsweredStatus>>({});

  // Merge server data with local optimistic overrides
  const items: TorItem[] = (torItems ?? []).map((t) =>
    localStatuses[t.id] !== undefined ? { ...t, cr5ab_answeredstatus: localStatuses[t.id] } : t
  );

  const total = items.length;
  const answered = items.filter((t) => t.cr5ab_answeredstatus === TorAnsweredStatus.Yes).length;
  const progress = total > 0 ? Math.round((answered / total) * 100) : 0;

  // ── group by section ───────────────────────────────────────────────────
  const sections = React.useMemo(() => {
    const map = new Map<string, TorItem[]>();
    for (const item of items) {
      const sec = item.cr5ab_section || "Uncategorised";
      if (!map.has(sec)) map.set(sec, []);
      map.get(sec)!.push(item);
    }
    return map;
  }, [items]);

  const sectionNames = Array.from(sections.keys());

  // ── collapsible section state ──────────────────────────────────────────
  const [collapsedSections, setCollapsedSections] = useState<Record<string, boolean>>({});
  function toggleSection(name: string) {
    setCollapsedSections((prev) => ({ ...prev, [name]: !prev[name] }));
  }

  // ── inline edit state ─────────────────────────────────────────────────
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editStatus, setEditStatus] = useState<TorAnsweredStatus>(TorAnsweredStatus.No);
  const [editComments, setEditComments] = useState("");
  const [showCommentFor, setShowCommentFor] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);

  function startEdit(item: TorItem) {
    setEditingId(item.id);
    setEditStatus(item.cr5ab_answeredstatus);
    setEditComments(item.cr5ab_comments ?? "");
    setShowCommentFor(null);
  }

  async function saveEdit(item: TorItem) {
    setSaving(true);
    try {
      await dataverseClient.updateTorItem(item.id, {
        cr5ab_answeredstatus: editStatus,
        cr5ab_comments: editComments,
      });
      // Apply optimistic status immediately so progress counters update
      setLocalStatuses((prev) => ({ ...prev, [item.id]: editStatus }));
      refresh();
      setEditingId(null);
      setShowCommentFor(null);
    } finally { setSaving(false); }
  }

  async function handleDelete(id: string) {
    if (!confirm("Delete this TOR row?")) return;
    await dataverseClient.deleteTorItem(id);
    setLocalStatuses((prev) => { const n = { ...prev }; delete n[id]; return n; });
    refresh();
  }

  // ── add row state ──────────────────────────────────────────────────────
  const [showAddRow, setShowAddRow] = useState(false);
  const [newSection, setNewSection] = useState("");
  const [newQNum, setNewQNum] = useState("");
  const [newQNumManual, setNewQNumManual] = useState(false);
  const [newDetail, setNewDetail] = useState("");
  const [newDepartment, setNewDepartment] = useState("");
  const [newWeighting, setNewWeighting] = useState("");
  const [newFirstDraft, setNewFirstDraft] = useState("");
  const [newFinalDraft, setNewFinalDraft] = useState("");
  const [newDeadline, setNewDeadline] = useState("");
  const [newInstructions, setNewInstructions] = useState("");
  const [newAllocatedTo, setNewAllocatedTo] = useState<DataverseUser | null>(null);
  const [addBusy, setAddBusy] = useState(false);

  // Auto-generate question number when section changes (unless user has manually edited it)
  useEffect(() => {
    if (newQNumManual) return;
    const sectionItems = sections.get(newSection) ?? [];
    const suggested = nextQuestionNumber(newSection, sectionItems);
    setNewQNum(suggested);
  }, [newSection, newQNumManual, sections]);

  function resetAddForm() {
    setNewSection("");
    setNewQNum("");
    setNewQNumManual(false);
    setNewDetail("");
    setNewDepartment("");
    setNewWeighting("");
    setNewFirstDraft("");
    setNewFinalDraft("");
    setNewDeadline("");
    setNewInstructions("");
    setNewAllocatedTo(null);
  }

  async function handleAddRow() {
    if (!newQNum || !newDetail) return;
    setAddBusy(true);
    try {
      await dataverseClient.createTorItem({
        cr5ab_bidworkspaceid: { id: workspaceId, cr5ab_title: "" },
        cr5ab_section: newSection || "Uncategorised",
        cr5ab_questionnumber: newQNum,
        cr5ab_questiondetail: newDetail,
        cr5ab_department: newDepartment || undefined,
        cr5ab_scoreweightingthreshold: newWeighting || undefined,
        cr5ab_firstdraftdeadline: newFirstDraft || undefined,
        cr5ab_finaldraftdeadline: newFinalDraft || undefined,
        cr5ab_actualdeadline: newDeadline || undefined,
        cr5ab_specialinstructions: newInstructions || undefined,
        cr5ab_answeredstatus: TorAnsweredStatus.No,
        ...(newAllocatedTo ? { cr5ab_allocatedto: newAllocatedTo } : {}),
      });
      refresh();
      setShowAddRow(false);
      resetAddForm();
    } finally { setAddBusy(false); }
  }

  // ── render ─────────────────────────────────────────────────────────────
  if (isLoading) return <LoadingState label="Loading TOR..." />;
  if (error) return <ErrorState message={error} onRetry={refresh} />;

  return (
    <div>
      {/* Progress */}
      <Card className={styles.progressCard}>
        <div className={styles.progressRow}>
          <Text weight="semibold" size={300}>TOR Progress — {answered} of {total} answered</Text>
          <Text size={300} weight="semibold">{progress}%</Text>
        </div>
        <ProgressBar value={progress / 100} color={progress >= 75 ? "success" : progress >= 40 ? "warning" : "error"} />
        <div style={{ display: "flex", gap: tokens.spacingHorizontalL, marginTop: tokens.spacingVerticalS }}>
          {([TorAnsweredStatus.Yes, TorAnsweredStatus.Partial, TorAnsweredStatus.No] as TorAnsweredStatus[]).map((s) => {
            const count = items.filter((t) => t.cr5ab_answeredstatus === s).length;
            return (
              <div key={s} style={{ display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXS }}>
                <Badge appearance="filled" color={TorAnsweredStatusColor[s]} size="small">{TorAnsweredStatusLabel[s]}</Badge>
                <Text size={200}>{count}</Text>
              </div>
            );
          })}
        </div>
      </Card>

      {/* Toolbar — Add Row only */}
      <div className={styles.torToolbar}>
        <Button
          appearance="primary"
          icon={<AddRegular />}
          onClick={() => setShowAddRow(true)}
        >
          Add Row
        </Button>
      </div>

      {/* Empty state */}
      {items.length === 0 && (
        <EmptyState icon={<TableRegular />} title="No TOR rows yet" description="Add the first row using the button above." />
      )}

      {/* Sections */}
      {Array.from(sections.entries()).map(([section, rows]) => {
        const secAnswered = rows.filter((r) => r.cr5ab_answeredstatus === TorAnsweredStatus.Yes).length;
        const isCollapsed = !!collapsedSections[section];

        return (
          <div key={section} className={styles.torSection}>
            {/* Section header — clickable to collapse */}
            <button
              className={styles.torSectionHeader}
              onClick={() => toggleSection(section)}
              type="button"
              aria-expanded={!isCollapsed}
            >
              {isCollapsed
                ? <ChevronRightRegular style={{ fontSize: "16px", flexShrink: 0, color: tokens.colorNeutralForeground3 }} />
                : <ChevronDownRegular  style={{ fontSize: "16px", flexShrink: 0, color: tokens.colorNeutralForeground3 }} />
              }
              <Text weight="semibold" size={400}>{section}</Text>
              <div className={styles.torSectionProgress}>
                <Badge
                  appearance="filled"
                  color={secAnswered === rows.length ? "success" : secAnswered > 0 ? "warning" : "danger"}
                  size="small"
                >
                  {secAnswered}/{rows.length}
                </Badge>
                <ProgressBar
                  value={rows.length > 0 ? secAnswered / rows.length : 0}
                  style={{ width: "80px" }}
                  color={secAnswered === rows.length ? "success" : "warning"}
                />
              </div>
            </button>

            {!isCollapsed && (
              <>
                {/* Column headers */}
                <div className={styles.torHeaderRow}>
                  <Text size={100} weight="semibold" style={{ color: tokens.colorNeutralForeground3 }}>Q#</Text>
                  <Text size={100} weight="semibold" style={{ color: tokens.colorNeutralForeground3 }}>Question Detail</Text>
                  <Text size={100} weight="semibold" style={{ color: tokens.colorNeutralForeground3 }}>Owner</Text>
                  <Text size={100} weight="semibold" style={{ color: tokens.colorNeutralForeground3 }}>Deadlines</Text>
                  <Text size={100} weight="semibold" style={{ color: tokens.colorNeutralForeground3 }}>Weighting</Text>
                  <Text size={100} weight="semibold" style={{ color: tokens.colorNeutralForeground3 }}>Status</Text>
                  <Text size={100} weight="semibold" style={{ color: tokens.colorNeutralForeground3 }}></Text>
                </div>

                {rows.map((item) => {
                  const isEditing = editingId === item.id;
                  const showComment = showCommentFor === item.id;
                  const rowClass = `${styles.torRow} ${
                    item.cr5ab_answeredstatus === TorAnsweredStatus.Yes ? styles.torRowAnswered
                    : item.cr5ab_answeredstatus === TorAnsweredStatus.Partial ? styles.torRowPartial
                    : styles.torRowNo
                  }`;
                  const deadline = item.cr5ab_actualdeadline;
                  const days = deadline ? daysUntil(deadline) : null;

                  const statusPillClass = `${styles.statusPill} ${
                    item.cr5ab_answeredstatus === TorAnsweredStatus.Yes ? styles.statusPillAnswered
                    : item.cr5ab_answeredstatus === TorAnsweredStatus.Partial ? styles.statusPillPartial
                    : styles.statusPillNo
                  }`;

                  return (
                    <div key={item.id} className={rowClass}>
                      {/* Q# */}
                      <Text size={200} weight="semibold">{item.cr5ab_questionnumber}</Text>

                      {/* Question Detail + comment area */}
                      <div>
                        <Text
                          size={200}
                          style={{
                            display: "-webkit-box",
                            WebkitLineClamp: 2,
                            WebkitBoxOrient: "vertical",
                            overflow: "hidden",
                          }}
                          title={item.cr5ab_questiondetail}
                        >
                          {item.cr5ab_questiondetail}
                        </Text>
                        {item.cr5ab_specialinstructions && (
                          <Text size={100} style={{ display: "block", color: tokens.colorNeutralForeground3, fontStyle: "italic" }}>
                            {item.cr5ab_specialinstructions}
                          </Text>
                        )}
                        {/* Comments — only shown when editing or explicitly expanded */}
                        {isEditing && (
                          <Textarea
                            placeholder="Add notes or comments..."
                            value={editComments}
                            onChange={(_, d) => setEditComments(d.value)}
                            rows={2}
                            style={{ marginTop: "4px", width: "100%" }}
                          />
                        )}
                        {!isEditing && item.cr5ab_comments && !showComment && (
                          <button
                            className={styles.commentToggle}
                            onClick={() => setShowCommentFor(item.id)}
                          >
                            Show comment
                          </button>
                        )}
                        {!isEditing && showComment && item.cr5ab_comments && (
                          <div style={{ marginTop: "4px" }}>
                            <Text size={100} style={{ color: tokens.colorNeutralForeground2 }}>{item.cr5ab_comments}</Text>
                            <button
                              className={styles.commentToggle}
                              onClick={() => setShowCommentFor(null)}
                              style={{ marginLeft: "4px" }}
                            >
                              hide
                            </button>
                          </div>
                        )}
                      </div>

                      {/* Allocated To — first name only if long */}
                      <Text
                        size={200}
                        style={{
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                          whiteSpace: "nowrap",
                          display: "block",
                        }}
                        title={item.cr5ab_allocatedto?.fullName}
                      >
                        {item.cr5ab_allocatedto?.fullName ?? "—"}
                      </Text>

                      {/* Deadlines — stacked: 1st Draft / Final / Submission */}
                      <div style={{ display: "flex", flexDirection: "column", gap: "2px" }}>
                        {item.cr5ab_firstdraftdeadline && (
                          <div style={{ display: "flex", alignItems: "center", gap: "3px" }}>
                            <Text size={100} style={{ color: tokens.colorNeutralForeground3, minWidth: "10px" }}>1</Text>
                            <Text size={100} style={{ color: tokens.colorNeutralForeground2 }}>{fmtDate(item.cr5ab_firstdraftdeadline)}</Text>
                          </div>
                        )}
                        {item.cr5ab_finaldraftdeadline && (
                          <div style={{ display: "flex", alignItems: "center", gap: "3px" }}>
                            <Text size={100} style={{ color: tokens.colorNeutralForeground3, minWidth: "10px" }}>F</Text>
                            <Text size={100} style={{ color: tokens.colorNeutralForeground2 }}>{fmtDate(item.cr5ab_finaldraftdeadline)}</Text>
                          </div>
                        )}
                        {deadline && (
                          <div style={{ display: "flex", alignItems: "center", gap: "3px" }}>
                            <Text size={100} style={{ color: tokens.colorNeutralForeground3, minWidth: "10px" }}>S</Text>
                            <Text
                              size={100}
                              weight="semibold"
                              style={{ color: days !== null && days < 0 ? tokens.colorPaletteRedForeground3 : days !== null && days <= 3 ? tokens.colorPaletteMarigoldForeground2 : tokens.colorNeutralForeground1 }}
                            >
                              {fmtDate(deadline)}
                            </Text>
                          </div>
                        )}
                        {!item.cr5ab_firstdraftdeadline && !item.cr5ab_finaldraftdeadline && !deadline && (
                          <Text size={100} style={{ color: tokens.colorNeutralForeground4 }}>—</Text>
                        )}
                        {days !== null && (
                          <Text size={100} style={{ color: days < 0 ? tokens.colorPaletteRedForeground3 : days <= 3 ? tokens.colorPaletteMarigoldForeground2 : tokens.colorNeutralForeground3, fontStyle: "italic" }}>
                            {days < 0 ? `${Math.abs(days)}d over` : days === 0 ? "Today" : `${days}d`}
                          </Text>
                        )}
                      </div>

                      {/* Weighting */}
                      <Text size={200} style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "block" }}>
                        {item.cr5ab_scoreweightingthreshold ?? "—"}
                      </Text>

                      {/* Status — compact pill / select */}
                      <div>
                        {isEditing ? (
                          <Select
                            value={editStatus}
                            onChange={(_, d) => setEditStatus(d.value as TorAnsweredStatus)}
                            size="small"
                            style={{ width: "100%", minWidth: 0 }}
                          >
                            <option value={TorAnsweredStatus.Yes}>Done</option>
                            <option value={TorAnsweredStatus.Partial}>Partial</option>
                            <option value={TorAnsweredStatus.No}>Not Started</option>
                          </Select>
                        ) : (
                          <span className={statusPillClass} style={{ fontSize: tokens.fontSizeBase100, whiteSpace: "nowrap" }}>
                            {item.cr5ab_answeredstatus === TorAnsweredStatus.Yes ? "Done"
                              : item.cr5ab_answeredstatus === TorAnsweredStatus.Partial ? "Partial"
                              : "Not Started"}
                          </span>
                        )}
                      </div>

                      {/* Actions */}
                      <div style={{ display: "flex", gap: "4px", justifyContent: "flex-end" }}>
                        {isEditing ? (
                          <>
                            <Tooltip content="Save" relationship="label">
                              <Button size="small" appearance="primary" icon={saving ? <Spinner size="tiny" /> : <CheckmarkRegular />} disabled={saving} onClick={() => saveEdit(item)} />
                            </Tooltip>
                            <Tooltip content="Cancel" relationship="label">
                              <Button size="small" appearance="subtle" icon={<DismissRegular />} onClick={() => { setEditingId(null); setShowCommentFor(null); }} />
                            </Tooltip>
                          </>
                        ) : (
                          <>
                            <Tooltip content="Edit" relationship="label">
                              <Button size="small" appearance="subtle" icon={<EditRegular />} onClick={() => startEdit(item)} />
                            </Tooltip>
                            <Tooltip content="Delete" relationship="label">
                              <Button size="small" appearance="subtle" icon={<DeleteRegular />} onClick={() => handleDelete(item.id)} />
                            </Tooltip>
                          </>
                        )}
                      </div>
                    </div>
                  );
                })}
              </>
            )}
          </div>
        );
      })}

      {/* ── Add Row Dialog ─────────────────────────────────────────────── */}
      <Dialog open={showAddRow} onOpenChange={(_, d) => { setShowAddRow(d.open); if (!d.open) resetAddForm(); }}>
        <DialogSurface style={{ maxWidth: "560px" }}>
          <DialogBody>
            <DialogTitle>Add TOR Row</DialogTitle>
            <DialogContent>
              <div style={{ display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM, marginTop: tokens.spacingVerticalM }}>

                {/* Section picker */}
                <Field label="Section">
                  <SectionPicker
                    value={newSection}
                    onChange={(val) => { setNewSection(val); setNewQNumManual(false); }}
                    existingSections={sectionNames}
                  />
                </Field>

                {/* Question Number — auto-generated + manually overridable */}
                <Field label="Question Number" required hint="Auto-generated from section. Override if needed.">
                  <Input
                    placeholder="e.g. 1.1"
                    value={newQNum}
                    onChange={(_, d) => { setNewQNum(d.value); setNewQNumManual(true); }}
                  />
                </Field>

                {/* Question Detail */}
                <Field label="Question Detail" required>
                  <Textarea
                    rows={3}
                    placeholder="Describe the question or requirement..."
                    value={newDetail}
                    onChange={(_, d) => setNewDetail(d.value)}
                  />
                </Field>

                {/* Department */}
                <Field label="Department">
                  <Input
                    placeholder="e.g. Technical, Commercial, Bid Management"
                    value={newDepartment}
                    onChange={(_, d) => setNewDepartment(d.value)}
                  />
                </Field>

                {/* Deadlines — 3-column grid */}
                <div>
                  <Text size={200} weight="semibold" style={{ display: "block", marginBottom: tokens.spacingVerticalXS, color: tokens.colorNeutralForeground2 }}>Deadlines</Text>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: tokens.spacingHorizontalS }}>
                    <Field label="1st Draft (internal)">
                      <Input type="date" value={newFirstDraft} onChange={(_, d) => setNewFirstDraft(d.value)} />
                    </Field>
                    <Field label="Final Draft (internal)">
                      <Input type="date" value={newFinalDraft} onChange={(_, d) => setNewFinalDraft(d.value)} />
                    </Field>
                    <Field label="Submission">
                      <Input type="date" value={newDeadline} onChange={(_, d) => setNewDeadline(d.value)} />
                    </Field>
                  </div>
                </div>

                {/* Weighting */}
                <Field label="Weighting / Threshold">
                  <WeightingField value={newWeighting} onChange={setNewWeighting} />
                </Field>

                {/* Special Instructions */}
                <Field label="Special Instructions">
                  <Textarea
                    rows={2}
                    placeholder="Any specific guidance..."
                    value={newInstructions}
                    onChange={(_, d) => setNewInstructions(d.value)}
                  />
                </Field>

                {/* People picker */}
                <Field label="Allocated To">
                  <PeoplePicker value={newAllocatedTo} onChange={setNewAllocatedTo} />
                </Field>

              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="subtle" onClick={() => { setShowAddRow(false); resetAddForm(); }}>
                Cancel
              </Button>
              <Button
                appearance="primary"
                disabled={addBusy || !newQNum || !newDetail}
                onClick={handleAddRow}
              >
                {addBusy ? <Spinner size="tiny" /> : "Add Row"}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Tab: Team
// ---------------------------------------------------------------------------

function TeamTab({ members: initialMembers, workspaceId }: { members: BidRoleAssignment[]; workspaceId: string }) {
  const styles = useStyles();
  const [members, setMembers] = useState(initialMembers);
  const [showDialog, setShowDialog] = useState(false);
  const [selectedUser, setSelectedUser] = useState<DataverseUser | null>(null);
  const [selectedRole, setSelectedRole] = useState<TeamRole>(TeamRole.SME);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  async function handleAdd() {
    if (!selectedUser) return;
    setBusy(true);
    setError(null);
    try {
      const newMember = await dataverseClient.createRoleAssignment({
        cr5ab_bidworkspaceid: { id: workspaceId, cr5ab_title: "" },
        cr5ab_userid: selectedUser,
        cr5ab_role: selectedRole,
        cr5ab_isactive: true,
        cr5ab_assigneddate: new Date().toISOString(),
      });
      setMembers((prev) => [...prev, newMember]);
      setShowDialog(false);
      setSelectedUser(null);
      setSelectedRole(TeamRole.SME);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to add member");
    } finally { setBusy(false); }
  }

  return (
    <Card>
      <CardHeader
        header={<Text weight="semibold">Team Members</Text>}
        action={
          <Button appearance="primary" size="small" icon={<AddRegular />} onClick={() => setShowDialog(true)}>
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
              <TableHeaderCell>Assigned</TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {members.map((m) => (
              <TableRow key={m.id}>
                <TableCell><TableCellLayout><Text weight="semibold" size={300}>{m.cr5ab_userid.fullName}</Text></TableCellLayout></TableCell>
                <TableCell><TableCellLayout><Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>{m.cr5ab_userid.email}</Text></TableCellLayout></TableCell>
                <TableCell><TableCellLayout><Badge appearance="outline" color="informative">{TeamRoleLabel[m.cr5ab_role as TeamRole]}</Badge></TableCellLayout></TableCell>
                <TableCell><TableCellLayout><Badge appearance="filled" color={m.cr5ab_isactive ? "success" : "subtle"} size="small">{m.cr5ab_isactive ? "Active" : "Inactive"}</Badge></TableCellLayout></TableCell>
                <TableCell><TableCellLayout><Text size={200}>{fmtDate(m.cr5ab_assigneddate)}</Text></TableCellLayout></TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      )}

      <Dialog open={showDialog} onOpenChange={(_, d) => setShowDialog(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Add Team Member</DialogTitle>
            <DialogContent>
              <div style={{ display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM, marginTop: tokens.spacingVerticalM }}>
                <Field label="User" required>
                  <PeoplePicker
                    value={selectedUser}
                    onChange={setSelectedUser}
                    placeholder="Search for team member..."
                  />
                </Field>
                <Field label="Role" required>
                  <Select value={String(selectedRole)} onChange={(_, d) => setSelectedRole(Number(d.value) as TeamRole)}>
                    {Object.entries(TeamRoleLabel).map(([code, label]) => (
                      <option key={code} value={code}>{label}</option>
                    ))}
                  </Select>
                </Field>
                {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="subtle" onClick={() => setShowDialog(false)}>Cancel</Button>
              <Button appearance="primary" disabled={busy || !selectedUser} onClick={handleAdd}>
                {busy ? <Spinner size="tiny" /> : "Add"}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Tab: Approvals
// ---------------------------------------------------------------------------

function ApprovalsTab({ approvals: initialApprovals, workspaceId }: { approvals: BidApproval[]; workspaceId: string }) {
  const styles = useStyles();
  const [approvals, setApprovals] = useState(initialApprovals);
  const [showDialog, setShowDialog] = useState(false);
  const [title, setTitle] = useState("");
  const [approverEmail, setApproverEmail] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [respondingId, setRespondingId] = useState<string | null>(null);
  const [respondComment, setRespondComment] = useState("");
  const [respondBusy, setRespondBusy] = useState(false);

  const nextStage = approvals.length + 1;

  async function handleRequest() {
    if (!title || !approverEmail) return;
    setBusy(true);
    setError(null);
    try {
      const fakeUser: DataverseUser = { id: `usr-${Date.now()}`, fullName: approverEmail.split("@")[0].replace(".", " "), email: approverEmail };
      const newAp = await dataverseClient.createApproval({
        cr5ab_bidworkspaceid: { id: workspaceId, cr5ab_title: "" },
        cr5ab_title: title, cr5ab_approverstage: nextStage,
        cr5ab_approverid: fakeUser, cr5ab_status: ApprovalStatus.Pending,
        cr5ab_requesteddate: new Date().toISOString(),
      });
      setApprovals((prev) => [...prev, newAp]);
      setShowDialog(false); setTitle(""); setApproverEmail("");
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to request approval");
    } finally { setBusy(false); }
  }

  async function handleRespond(ap: BidApproval, status: ApprovalStatus.Approved | ApprovalStatus.Rejected) {
    setRespondBusy(true);
    try {
      await dataverseClient.updateApproval(ap.id, { cr5ab_status: status, cr5ab_respondeddate: new Date().toISOString(), cr5ab_comments: respondComment || undefined });
      setApprovals((prev) => prev.map((a) => a.id === ap.id ? { ...a, cr5ab_status: status, cr5ab_respondeddate: new Date().toISOString(), cr5ab_comments: respondComment || a.cr5ab_comments } : a));
      setRespondingId(null); setRespondComment("");
    } finally { setRespondBusy(false); }
  }

  return (
    <Card>
      <CardHeader
        header={<Text weight="semibold">Approval Chain</Text>}
        action={
          <Button appearance="primary" size="small" icon={<AddRegular />} onClick={() => setShowDialog(true)}>
            Request Approval
          </Button>
        }
      />
      {approvals.length === 0 ? (
        <EmptyState title="No approvals" description="No approvals have been requested yet." />
      ) : (
        approvals.map((ap) => {
          const isResponding = respondingId === ap.id;
          const stageClass = `${styles.stageNumber} ${ap.cr5ab_status === ApprovalStatus.Approved ? styles.stageNumberApproved : ap.cr5ab_status === ApprovalStatus.Pending ? styles.stageNumberPending : ""}`;
          return (
            <div key={ap.id} className={styles.approvalRow}>
              <div className={stageClass}>{ap.cr5ab_approverstage}</div>
              <div style={{ flexGrow: 1 }}>
                <Text weight="semibold" size={300}>{ap.cr5ab_title}</Text>
                <Text size={200} style={{ display: "block", color: tokens.colorNeutralForeground3 }}>
                  {ap.cr5ab_approverid.fullName} — {ap.cr5ab_approverid.email}
                </Text>
                <Text size={100} style={{ display: "block", color: tokens.colorNeutralForeground4 }}>
                  Requested: {fmtDate(ap.cr5ab_requesteddate)}
                  {ap.cr5ab_respondeddate && ` · Responded: ${fmtDate(ap.cr5ab_respondeddate)}`}
                </Text>
                {ap.cr5ab_comments && (
                  <Text size={200} style={{ color: tokens.colorNeutralForeground2, fontStyle: "italic", marginTop: "2px" }}>
                    "{ap.cr5ab_comments}"
                  </Text>
                )}
                {isResponding && (
                  <div style={{ marginTop: tokens.spacingVerticalS, display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS }}>
                    <Textarea placeholder="Optional comments..." value={respondComment} onChange={(_, d) => setRespondComment(d.value)} rows={2} />
                    <div style={{ display: "flex", gap: tokens.spacingHorizontalS }}>
                      <Button size="small" appearance="primary" icon={<CheckmarkRegular />} disabled={respondBusy} onClick={() => handleRespond(ap, ApprovalStatus.Approved)}>Approve</Button>
                      <Button size="small" appearance="secondary" icon={<DismissRegular />} disabled={respondBusy} onClick={() => handleRespond(ap, ApprovalStatus.Rejected)}>Reject</Button>
                      <Button size="small" appearance="subtle" onClick={() => setRespondingId(null)}>Cancel</Button>
                    </div>
                  </div>
                )}
                {ap.cr5ab_status === ApprovalStatus.Pending && !isResponding && (
                  <Button size="small" appearance="subtle" style={{ marginTop: tokens.spacingVerticalXS }} onClick={() => setRespondingId(ap.id)}>
                    Respond
                  </Button>
                )}
              </div>
              <div style={{ minWidth: "90px", textAlign: "right" }}>
                <ApprovalBadge status={ap.cr5ab_status} />
              </div>
            </div>
          );
        })
      )}

      <Dialog open={showDialog} onOpenChange={(_, d) => setShowDialog(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Request Approval</DialogTitle>
            <DialogContent>
              <div style={{ display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM, marginTop: tokens.spacingVerticalM }}>
                <Field label="Approval title" required>
                  <Input placeholder="e.g. Technical Review" value={title} onChange={(_, d) => setTitle(d.value)} />
                </Field>
                <Field label="Approver email" required>
                  <Input type="email" placeholder="approver@ricoh.co.uk" value={approverEmail} onChange={(_, d) => setApproverEmail(d.value)} />
                </Field>
                <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>This will be stage {nextStage} in the approval chain.</Text>
                {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="subtle" onClick={() => setShowDialog(false)}>Cancel</Button>
              <Button appearance="primary" disabled={busy || !title || !approverEmail} onClick={handleRequest}>
                {busy ? <Spinner size="tiny" /> : "Request"}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Tab: Documents
// ---------------------------------------------------------------------------

const DOC_CATEGORIES: DocumentCategory[] = ["originals", "working", "submission"];

function DocumentsTab({ documents: initialDocs, workspaceId, sharepointBase }: {
  documents: BidDocumentRecord[];
  workspaceId: string;
  sharepointBase?: string;
}) {
  const styles = useStyles();
  const [docs, setDocs] = useState(initialDocs);
  const [activeCategory, setActiveCategory] = useState<DocumentCategory>("originals");

  // ── Link Document modal ────────────────────────────────────────────────
  const [showLink, setShowLink] = useState(false);
  const [linkTitle, setLinkTitle] = useState("");
  const [linkUrl, setLinkUrl] = useState(sharepointBase ?? "");
  const [linkType, setLinkType] = useState("");
  const [linkCategory, setLinkCategory] = useState<DocumentCategory>("working");
  const [linkBusy, setLinkBusy] = useState(false);
  const [linkError, setLinkError] = useState<string | null>(null);

  // ── Upload Document modal ──────────────────────────────────────────────
  const [showUpload, setShowUpload] = useState(false);
  const [uploadFile, setUploadFile] = useState<File | null>(null);
  const [uploadTitle, setUploadTitle] = useState("");
  const [uploadType, setUploadType] = useState("");
  const [uploadCategory, setUploadCategory] = useState<DocumentCategory>("working");
  const [uploadProgress, setUploadProgress] = useState<"idle" | "uploading" | "saving" | "done" | "error">("idle");
  const [uploadError, setUploadError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const SYSTEM_USER = { id: "00000000-0000-0000-0000-000000000001", fullName: "Dev User", email: "dev@ricoh.co.uk" };

  const filteredDocs = docs.filter((d) => d.cr5ab_category === activeCategory);
  const isOriginalsTab = activeCategory === "originals";

  // ── Link handler ───────────────────────────────────────────────────────
  async function handleLink() {
    if (!linkTitle || !linkUrl) return;
    setLinkBusy(true);
    setLinkError(null);
    try {
      const newDoc: BidDocumentRecord = {
        id: `doc-${Date.now()}`,
        cr5ab_title: linkTitle,
        cr5ab_documenttype: linkType || "Document",
        cr5ab_category: linkCategory,
        cr5ab_sharepointurl: linkUrl,
        cr5ab_version: "1.0",
        cr5ab_uploadedby: SYSTEM_USER,
        cr5ab_bidworkspaceid: { id: workspaceId, cr5ab_title: "" },
        createdOn: new Date().toISOString(),
        modifiedOn: new Date().toISOString(),
        createdBy: SYSTEM_USER,
        modifiedBy: SYSTEM_USER,
      };
      setDocs((prev) => [...prev, newDoc]);
      setActiveCategory(linkCategory);
      setShowLink(false);
      setLinkTitle(""); setLinkUrl(sharepointBase ?? ""); setLinkType("");
    } catch (e) {
      setLinkError(e instanceof Error ? e.message : "Failed to link document");
    } finally { setLinkBusy(false); }
  }

  // ── Upload handler ─────────────────────────────────────────────────────
  async function handleUpload() {
    if (!uploadFile || !uploadTitle) return;
    setUploadProgress("uploading");
    setUploadError(null);
    try {
      // Read file as base64
      const base64 = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve((reader.result as string).split(",")[1] ?? "");
        reader.onerror = reject;
        reader.readAsDataURL(uploadFile);
      });

      setUploadProgress("saving");
      const { sharepointUrl } = await dataverseClient.triggerUploadDocumentFlow({
        workspaceId,
        fileName: uploadFile.name,
        fileBase64: base64,
        documentType: uploadType || "Document",
        category: uploadCategory,
      });

      const newDoc: BidDocumentRecord = {
        id: `doc-${Date.now()}`,
        cr5ab_title: uploadTitle,
        cr5ab_filename: uploadFile.name,
        cr5ab_documenttype: uploadType || "Document",
        cr5ab_category: uploadCategory,
        cr5ab_sharepointurl: sharepointUrl,
        cr5ab_filesize: uploadFile.size,
        cr5ab_version: "1.0",
        cr5ab_uploadedby: SYSTEM_USER,
        cr5ab_bidworkspaceid: { id: workspaceId, cr5ab_title: "" },
        createdOn: new Date().toISOString(),
        modifiedOn: new Date().toISOString(),
        createdBy: SYSTEM_USER,
        modifiedBy: SYSTEM_USER,
      };
      setDocs((prev) => [...prev, newDoc]);
      setActiveCategory(uploadCategory);
      setUploadProgress("done");

      // Auto-close after brief "done" display
      setTimeout(() => {
        setShowUpload(false);
        setUploadProgress("idle");
        setUploadFile(null);
        setUploadTitle("");
        setUploadType("");
      }, 1200);
    } catch (e) {
      setUploadError(e instanceof Error ? e.message : "Upload failed");
      setUploadProgress("error");
    }
  }

  function formatBytes(bytes?: number) {
    if (!bytes) return "";
    if (bytes >= 1_000_000) return ` · ${(bytes / 1_000_000).toFixed(1)} MB`;
    if (bytes >= 1_000) return ` · ${(bytes / 1_000).toFixed(0)} KB`;
    return ` · ${bytes} B`;
  }

  return (
    <Card>
      <CardHeader
        header={<Text weight="semibold">Documents</Text>}
        action={
          <div style={{ display: "flex", gap: tokens.spacingHorizontalS }}>
            <Button appearance="secondary" size="small" icon={<LinkRegular />} onClick={() => setShowLink(true)}>
              Link Document
            </Button>
            <Button appearance="primary" size="small" icon={<ArrowUploadRegular />} onClick={() => setShowUpload(true)}>
              Upload Document
            </Button>
          </div>
        }
      />

      {/* Category tabs */}
      <div className={styles.docTabRow}>
        {DOC_CATEGORIES.map((cat) => {
          const count = docs.filter((d) => d.cr5ab_category === cat).length;
          return (
            <button
              key={cat}
              className={`${styles.docTab} ${activeCategory === cat ? styles.docTabActive : ""}`}
              onClick={() => setActiveCategory(cat)}
            >
              {cat === "originals" && <LockClosedRegular style={{ fontSize: "12px" }} />}
              {DocumentCategoryLabel[cat]} ({count})
            </button>
          );
        })}
      </div>

      {/* Originals read-only notice */}
      {isOriginalsTab && (
        <MessageBar intent="info" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>
            <strong>Read-only.</strong> Originals are source documents from the customer. Link here for reference — these cannot be edited.
          </MessageBarBody>
        </MessageBar>
      )}

      {/* Document list */}
      {filteredDocs.length === 0 ? (
        <EmptyState
          icon={<DocumentRegular />}
          title={`No ${DocumentCategoryLabel[activeCategory].toLowerCase()}`}
          description={isOriginalsTab ? "Link original customer documents here." : "Upload or link SharePoint documents."}
        />
      ) : (
        filteredDocs.map((doc) => (
          <div key={doc.id} className={styles.docRow}>
            <DocumentRegular style={{ fontSize: "20px", color: tokens.colorNeutralForeground3, flexShrink: 0 }} />
            <div style={{ flexGrow: 1 }}>
              <Text weight="semibold" size={300}>{doc.cr5ab_title}</Text>
              <Text size={200} style={{ display: "block", color: tokens.colorNeutralForeground3 }}>
                {doc.cr5ab_documenttype} · v{doc.cr5ab_version}{formatBytes(doc.cr5ab_filesize)} · {fmtDate(doc.createdOn)} · {doc.cr5ab_uploadedby.fullName}
              </Text>
              {doc.cr5ab_filename && (
                <Text size={100} style={{ color: tokens.colorNeutralForeground4 }}>{doc.cr5ab_filename}</Text>
              )}
            </div>
            {isOriginalsTab && (
              <Badge appearance="outline" size="small" icon={<LockClosedRegular />} color="subtle">
                Read-only
              </Badge>
            )}
            <Button
              appearance="subtle" size="small" icon={<LinkRegular />}
              as="a" href={doc.cr5ab_sharepointurl} target="_blank" rel="noopener noreferrer"
            >
              Open
            </Button>
          </div>
        ))
      )}

      {/* ── Link Document Dialog ───────────────────────────────────────── */}
      <Dialog open={showLink} onOpenChange={(_, d) => setShowLink(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Link Document</DialogTitle>
            <DialogContent>
              <div style={{ display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM, marginTop: tokens.spacingVerticalM }}>
                <Field label="Document title" required>
                  <Input placeholder="e.g. ITT Document v1" value={linkTitle} onChange={(_, d) => setLinkTitle(d.value)} />
                </Field>
                <Field label="SharePoint URL" required>
                  <Input placeholder="https://ricoh.sharepoint.com/..." value={linkUrl} onChange={(_, d) => setLinkUrl(d.value)} />
                </Field>
                <Field label="Document type">
                  <Input placeholder="e.g. Tender Document, Response, Summary" value={linkType} onChange={(_, d) => setLinkType(d.value)} />
                </Field>
                <Field label="Category" required>
                  <Select value={linkCategory} onChange={(_, d) => setLinkCategory(d.value as DocumentCategory)}>
                    {DOC_CATEGORIES.map((cat) => (
                      <option key={cat} value={cat}>{DocumentCategoryLabel[cat]}</option>
                    ))}
                  </Select>
                </Field>
                <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                  The document will open in a new tab when clicked from the list.
                </Text>
                {linkError && <MessageBar intent="error"><MessageBarBody>{linkError}</MessageBarBody></MessageBar>}
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="subtle" onClick={() => setShowLink(false)}>Cancel</Button>
              <Button appearance="primary" disabled={linkBusy || !linkTitle || !linkUrl} onClick={handleLink}>
                {linkBusy ? <Spinner size="tiny" /> : "Link Document"}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* ── Upload Document Dialog ─────────────────────────────────────── */}
      <Dialog open={showUpload} onOpenChange={(_, d) => { if (uploadProgress !== "uploading" && uploadProgress !== "saving") setShowUpload(d.open); }}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Upload Document</DialogTitle>
            <DialogContent>
              <div style={{ display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM, marginTop: tokens.spacingVerticalM }}>
                {/* Hidden file input */}
                <input
                  ref={fileInputRef}
                  type="file"
                  style={{ display: "none" }}
                  onChange={(e) => {
                    const f = e.target.files?.[0] ?? null;
                    setUploadFile(f);
                    if (f && !uploadTitle) setUploadTitle(f.name.replace(/\.[^.]+$/, ""));
                  }}
                />

                {/* File picker */}
                <Field label="File" required>
                  <div style={{ display: "flex", gap: tokens.spacingHorizontalS, alignItems: "center" }}>
                    <Button
                      appearance="secondary"
                      size="small"
                      icon={<ArrowUploadRegular />}
                      onClick={() => fileInputRef.current?.click()}
                    >
                      Choose file
                    </Button>
                    {uploadFile ? (
                      <Text size={200} style={{ color: tokens.colorBrandForeground1 }}>
                        {uploadFile.name}{formatBytes(uploadFile.size)}
                      </Text>
                    ) : (
                      <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>No file chosen</Text>
                    )}
                  </div>
                </Field>

                <Field label="Document title" required>
                  <Input
                    placeholder="e.g. Technical Response v1"
                    value={uploadTitle}
                    onChange={(_, d) => setUploadTitle(d.value)}
                  />
                </Field>

                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: tokens.spacingHorizontalM }}>
                  <Field label="Document type">
                    <Input placeholder="e.g. Response, Summary" value={uploadType} onChange={(_, d) => setUploadType(d.value)} />
                  </Field>
                  <Field label="Category" required>
                    <Select value={uploadCategory} onChange={(_, d) => setUploadCategory(d.value as DocumentCategory)}>
                      {DOC_CATEGORIES.map((cat) => (
                        <option key={cat} value={cat}>{DocumentCategoryLabel[cat]}</option>
                      ))}
                    </Select>
                  </Field>
                </div>

                {/* Progress states */}
                {uploadProgress === "uploading" && (
                  <div className={styles.uploadProgress}>
                    <div style={{ display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS }}>
                      <Spinner size="tiny" />
                      <Text size={200}>Reading file...</Text>
                    </div>
                    <ProgressBar />
                  </div>
                )}
                {uploadProgress === "saving" && (
                  <div className={styles.uploadProgress}>
                    <div style={{ display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS }}>
                      <Spinner size="tiny" />
                      <Text size={200}>Uploading to SharePoint via Power Automate...</Text>
                    </div>
                    <ProgressBar />
                  </div>
                )}
                {uploadProgress === "done" && (
                  <MessageBar intent="success">
                    <MessageBarBody>Upload complete. Document saved to SharePoint.</MessageBarBody>
                  </MessageBar>
                )}
                {uploadProgress === "error" && uploadError && (
                  <MessageBar intent="error">
                    <MessageBarBody>{uploadError}</MessageBarBody>
                  </MessageBar>
                )}
              </div>
            </DialogContent>
            <DialogActions>
              <Button
                appearance="subtle"
                disabled={uploadProgress === "uploading" || uploadProgress === "saving"}
                onClick={() => { setShowUpload(false); setUploadProgress("idle"); setUploadFile(null); setUploadTitle(""); setUploadType(""); }}
              >
                Cancel
              </Button>
              <Button
                appearance="primary"
                icon={<ArrowUploadRegular />}
                disabled={!uploadFile || !uploadTitle || uploadProgress === "uploading" || uploadProgress === "saving" || uploadProgress === "done"}
                onClick={handleUpload}
              >
                {uploadProgress === "uploading" || uploadProgress === "saving"
                  ? <Spinner size="tiny" />
                  : "Upload"
                }
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Tab: Clarifications
// ---------------------------------------------------------------------------

function ClarificationsTab({ workspaceId }: { workspaceId: string }) {
  const styles = useStyles();

  const { data: cqItems, isLoading, error, refresh } = useDataverse(
    () => dataverseClient.getClarifications(workspaceId),
    [workspaceId]
  );

  const items: BidClarification[] = cqItems ?? [];

  const [respondingId, setRespondingId] = useState<string | null>(null);
  const [responseText, setResponseText] = useState("");
  const [respondBusy, setRespondBusy] = useState(false);
  const [showAdd, setShowAdd] = useState(false);

  // Add form state
  const [newDirection, setNewDirection] = useState<"us" | "customer">("us");
  const [newQuestion, setNewQuestion] = useState("");
  const [newDeadline, setNewDeadline] = useState("");
  const [addBusy, setAddBusy] = useState(false);

  const SYSTEM_USER = { id: "00000000-0000-0000-0000-000000000001", fullName: "Dev User", email: "dev@ricoh.co.uk" };

  // Auto-number: CQ-XX
  const nextCqNum = `CQ-${String(items.length + 1).padStart(2, "0")}`;

  async function handleRespond(cq: BidClarification) {
    setRespondBusy(true);
    try {
      await dataverseClient.updateClarification(cq.id, {
        cr5ab_responsetext: responseText,
        cr5ab_respondeddate: new Date().toISOString(),
        cr5ab_status: ClarificationStatus.AnswerReceived,
      });
      refresh();
      setRespondingId(null);
      setResponseText("");
    } finally { setRespondBusy(false); }
  }

  async function handleClose(id: string) {
    await dataverseClient.updateClarification(id, { cr5ab_status: ClarificationStatus.Closed });
    refresh();
  }

  async function handleAdd() {
    if (!newQuestion) return;
    setAddBusy(true);
    try {
      await dataverseClient.createClarification({
        cr5ab_bidworkspaceid: { id: workspaceId, cr5ab_title: "" },
        cr5ab_questionnumber: nextCqNum,
        cr5ab_questiontext: newQuestion,
        cr5ab_raisedby: SYSTEM_USER,
        cr5ab_raiseddate: new Date().toISOString(),
        cr5ab_deadline: newDeadline || undefined,
        cr5ab_status: ClarificationStatus.Open,
        cr5ab_iscustomerraised: newDirection === "customer",
      });
      refresh();
      setShowAdd(false);
      setNewQuestion("");
      setNewDeadline("");
    } finally { setAddBusy(false); }
  }

  if (isLoading) return <LoadingState label="Loading clarifications..." />;
  if (error) return <ErrorState message={error} onRetry={refresh} />;

  const open   = items.filter((c) => c.cr5ab_status === ClarificationStatus.Open).length;
  const answered = items.filter((c) => c.cr5ab_status === ClarificationStatus.AnswerReceived).length;
  const closed = items.filter((c) => c.cr5ab_status === ClarificationStatus.Closed).length;

  return (
    <div>
      {/* Summary badges */}
      <div style={{ display: "flex", gap: tokens.spacingHorizontalM, marginBottom: tokens.spacingVerticalM, flexWrap: "wrap", alignItems: "center" }}>
        <Badge appearance="filled" color="warning" size="medium">{open} Open</Badge>
        <Badge appearance="filled" color="informative" size="medium">{answered} Answer Received</Badge>
        <Badge appearance="filled" color="success" size="medium">{closed} Closed</Badge>
        <Button
          appearance="primary"
          size="small"
          icon={<AddRegular />}
          style={{ marginLeft: "auto" }}
          onClick={() => setShowAdd(true)}
        >
          Raise Clarification
        </Button>
      </div>

      {items.length === 0 ? (
        <EmptyState icon={<QuestionCircleRegular />} title="No clarifications yet" description="Raise a clarification question to track customer or internal queries." />
      ) : (
        <Card>
          {items.map((cq) => {
            const isResponding = respondingId === cq.id;
            const days = cq.cr5ab_deadline ? Math.ceil((new Date(cq.cr5ab_deadline).getTime() - Date.now()) / 86400000) : null;
            const isOverdue = days !== null && days < 0;

            return (
              <div key={cq.id} className={styles.cqRow}>
                {/* Direction indicator */}
                <Tooltip content={cq.cr5ab_iscustomerraised ? "Customer raised this" : "We raised this"} relationship="label">
                  <div className={`${styles.cqDirectionBadge} ${cq.cr5ab_iscustomerraised ? styles.cqDirectionCustomer : styles.cqDirectionUs}`}>
                    {cq.cr5ab_iscustomerraised ? "↓" : "↑"}
                  </div>
                </Tooltip>

                <div style={{ flexGrow: 1, minWidth: 0 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS, flexWrap: "wrap", marginBottom: "2px" }}>
                    <Text weight="semibold" size={200} style={{ color: tokens.colorNeutralForeground3, flexShrink: 0 }}>{cq.cr5ab_questionnumber}</Text>
                    <Badge appearance="filled" color={ClarificationStatusColor[cq.cr5ab_status]} size="small">
                      {ClarificationStatusLabel[cq.cr5ab_status]}
                    </Badge>
                    {cq.cr5ab_deadline && (
                      <Badge appearance="outline" color={isOverdue ? "danger" : "subtle"} size="small">
                        {isOverdue ? `${Math.abs(days!)}d overdue` : `Due ${fmtDate(cq.cr5ab_deadline)}`}
                      </Badge>
                    )}
                  </div>

                  <Text size={300}>{cq.cr5ab_questiontext}</Text>
                  <Text size={100} style={{ display: "block", color: tokens.colorNeutralForeground3, marginTop: "2px" }}>
                    Raised by {cq.cr5ab_raisedby.fullName} · {fmtDate(cq.cr5ab_raiseddate)}
                  </Text>

                  {/* Response */}
                  {cq.cr5ab_responsetext && !isResponding && (
                    <div style={{ marginTop: tokens.spacingVerticalXS, padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`, backgroundColor: tokens.colorNeutralBackground2, borderRadius: tokens.borderRadiusMedium, borderLeft: `3px solid ${tokens.colorBrandBackground}` }}>
                      <Text size={200} style={{ color: tokens.colorNeutralForeground2, fontStyle: "italic" }}>
                        {cq.cr5ab_responsetext}
                      </Text>
                      <Text size={100} style={{ display: "block", color: tokens.colorNeutralForeground3, marginTop: "2px" }}>
                        Responded {fmtDate(cq.cr5ab_respondeddate)}
                      </Text>
                    </div>
                  )}

                  {/* Respond area */}
                  {isResponding && (
                    <div style={{ marginTop: tokens.spacingVerticalS, display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS }}>
                      <Textarea
                        placeholder="Type the response..."
                        rows={3}
                        value={responseText}
                        onChange={(_, d) => setResponseText(d.value)}
                      />
                      <div style={{ display: "flex", gap: tokens.spacingHorizontalS }}>
                        <Button size="small" appearance="primary" icon={<CheckmarkRegular />} disabled={respondBusy || !responseText} onClick={() => handleRespond(cq)}>
                          {respondBusy ? <Spinner size="tiny" /> : "Save Response"}
                        </Button>
                        <Button size="small" appearance="subtle" onClick={() => { setRespondingId(null); setResponseText(""); }}>Cancel</Button>
                      </div>
                    </div>
                  )}

                  {/* Action buttons */}
                  {!isResponding && cq.cr5ab_status !== ClarificationStatus.Closed && (
                    <div style={{ marginTop: tokens.spacingVerticalXS, display: "flex", gap: tokens.spacingHorizontalS }}>
                      {cq.cr5ab_status === ClarificationStatus.Open && (
                        <Button size="small" appearance="subtle" icon={<ArrowReplyRegular />} onClick={() => { setRespondingId(cq.id); setResponseText(cq.cr5ab_responsetext ?? ""); }}>
                          {cq.cr5ab_responsetext ? "Edit Response" : "Respond"}
                        </Button>
                      )}
                      <Button size="small" appearance="subtle" icon={<CheckmarkRegular />} onClick={() => handleClose(cq.id)}>
                        Close
                      </Button>
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </Card>
      )}

      {/* Raise Clarification Dialog */}
      <Dialog open={showAdd} onOpenChange={(_, d) => setShowAdd(d.open)}>
        <DialogSurface style={{ maxWidth: "520px" }}>
          <DialogBody>
            <DialogTitle>Raise Clarification</DialogTitle>
            <DialogContent>
              <div style={{ display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM, marginTop: tokens.spacingVerticalM }}>
                {/* Direction toggle */}
                <Field label="Direction">
                  <div style={{ display: "flex", gap: tokens.spacingHorizontalM }}>
                    {(["us", "customer"] as const).map((dir) => (
                      <button
                        key={dir}
                        type="button"
                        onClick={() => setNewDirection(dir)}
                        style={{
                          flex: 1,
                          padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
                          border: `2px solid ${newDirection === dir ? tokens.colorBrandBackground : tokens.colorNeutralStroke2}`,
                          borderRadius: tokens.borderRadiusMedium,
                          backgroundColor: newDirection === dir ? tokens.colorBrandBackground2 : "transparent",
                          cursor: "pointer",
                          fontWeight: newDirection === dir ? tokens.fontWeightSemibold : tokens.fontWeightRegular,
                          color: newDirection === dir ? tokens.colorBrandForeground1 : tokens.colorNeutralForeground2,
                          display: "flex", alignItems: "center", justifyContent: "center", gap: "6px",
                        }}
                      >
                        <span style={{ fontSize: "16px" }}>{dir === "us" ? "↑" : "↓"}</span>
                        {dir === "us" ? "We raised this" : "Customer raised this"}
                      </button>
                    ))}
                  </div>
                </Field>

                <Field label="Question / Clarification" required hint={`Will be numbered ${nextCqNum}`}>
                  <Textarea rows={4} placeholder="Type the clarification question..." value={newQuestion} onChange={(_, d) => setNewQuestion(d.value)} />
                </Field>

                <Field label="Response Deadline">
                  <Input type="date" value={newDeadline} onChange={(_, d) => setNewDeadline(d.value)} />
                </Field>
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="subtle" onClick={() => setShowAdd(false)}>Cancel</Button>
              <Button appearance="primary" disabled={addBusy || !newQuestion} onClick={handleAdd}>
                {addBusy ? <Spinner size="tiny" /> : "Raise"}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Page
// ---------------------------------------------------------------------------

type WorkspaceTab = "overview" | "tor" | "team" | "approvals" | "documents" | "clarifications";

export function BidWorkspacePage() {
  const styles = useStyles();
  const navigate = useNavigate();
  const [searchParams] = useSearchParams();
  const [activeTab, setActiveTab] = useState<WorkspaceTab>("overview");

  const bidId = searchParams.get("bidId");
  const workspaceIdParam = searchParams.get("workspaceId");

  const resolvedWorkspaceId = workspaceIdParam ?? null;

  const { data: workspaceByParam, isLoading: loadingByParam, error: errorByParam, refresh: refreshByParam } = useDataverse(
    () => resolvedWorkspaceId
      ? dataverseClient.getBidWorkspace(resolvedWorkspaceId)
      : bidId
        ? dataverseClient.getWorkspaceForBid(bidId)
        : Promise.resolve(null),
    [resolvedWorkspaceId, bidId]
  );

  const { data: bidRequest } = useDataverse(
    () => (bidId ? dataverseClient.getBidRequest(bidId) : Promise.resolve(null)),
    [bidId]
  );

  const [localStatus, setLocalStatus] = useState<BidStatus | null>(null);
  const [localBidRequest, setLocalBidRequest] = useState<Partial<BidRequest>>({});

  const workspace = workspaceByParam;
  const effectiveStatus = localStatus ?? workspace?.cr5ab_status ?? bidRequest?.cr5ab_status;
  const effectiveBidRequest = bidRequest ? { ...bidRequest, ...localBidRequest } : null;

  const handleStatusChange = useCallback((newStatus: BidStatus) => {
    setLocalStatus(newStatus);
  }, []);

  const handleQualificationDecision = useCallback((changes: Partial<BidRequest>) => {
    setLocalBidRequest((prev) => ({ ...prev, ...changes }));
    if (changes.cr5ab_status) setLocalStatus(changes.cr5ab_status);
  }, []);

  if (loadingByParam) return <LoadingState label="Loading workspace..." />;
  if (errorByParam) return <ErrorState message={errorByParam} onRetry={refreshByParam} />;

  const title = workspace?.cr5ab_title ?? effectiveBidRequest?.cr5ab_title ?? "Bid Workspace";
  const progress = workspace?.cr5ab_completionpercentage ?? 0;
  const teamMembers = workspace?.teamMembers ?? [];
  const approvals = workspace?.approvals ?? [];
  const documents = (workspace?.documents ?? []) as BidDocumentRecord[];
  const workspaceId = workspace?.id ?? `ws-fallback-${bidId}`;

  // Show qualification panel when bid hasn't been qualified yet, or is pending/under qualification
  const needsQualification = effectiveBidRequest && (
    effectiveBidRequest.cr5ab_opportunitystage === OpportunityStage.PendingOAF ||
    effectiveBidRequest.cr5ab_opportunitystage === OpportunityStage.UnderQualification ||
    effectiveBidRequest.cr5ab_opportunitystage === undefined ||
    effectiveBidRequest.cr5ab_qualificationoutcome !== undefined  // show read-only summary even if done
  );

  return (
    <div>
      <PageHeader
        title={title}
        subtitle={
          workspace?.cr5ab_bidrequestid?.cr5ab_bidreferencenumber ??
          bidRequest?.cr5ab_bidreferencenumber ??
          "Workspace"
        }
        actions={
          <>
            <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={() => navigate("/bid-register")}>
              Back to Register
            </Button>
            {workspace?.cr5ab_sharepointfolderurl && (
              <Button appearance="secondary" icon={<LinkRegular />} as="a" href={workspace.cr5ab_sharepointfolderurl} target="_blank" rel="noopener noreferrer">
                SharePoint
              </Button>
            )}
          </>
        }
      />

      {/* Status meta strip */}
      <div className={styles.headerGrid}>
        <div className={styles.metaItem}>
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Status</Text>
          {effectiveStatus !== undefined ? <StatusBadge status={effectiveStatus} /> : <Text>—</Text>}
        </div>
        <div className={styles.metaItem}>
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Stage</Text>
          {effectiveBidRequest?.cr5ab_opportunitystage !== undefined ? (
            <Badge appearance="filled" color={OpportunityStageColor[effectiveBidRequest.cr5ab_opportunitystage!]} size="small">
              {OpportunityStageLabel[effectiveBidRequest.cr5ab_opportunitystage!]}
            </Badge>
          ) : <Text size={200}>—</Text>}
        </div>
        <div className={styles.metaItem}>
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Bid Manager</Text>
          <Text size={300} weight="semibold">{workspace?.cr5ab_bidmanagerid?.fullName ?? effectiveBidRequest?.cr5ab_assignedto?.fullName ?? "—"}</Text>
        </div>
        <div className={styles.metaItem}>
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Submission Deadline</Text>
          <Text size={300} weight="semibold">{fmtDate(effectiveBidRequest?.cr5ab_submissiondeadline)}</Text>
        </div>
      </div>

      {/* Qualification panel */}
      {needsQualification && effectiveBidRequest && (
        <QualificationPanel
          bidRequest={effectiveBidRequest as BidRequest}
          onDecisionRecorded={handleQualificationDecision}
        />
      )}

      {/* Status transition bar */}
      {workspace && (
        <StatusTransitionBar
          workspace={{ ...workspace, cr5ab_status: localStatus ?? workspace.cr5ab_status }}
          onStatusChange={handleStatusChange}
        />
      )}

      {/* Progress bar */}
      <Card className={styles.progressCard}>
        <div className={styles.progressRow}>
          <Text size={300} weight="semibold">Bid Progress</Text>
          <Text size={300}>{progress}%</Text>
        </div>
        <ProgressBar value={progress / 100} color={progress >= 75 ? "success" : progress >= 40 ? "warning" : "error"} />
        <Text size={200} style={{ color: tokens.colorNeutralForeground3, marginTop: tokens.spacingVerticalXS }}>
          {progress < 25 && "Just getting started"}
          {progress >= 25 && progress < 50 && "In progress — key sections underway"}
          {progress >= 50 && progress < 75 && "Good progress — final sections remaining"}
          {progress >= 75 && progress < 100 && "Nearly there — review stage"}
          {progress === 100 && "Complete — ready for submission"}
        </Text>
      </Card>

      {/* Tab navigation */}
      <TabList selectedValue={activeTab} onTabSelect={(_, d) => setActiveTab(d.value as WorkspaceTab)}>
        <Tab value="overview" icon={<CheckmarkCircleRegular />}>Overview</Tab>
        <Tab value="tor" icon={<TableRegular />}>TOR</Tab>
        <Tab value="clarifications" icon={<ChatRegular />}>
          Clarifications
        </Tab>
        <Tab value="team" icon={<PeopleRegular />}>Team</Tab>
        <Tab value="approvals" icon={<ClockRegular />}>Approvals</Tab>
        <Tab value="documents" icon={<DocumentRegular />}>Documents</Tab>
      </TabList>

      <div className={styles.tabContent}>
        {activeTab === "overview" && (
          <OverviewTab workspace={workspace ?? null} bidRequest={effectiveBidRequest ?? null} />
        )}
        {activeTab === "tor" && workspace && (
          <TorTab workspaceId={workspaceId} />
        )}
        {activeTab === "tor" && !workspace && (
          <MessageBar intent="warning">
            <MessageBarBody>A workspace must exist before managing the TOR. Submit the bid and create a workspace first.</MessageBarBody>
          </MessageBar>
        )}
        {activeTab === "clarifications" && workspace && (
          <ClarificationsTab workspaceId={workspaceId} />
        )}
        {activeTab === "clarifications" && !workspace && (
          <MessageBar intent="warning">
            <MessageBarBody>A workspace must exist before managing clarifications.</MessageBarBody>
          </MessageBar>
        )}
        {activeTab === "team" && (
          <TeamTab members={teamMembers} workspaceId={workspaceId} />
        )}
        {activeTab === "approvals" && (
          <ApprovalsTab approvals={approvals} workspaceId={workspaceId} />
        )}
        {activeTab === "documents" && (
          <DocumentsTab documents={documents} workspaceId={workspaceId} sharepointBase={workspace?.cr5ab_sharepointfolderurl} />
        )}
      </div>
    </div>
  );
}
