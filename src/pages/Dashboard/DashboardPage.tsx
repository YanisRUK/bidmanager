/**
 * Dashboard Page
 *
 * - KPI cards: active bids, won, pipeline value, win rate
 * - Active pipeline summary table
 * - Full calendar month view showing submission deadlines, award dates
 * - My Tasks panel — TOR items assigned to the logged-in user, colour-coded by urgency
 * - Win rate progress bar
 */

import { useMemo } from "react";
import { useNavigate } from "react-router-dom";
import {
  makeStyles,
  shorthands,
  tokens,
  Text,
  Card,
  CardHeader,
  Button,
  ProgressBar,
  Badge,
  Tooltip,
} from "@fluentui/react-components";
import type { ReactNode } from "react";
import {
  AddSquareRegular,
  ArrowRightRegular,
  TrophyRegular,
  DocumentRegular,
  ClockRegular,
  CheckmarkCircleRegular,
  PersonRegular,
  CalendarRegular,
} from "@fluentui/react-icons";
import { PageHeader } from "../../components/common/PageHeader";
import { StatusBadge } from "../../components/common/StatusBadge";
import { LoadingState } from "../../components/common/LoadingState";
import { ErrorState } from "../../components/common/ErrorState";
import { EmptyState } from "../../components/common/EmptyState";
import { useDataverse } from "../../hooks/useDataverse";
import { dataverseClient } from "../../lib/dataverse";
import { BidStatus, TorAnsweredStatus, ClarificationStatus } from "../../types/dataverse";
import { useAuth } from "../../context/AuthContext";
import type { BidRequest, BidClarification, TorItem } from "../../types/dataverse";

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

const useStyles = makeStyles({
  grid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))",
    gap: tokens.spacingVerticalL,
    marginBottom: tokens.spacingVerticalXL,
  },
  twoCol: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: tokens.spacingVerticalL,
    marginBottom: tokens.spacingVerticalXL,
    "@media (max-width: 900px)": { gridTemplateColumns: "1fr" },
  },
  threeCol: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr 1fr",
    gap: tokens.spacingVerticalL,
    marginBottom: tokens.spacingVerticalXL,
    "@media (max-width: 1100px)": { gridTemplateColumns: "1fr 1fr" },
    "@media (max-width: 700px)": { gridTemplateColumns: "1fr" },
  },
  kpiCard: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  kpiIcon: {
    width: "40px", height: "40px", borderRadius: tokens.borderRadiusMedium,
    display: "flex", alignItems: "center", justifyContent: "center", fontSize: "20px",
  },
  pipelineRow: {
    display: "flex", alignItems: "center", gap: tokens.spacingHorizontalM,
    ...shorthands.padding(tokens.spacingVerticalS, "0"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ":last-child": { borderBottom: "none" },
  },
  pipelineRef: {
    minWidth: "130px", color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200,
  },
  pipelineTitle: { flexGrow: 1, minWidth: "0", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  pipelineValue: { minWidth: "90px", textAlign: "right", color: tokens.colorNeutralForeground2 },
  // Calendar
  calendarGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(7, 1fr)",
    gap: "2px",
  },
  calDayHeader: {
    textAlign: "center",
    ...shorthands.padding(tokens.spacingVerticalXS),
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    fontWeight: tokens.fontWeightSemibold,
  },
  calDay: {
    minHeight: "60px",
    ...shorthands.padding("4px"),
    borderRadius: tokens.borderRadiusSmall,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    display: "flex",
    flexDirection: "column",
    gap: "2px",
  },
  calDayToday: { border: `2px solid ${tokens.colorBrandBackground}`, backgroundColor: tokens.colorBrandBackground2 },
  calDayOtherMonth: { backgroundColor: tokens.colorNeutralBackground3, opacity: "0.5" },
  calDayNum: { fontSize: tokens.fontSizeBase100, fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground2 },
  calEvent: {
    fontSize: "10px",
    ...shorthands.borderRadius("2px"),
    ...shorthands.padding("1px", "3px"),
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    cursor: "pointer",
    lineHeight: "14px",
  },
  // My Tasks
  taskRow: {
    display: "flex", alignItems: "flex-start", gap: tokens.spacingHorizontalM,
    ...shorthands.padding(tokens.spacingVerticalS, "0"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ":last-child": { borderBottom: "none" },
  },
  taskUrgencyBar: { width: "4px", borderRadius: "2px", flexShrink: 0, alignSelf: "stretch" },
});

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function formatCurrency(value: number): string {
  if (value >= 1_000_000) return `£${(value / 1_000_000).toFixed(1)}M`;
  if (value >= 1_000) return `£${(value / 1_000).toFixed(0)}K`;
  return `£${value}`;
}

function daysUntil(iso: string): number {
  return Math.ceil((new Date(iso).getTime() - Date.now()) / 86400000);
}

function fmtShort(iso: string) {
  return new Date(iso).toLocaleDateString("en-GB", { day: "numeric", month: "short" });
}

// ---------------------------------------------------------------------------
// KPI Card
// ---------------------------------------------------------------------------

interface KpiCardProps {
  label: string; value: string | number; icon: ReactNode; iconBg: string; iconColor: string; trend?: string;
}

function KpiCard({ label, value, icon, iconBg, iconColor, trend }: KpiCardProps) {
  const styles = useStyles();
  return (
    <Card>
      <div className={styles.kpiCard}>
        <div className={styles.kpiIcon} style={{ backgroundColor: iconBg, color: iconColor }}>{icon}</div>
        <Text size={600} weight="bold">{value}</Text>
        <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>{label}</Text>
        {trend && <Text size={200} style={{ color: tokens.colorPaletteGreenForeground1 }}>{trend}</Text>}
      </div>
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Pipeline Row
// ---------------------------------------------------------------------------

function PipelineRow({ bid }: { bid: BidRequest }) {
  const styles = useStyles();
  const navigate = useNavigate();
  return (
    <div className={styles.pipelineRow} style={{ cursor: "pointer" }} onClick={() => navigate(`/workspace?bidId=${bid.id}`)}>
      <Text className={styles.pipelineRef} size={200}>{bid.cr5ab_bidreferencenumber}</Text>
      <Text className={styles.pipelineTitle} size={300}>{bid.cr5ab_title}</Text>
      <div style={{ minWidth: "110px", display: "flex", justifyContent: "center" }}>
        <StatusBadge status={bid.cr5ab_status} />
      </div>
      <Text className={styles.pipelineValue} size={200}>
        {bid.cr5ab_estimatedvalue ? formatCurrency(bid.cr5ab_estimatedvalue) : "—"}
      </Text>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Calendar Month View
// ---------------------------------------------------------------------------

interface CalendarEvent {
  date: Date;
  label: string;
  color: string;
  bgColor: string;
  bidId?: string;
}

function CalendarMonthView({ events }: { events: CalendarEvent[] }) {
  const styles = useStyles();
  const navigate = useNavigate();

  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth();

  const firstDay = new Date(year, month, 1);
  const lastDay = new Date(year, month + 1, 0);
  // Start from Sunday of the week containing the 1st
  const startDate = new Date(firstDay);
  startDate.setDate(firstDay.getDate() - firstDay.getDay());

  // Build calendar cells (6 rows × 7 = 42 cells max)
  const cells: Date[] = [];
  const d = new Date(startDate);
  while (d <= lastDay || cells.length % 7 !== 0 || cells.length < 35) {
    cells.push(new Date(d));
    d.setDate(d.getDate() + 1);
    if (cells.length >= 42) break;
  }

  const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

  const eventMap = new Map<string, CalendarEvent[]>();
  for (const ev of events) {
    const key = ev.date.toDateString();
    if (!eventMap.has(key)) eventMap.set(key, []);
    eventMap.get(key)!.push(ev);
  }

  const monthLabel = today.toLocaleDateString("en-GB", { month: "long", year: "numeric" });

  return (
    <div>
      <Text weight="semibold" size={300} style={{ display: "block", marginBottom: tokens.spacingVerticalS, textAlign: "center" }}>{monthLabel}</Text>
      <div className={styles.calendarGrid}>
        {dayNames.map((d) => (
          <div key={d} className={styles.calDayHeader}>{d}</div>
        ))}
        {cells.map((cell, i) => {
          const isToday = cell.toDateString() === today.toDateString();
          const isCurrentMonth = cell.getMonth() === month;
          const cellEvents = eventMap.get(cell.toDateString()) ?? [];
          return (
            <div
              key={i}
              className={`${styles.calDay} ${isToday ? styles.calDayToday : ""} ${!isCurrentMonth ? styles.calDayOtherMonth : ""}`}
            >
              <span className={styles.calDayNum}>{cell.getDate()}</span>
              {cellEvents.slice(0, 3).map((ev, j) => (
                <Tooltip key={j} content={ev.label} relationship="label">
                  <div
                    className={styles.calEvent}
                    style={{ backgroundColor: ev.bgColor, color: ev.color }}
                    onClick={() => ev.bidId && navigate(`/workspace?bidId=${ev.bidId}`)}
                  >
                    {ev.label}
                  </div>
                </Tooltip>
              ))}
              {cellEvents.length > 3 && (
                <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>+{cellEvents.length - 3} more</Text>
              )}
            </div>
          );
        })}
      </div>
      {/* Legend */}
      <div style={{ display: "flex", gap: tokens.spacingHorizontalL, marginTop: tokens.spacingVerticalS, flexWrap: "wrap" }}>
        {[
          { color: "#c50f1f", bg: "#fde7e9", label: "Submission Deadline" },
          { color: "#8764b8", bg: "#f4ecff", label: "Expected Award" },
          { color: "#0078d4", bg: "#dce9ff", label: "Clarification Deadline" },
        ].map(({ color, bg, label }) => (
          <div key={label} style={{ display: "flex", alignItems: "center", gap: "6px" }}>
            <div style={{ width: "12px", height: "12px", borderRadius: "2px", backgroundColor: bg, border: `1px solid ${color}` }} />
            <Text size={100}>{label}</Text>
          </div>
        ))}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// My Tasks Panel
// ---------------------------------------------------------------------------

// Unified task item — covers TOR rows and open clarifications
interface UnifiedTask {
  id: string;
  type: "tor" | "clarification";
  label: string;
  sublabel: string;
  deadline: string | null;
  workspaceId: string;
}

function MyTasksPanel({ userId }: { userId: string }) {
  const styles = useStyles();
  const navigate = useNavigate();

  const { data: torTasks, isLoading: torLoading, error: torError, refresh: torRefresh } = useDataverse(
    () => dataverseClient.getTorItemsForUser(userId),
    [userId]
  );
  const { data: cqTasks, isLoading: cqLoading, error: cqError, refresh: cqRefresh } = useDataverse(
    () => dataverseClient.getClarificationsForUser(userId),
    [userId]
  );

  const isLoading = torLoading || cqLoading;
  const error = torError || cqError;

  if (isLoading) return <LoadingState label="Loading your tasks..." />;
  if (error) return <ErrorState message={error} onRetry={() => { torRefresh(); cqRefresh(); }} />;

  const pendingTor: UnifiedTask[] = ((torTasks ?? []) as TorItem[])
    .filter((t) => t.cr5ab_answeredstatus !== TorAnsweredStatus.Yes)
    .map((t) => ({
      id: t.id,
      type: "tor",
      label: `${t.cr5ab_questionnumber} — ${t.cr5ab_questiondetail.slice(0, 75)}${t.cr5ab_questiondetail.length > 75 ? "…" : ""}`,
      sublabel: `TOR · ${t.cr5ab_section}`,
      deadline: t.cr5ab_actualdeadline ?? null,
      workspaceId: t.cr5ab_bidworkspaceid.id,
    }));

  const pendingCq: UnifiedTask[] = ((cqTasks ?? []) as BidClarification[])
    .filter((c) => c.cr5ab_status !== ClarificationStatus.Closed)
    .map((c) => ({
      id: c.id,
      type: "clarification",
      label: `${c.cr5ab_questionnumber} — ${c.cr5ab_questiontext.slice(0, 75)}${c.cr5ab_questiontext.length > 75 ? "…" : ""}`,
      sublabel: `Clarification · ${c.cr5ab_status === ClarificationStatus.AnswerReceived ? "Answer received" : "Awaiting response"}`,
      deadline: c.cr5ab_deadline ?? null,
      workspaceId: c.cr5ab_bidworkspaceid.id,
    }));

  const allTasks = [...pendingTor, ...pendingCq].sort((a, b) => {
    const da = a.deadline ? new Date(a.deadline).getTime() : Infinity;
    const db = b.deadline ? new Date(b.deadline).getTime() : Infinity;
    return da - db;
  });

  if (allTasks.length === 0) {
    return (
      <EmptyState
        icon={<CheckmarkCircleRegular />}
        title="All caught up"
        description="No outstanding TOR tasks or clarifications assigned to you."
      />
    );
  }

  return (
    <div>
      {allTasks.map((task) => {
        const days = task.deadline ? daysUntil(task.deadline) : null;
        const isOverdue = days !== null && days < 0;
        const isUrgent = days !== null && days >= 0 && days <= 3;
        const urgencyColor = isOverdue ? tokens.colorPaletteRedBackground3
          : isUrgent ? tokens.colorPaletteMarigoldBackground3
          : tokens.colorPaletteGreenBackground3;

        return (
          <div
            key={task.id}
            className={styles.taskRow}
            style={{ cursor: "pointer" }}
            onClick={() => navigate(`/workspace?workspaceId=${task.workspaceId}`)}
          >
            <div className={styles.taskUrgencyBar} style={{ backgroundColor: urgencyColor }} />
            <div style={{ flexGrow: 1, minWidth: 0 }}>
              <Text size={300} weight="semibold" style={{ display: "block" }}>
                {task.label}
              </Text>
              <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>
                {task.sublabel}
              </Text>
            </div>
            <div style={{ textAlign: "right", flexShrink: 0 }}>
              {days !== null ? (
                <Badge
                  appearance="filled"
                  color={isOverdue ? "danger" : isUrgent ? "warning" : "success"}
                  size="small"
                >
                  {isOverdue ? `${Math.abs(days)}d overdue` : days === 0 ? "Due today" : `${days}d`}
                </Badge>
              ) : (
                <Badge appearance="outline" color="subtle" size="small">No deadline</Badge>
              )}
              <Text size={100} style={{ display: "block", color: tokens.colorNeutralForeground3, marginTop: "2px" }}>
                {task.deadline ? fmtShort(task.deadline) : "—"}
              </Text>
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Page
// ---------------------------------------------------------------------------

export function DashboardPage() {
  const styles = useStyles();
  const navigate = useNavigate();
  const { user } = useAuth();

  const { data, isLoading, error, refresh } = useDataverse(
    () => dataverseClient.getBidRequests(),
    []
  );

  const bids = data?.value ?? [];

  const kpis = useMemo(() => {
    const active = bids.filter((b) =>
      [BidStatus.InProgress, BidStatus.InReview, BidStatus.Qualified, BidStatus.Submitted].includes(b.cr5ab_status)
    );
    const won = bids.filter((b) => b.cr5ab_status === BidStatus.Won);
    const pipeline = bids.filter((b) => b.cr5ab_estimatedvalue).reduce((sum, b) => sum + (b.cr5ab_estimatedvalue ?? 0), 0);
    const winRate = bids.length > 0 ? Math.round((won.length / bids.length) * 100) : 0;
    return { activeBids: active.length, wonBids: won.length, pipeline, winRate };
  }, [bids]);

  const activePipeline = useMemo(
    () => bids.filter((b) => [BidStatus.InProgress, BidStatus.InReview, BidStatus.Submitted, BidStatus.Qualified].includes(b.cr5ab_status)).slice(0, 6),
    [bids]
  );

  // Fetch all clarifications for active workspaces to populate calendar
  const { data: allClarifications } = useDataverse(
    () => dataverseClient.getClarificationsForUser("all"),
    []
  );

  // Build calendar events from bids + clarification deadlines
  const calendarEvents = useMemo((): CalendarEvent[] => {
    const events: CalendarEvent[] = [];
    const activeBids = bids.filter((b) => ![BidStatus.Won, BidStatus.Lost, BidStatus.Withdrawn].includes(b.cr5ab_status));
    for (const bid of activeBids) {
      if (bid.cr5ab_submissiondeadline) {
        events.push({
          date: new Date(bid.cr5ab_submissiondeadline),
          label: bid.cr5ab_bidreferencenumber || bid.cr5ab_title.slice(0, 12),
          color: "#c50f1f", bgColor: "#fde7e9",
          bidId: bid.id,
        });
      }
      if (bid.cr5ab_expectedawarddate) {
        events.push({
          date: new Date(bid.cr5ab_expectedawarddate),
          label: `${bid.cr5ab_bidreferencenumber || bid.cr5ab_title.slice(0, 12)} (award)`,
          color: "#8764b8", bgColor: "#f4ecff",
          bidId: bid.id,
        });
      }
    }
    // Add clarification deadlines (blue)
    for (const cq of (allClarifications ?? []) as BidClarification[]) {
      if (cq.cr5ab_deadline && cq.cr5ab_status !== ClarificationStatus.Closed) {
        events.push({
          date: new Date(cq.cr5ab_deadline),
          label: `${cq.cr5ab_questionnumber} CQ`,
          color: "#0078d4", bgColor: "#dce9ff",
        });
      }
    }
    return events;
  }, [bids, allClarifications]);

  if (isLoading) return <LoadingState label="Loading dashboard..." />;
  if (error) return <ErrorState message={error} onRetry={refresh} />;

  return (
    <div>
      <PageHeader
        title="Dashboard"
        subtitle="Ricoh Bid Management — live overview"
        actions={
          <Button appearance="primary" icon={<AddSquareRegular />} onClick={() => navigate("/new-bid")}>
            New Bid
          </Button>
        }
      />

      {/* KPI Cards */}
      <div className={styles.grid}>
        <KpiCard label="Active Bids"    value={kpis.activeBids}             icon={<DocumentRegular />}        iconBg={tokens.colorBrandBackground2}           iconColor={tokens.colorBrandForeground1} />
        <KpiCard label="Bids Won"       value={kpis.wonBids}                icon={<TrophyRegular />}          iconBg={tokens.colorPaletteGreenBackground2}     iconColor={tokens.colorPaletteGreenForeground1} />
        <KpiCard label="Total Pipeline" value={formatCurrency(kpis.pipeline)} icon={<CheckmarkCircleRegular />} iconBg={tokens.colorPalettePurpleBackground2}    iconColor={tokens.colorPalettePurpleForeground2} />
        <KpiCard label="Win Rate"       value={`${kpis.winRate}%`}          icon={<ClockRegular />}           iconBg={tokens.colorPaletteMarigoldBackground2}  iconColor={tokens.colorPaletteMarigoldForeground2} />
      </div>

      {/* Pipeline + Calendar */}
      <div className={styles.twoCol}>
        {/* Active Pipeline */}
        <Card>
          <CardHeader
            header={<Text weight="semibold" size={400}>Active Pipeline</Text>}
            action={
              <Button appearance="transparent" size="small" icon={<ArrowRightRegular />} iconPosition="after" onClick={() => navigate("/bid-register")}>
                View all
              </Button>
            }
          />
          {activePipeline.length === 0 ? (
            <Text style={{ color: tokens.colorNeutralForeground3 }}>No active bids.</Text>
          ) : (
            activePipeline.map((bid) => <PipelineRow key={bid.id} bid={bid} />)
          )}
        </Card>

        {/* Calendar Month View */}
        <Card>
          <CardHeader
            header={<Text weight="semibold" size={400}>Bid Calendar</Text>}
            description={<Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Deadlines and award dates this month</Text>}
            action={<CalendarRegular style={{ fontSize: "20px", color: tokens.colorNeutralForeground3 }} />}
          />
          <CalendarMonthView events={calendarEvents} />
        </Card>
      </div>

      {/* My Tasks + Win Rate */}
      <div className={styles.twoCol}>
        {/* My Tasks */}
        <Card>
          <CardHeader
            header={<Text weight="semibold" size={400}>My Tasks</Text>}
            description={<Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>TOR items assigned to you — sorted by urgency</Text>}
            action={<PersonRegular style={{ fontSize: "20px", color: tokens.colorNeutralForeground3 }} />}
          />
          {user ? (
            <MyTasksPanel userId={user.id} />
          ) : (
            <EmptyState title="Not signed in" description="Sign in to see your tasks." />
          )}
        </Card>

        {/* Win Rate */}
        <Card>
          <CardHeader header={<Text weight="semibold" size={400}>Win Rate Progress</Text>} />
          <div style={{ display: "flex", flexDirection: "column", gap: tokens.spacingVerticalL }}>
            <div style={{ display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS }}>
              <div style={{ display: "flex", justifyContent: "space-between" }}>
                <Text size={300}>Overall win rate</Text>
                <Text size={300} weight="semibold">{kpis.winRate}%</Text>
              </div>
              <ProgressBar value={kpis.winRate / 100} color={kpis.winRate >= 50 ? "success" : kpis.winRate >= 30 ? "warning" : "error"} />
              <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                Target: 50% | Based on {bids.length} bids
              </Text>
            </div>

            {/* Per-status breakdown */}
            {[
              { label: "In Progress", filter: BidStatus.InProgress },
              { label: "Submitted",   filter: BidStatus.Submitted },
              { label: "In Review",   filter: BidStatus.InReview },
              { label: "Qualified",   filter: BidStatus.Qualified },
            ].map(({ label, filter }) => {
              const count = bids.filter((b) => b.cr5ab_status === filter).length;
              return (
                <div key={label} style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <Text size={200}>{label}</Text>
                    <Text size={200} weight="semibold">{count}</Text>
                  </div>
                  <ProgressBar value={bids.length > 0 ? count / bids.length : 0} color="brand" />
                </div>
              );
            })}
          </div>
        </Card>
      </div>
    </div>
  );
}
