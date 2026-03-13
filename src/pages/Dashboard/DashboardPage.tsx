/**
 * Dashboard Page
 *
 * Live KPI cards, bid pipeline summary, bid calendar strip, and recent activity.
 * All data is sourced from the Dataverse client (mock in dev).
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
} from "@fluentui/react-components";
import type { ReactNode } from "react";
import {
  AddSquareRegular,
  ArrowRightRegular,
  TrophyRegular,
  DocumentRegular,
  ClockRegular,
  CheckmarkCircleRegular,
} from "@fluentui/react-icons";
import { PageHeader } from "../../components/common/PageHeader";
import { StatusBadge } from "../../components/common/StatusBadge";
import { LoadingState } from "../../components/common/LoadingState";
import { ErrorState } from "../../components/common/ErrorState";
import { useDataverse } from "../../hooks/useDataverse";
import { dataverseClient } from "../../lib/dataverse";
import { BidStatus } from "../../types/dataverse";
import type { BidRequest } from "../../types/dataverse";

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
    "@media (max-width: 900px)": {
      gridTemplateColumns: "1fr",
    },
  },
  kpiCard: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  kpiIcon: {
    width: "40px",
    height: "40px",
    borderRadius: tokens.borderRadiusMedium,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: "20px",
  },
  kpiValue: {},
  pipelineRow: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    ...shorthands.padding(tokens.spacingVerticalS, "0"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ":last-child": {
      borderBottom: "none",
    },
  },
  pipelineRef: {
    minWidth: "130px",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  pipelineTitle: {
    flexGrow: 1,
    minWidth: "0",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  pipelineValue: {
    minWidth: "90px",
    textAlign: "right",
    color: tokens.colorNeutralForeground2,
  },
  calendarStrip: {
    display: "flex",
    gap: tokens.spacingHorizontalM,
    ...shorthands.overflow("hidden", "auto"),
    ...shorthands.padding(tokens.spacingVerticalS, "0"),
  },
  calendarItem: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalL),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    minWidth: "120px",
    gap: tokens.spacingVerticalXS,
    flexShrink: 0,
  },
  calendarItemUrgent: {
    ...shorthands.borderColor(tokens.colorPaletteRedBorder2),
    backgroundColor: tokens.colorPaletteRedBackground1,
  },
});

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function formatCurrency(value: number): string {
  if (value >= 1_000_000) return `£${(value / 1_000_000).toFixed(1)}M`;
  if (value >= 1_000) return `£${(value / 1_000).toFixed(0)}K`;
  return `£${value}`;
}

function formatDeadline(iso: string): string {
  return new Date(iso).toLocaleDateString("en-GB", {
    day: "numeric",
    month: "short",
  });
}

function daysUntil(iso: string): number {
  const ms = new Date(iso).getTime() - Date.now();
  return Math.ceil(ms / (1000 * 60 * 60 * 24));
}

// ---------------------------------------------------------------------------
// Sub-components
// ---------------------------------------------------------------------------

interface KpiCardProps {
  label: string;
  value: string | number;
  icon: ReactNode;
  iconBg: string;
  iconColor: string;
  trend?: string;
}

function KpiCard({ label, value, icon, iconBg, iconColor, trend }: KpiCardProps) {
  const styles = useStyles();
  return (
    <Card>
      <div className={styles.kpiCard}>
        <div
          className={styles.kpiIcon}
          style={{ backgroundColor: iconBg, color: iconColor }}
        >
          {icon}
        </div>
        <Text size={600} weight="bold" className={styles.kpiValue}>
          {value}
        </Text>
        <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
          {label}
        </Text>
        {trend && (
          <Text size={200} style={{ color: tokens.colorPaletteGreenForeground1 }}>
            {trend}
          </Text>
        )}
      </div>
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Page
// ---------------------------------------------------------------------------

export function DashboardPage() {
  const styles = useStyles();
  const navigate = useNavigate();

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
    const pipeline = bids
      .filter((b) => b.cr5ab_estimatedvalue)
      .reduce((sum, b) => sum + (b.cr5ab_estimatedvalue ?? 0), 0);
    const winRate = bids.length > 0 ? Math.round((won.length / bids.length) * 100) : 0;

    return { activeBids: active.length, wonBids: won.length, pipeline, winRate };
  }, [bids]);

  const upcomingDeadlines = useMemo(
    () =>
      bids
        .filter((b) =>
          [BidStatus.Draft, BidStatus.Submitted, BidStatus.InProgress, BidStatus.InReview].includes(b.cr5ab_status)
        )
        .sort(
          (a, b) =>
            new Date(a.cr5ab_submissiondeadline).getTime() -
            new Date(b.cr5ab_submissiondeadline).getTime()
        )
        .slice(0, 5),
    [bids]
  );

  const activePipeline = useMemo(
    () =>
      bids
        .filter((b) =>
          [BidStatus.InProgress, BidStatus.InReview, BidStatus.Submitted].includes(b.cr5ab_status)
        )
        .slice(0, 6),
    [bids]
  );

  if (isLoading) return <LoadingState label="Loading dashboard..." />;
  if (error) return <ErrorState message={error} onRetry={refresh} />;

  return (
    <div>
      <PageHeader
        title="Dashboard"
        subtitle="Ricoh Bid Management — live overview"
        actions={
          <Button
            appearance="primary"
            icon={<AddSquareRegular />}
            onClick={() => navigate("/new-bid")}
          >
            New Bid
          </Button>
        }
      />

      {/* KPI Cards */}
      <div className={styles.grid}>
        <KpiCard
          label="Active Bids"
          value={kpis.activeBids}
          icon={<DocumentRegular />}
          iconBg={tokens.colorBrandBackground2}
          iconColor={tokens.colorBrandForeground1}
        />
        <KpiCard
          label="Bids Won"
          value={kpis.wonBids}
          icon={<TrophyRegular />}
          iconBg={tokens.colorPaletteGreenBackground2}
          iconColor={tokens.colorPaletteGreenForeground1}
        />
        <KpiCard
          label="Total Pipeline"
          value={formatCurrency(kpis.pipeline)}
          icon={<CheckmarkCircleRegular />}
          iconBg={tokens.colorPalettePurpleBackground2}
          iconColor={tokens.colorPalettePurpleForeground2}
        />
        <KpiCard
          label="Win Rate"
          value={`${kpis.winRate}%`}
          icon={<ClockRegular />}
          iconBg={tokens.colorPaletteMarigoldBackground2}
          iconColor={tokens.colorPaletteMarigoldForeground2}
        />
      </div>

      {/* Pipeline + Calendar */}
      <div className={styles.twoCol}>
        {/* Active Pipeline */}
        <Card>
          <CardHeader
            header={
              <Text weight="semibold" size={400}>
                Active Pipeline
              </Text>
            }
            action={
              <Button
                appearance="transparent"
                size="small"
                icon={<ArrowRightRegular />}
                iconPosition="after"
                onClick={() => navigate("/bid-register")}
              >
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

        {/* Upcoming Deadlines */}
        <Card>
          <CardHeader
            header={
              <Text weight="semibold" size={400}>
                Upcoming Deadlines
              </Text>
            }
          />
          <div className={styles.calendarStrip}>
            {upcomingDeadlines.length === 0 ? (
              <Text style={{ color: tokens.colorNeutralForeground3 }}>
                No upcoming deadlines.
              </Text>
            ) : (
              upcomingDeadlines.map((bid) => {
                const days = daysUntil(bid.cr5ab_submissiondeadline);
                const urgent = days <= 7;
                return (
                  <div
                    key={bid.id}
                    className={`${styles.calendarItem} ${urgent ? styles.calendarItemUrgent : ""}`}
                  >
                    <Text size={500} weight="bold">
                      {formatDeadline(bid.cr5ab_submissiondeadline)}
                    </Text>
                    <Text
                      size={200}
                      style={{
                        textAlign: "center",
                        color: tokens.colorNeutralForeground2,
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        whiteSpace: "nowrap",
                        maxWidth: "100px",
                      }}
                    >
                      {bid.cr5ab_title}
                    </Text>
                    <Badge
                      appearance="filled"
                      color={urgent ? "danger" : "informative"}
                      size="small"
                    >
                      {days <= 0 ? "Overdue" : `${days}d`}
                    </Badge>
                  </div>
                );
              })
            )}
          </div>
        </Card>
      </div>

      {/* Win Rate progress */}
      <Card>
        <CardHeader
          header={
            <Text weight="semibold" size={400}>
              Win Rate Progress
            </Text>
          }
        />
        <div style={{ display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS }}>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <Text size={300}>Overall win rate</Text>
            <Text size={300} weight="semibold">{kpis.winRate}%</Text>
          </div>
          <ProgressBar
            value={kpis.winRate / 100}
            color={kpis.winRate >= 50 ? "success" : kpis.winRate >= 30 ? "warning" : "error"}
          />
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
            Target: 50% | Based on {bids.length} bids
          </Text>
        </div>
      </Card>
    </div>
  );
}

function PipelineRow({ bid }: { bid: BidRequest }) {
  const styles = useStyles();
  return (
    <div className={styles.pipelineRow}>
      <Text className={styles.pipelineRef} size={200}>
        {bid.cr5ab_bidreferencenumber}
      </Text>
      <Text className={styles.pipelineTitle} size={300}>
        {bid.cr5ab_title}
      </Text>
      <div style={{ minWidth: "110px", display: "flex", justifyContent: "center" }}>
        <StatusBadge status={bid.cr5ab_status} />
      </div>
      <Text className={styles.pipelineValue} size={200}>
        {bid.cr5ab_estimatedvalue
          ? formatCurrency(bid.cr5ab_estimatedvalue)
          : "—"}
      </Text>
    </div>
  );
}
