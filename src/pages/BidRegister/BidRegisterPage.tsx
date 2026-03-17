/**
 * Bid Register Page
 *
 * Central searchable/filterable table of all bids across all types and statuses.
 */

import { useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import {
  makeStyles,
  shorthands,
  tokens,
  Text,
  Button,
  Input,
  Select,
  Table,
  TableHeader,
  TableHeaderCell,
  TableBody,
  TableRow,
  TableCell,
  TableCellLayout,
  Card,
  Badge,
} from "@fluentui/react-components";
import {
  AddSquareRegular,
  SearchRegular,
  FilterRegular,
  ArrowRightRegular,
} from "@fluentui/react-icons";
import { PageHeader } from "../../components/common/PageHeader";
import { StatusBadge } from "../../components/common/StatusBadge";
import { LoadingState } from "../../components/common/LoadingState";
import { ErrorState } from "../../components/common/ErrorState";
import { EmptyState } from "../../components/common/EmptyState";
import { useDataverse } from "../../hooks/useDataverse";
import { dataverseClient } from "../../lib/dataverse";
import {
  BidStatusLabel,
  BidTypeLabel,
  BidTypeCode,
  BidSourceLabel,
  OpportunityStageLabel,
  OpportunityStageColor,
} from "../../types/dataverse";

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

const useStyles = makeStyles({
  toolbar: {
    display: "flex",
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
    flexWrap: "wrap",
    alignItems: "center",
  },
  searchInput: {
    flexGrow: 1,
    minWidth: "200px",
    maxWidth: "360px",
  },
  filterGroup: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    flexWrap: "wrap",
    alignItems: "center",
    marginLeft: "auto",
  },
  tableCard: {
    ...shorthands.overflow("hidden"),
  },
  refCell: {
    color: tokens.colorNeutralForeground3,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
  },
  valueCell: {
    fontVariantNumeric: "tabular-nums",
  },
  deadlineCell: {
    whiteSpace: "nowrap",
  },
  actionCell: {
    width: "50px",
  },
  rowCount: {
    color: tokens.colorNeutralForeground3,
  },
});

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function formatCurrency(value?: number) {
  if (!value) return "—";
  if (value >= 1_000_000) return `£${(value / 1_000_000).toFixed(1)}M`;
  if (value >= 1_000) return `£${(value / 1_000).toFixed(0)}K`;
  return `£${value}`;
}

function formatDate(iso: string) {
  return new Date(iso).toLocaleDateString("en-GB", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
}

const ALL_STATUSES = "all";
const ALL_TYPES = "all";
const ALL_STAGES = "all";

// ---------------------------------------------------------------------------
// Page
// ---------------------------------------------------------------------------

export function BidRegisterPage() {
  const styles = useStyles();
  const navigate = useNavigate();

  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState<string>(ALL_STATUSES);
  const [typeFilter, setTypeFilter] = useState<string>(ALL_TYPES);
  const [stageFilter, setStageFilter] = useState<string>(ALL_STAGES);

  const { data, isLoading, error, refresh } = useDataverse(
    () => dataverseClient.getBidRequests(),
    []
  );

  const filtered = useMemo(() => {
    let bids = data?.value ?? [];

    if (search.trim()) {
      const q = search.toLowerCase();
      bids = bids.filter(
        (b) =>
          b.cr5ab_title.toLowerCase().includes(q) ||
          b.cr5ab_bidreferencenumber.toLowerCase().includes(q) ||
          b.cr5ab_customername.toLowerCase().includes(q)
      );
    }

    if (statusFilter !== ALL_STATUSES) {
      bids = bids.filter((b) => b.cr5ab_status === Number(statusFilter));
    }

    if (typeFilter !== ALL_TYPES) {
      bids = bids.filter(
        (b) => b.cr5ab_bidtypeid.cr5ab_code === Number(typeFilter)
      );
    }

    if (stageFilter !== ALL_STAGES) {
      bids = bids.filter((b) => String(b.cr5ab_opportunitystage ?? "") === stageFilter);
    }

    return bids;
  }, [data, search, statusFilter, typeFilter, stageFilter]);

  if (isLoading) return <LoadingState label="Loading bid register..." />;
  if (error) return <ErrorState message={error} onRetry={refresh} />;

  return (
    <div>
      <PageHeader
        title="Bid Workspaces"
        subtitle={`${data?.totalCount ?? 0} bids — select one to open its workspace`}
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

      {/* Toolbar */}
      <div className={styles.toolbar}>
        <Input
          className={styles.searchInput}
          placeholder="Search by title, reference or customer..."
          contentBefore={<SearchRegular />}
          value={search}
          onChange={(_, d) => setSearch(d.value)}
        />
        <div className={styles.filterGroup}>
          <FilterRegular style={{ color: tokens.colorNeutralForeground3 }} />
          <Select
            value={statusFilter}
            onChange={(_, d) => setStatusFilter(d.value)}
            aria-label="Filter by status"
          >
            <option value={ALL_STATUSES}>All Statuses</option>
            {Object.entries(BidStatusLabel).map(([code, label]) => (
              <option key={code} value={code}>
                {label}
              </option>
            ))}
          </Select>
          <Select
            value={typeFilter}
            onChange={(_, d) => setTypeFilter(d.value)}
            aria-label="Filter by bid type"
          >
            <option value={ALL_TYPES}>All Types</option>
            {Object.entries(BidTypeLabel).map(([code, label]) => (
              <option key={code} value={code}>
                {label}
              </option>
            ))}
          </Select>
          <Select
            value={stageFilter}
            onChange={(_, d) => setStageFilter(d.value)}
            aria-label="Filter by stage"
          >
            <option value={ALL_STAGES}>All Stages</option>
            {Object.entries(OpportunityStageLabel).map(([code, label]) => (
              <option key={code} value={code}>{label}</option>
            ))}
          </Select>
        </div>
        <Text size={200} className={styles.rowCount}>
          {filtered.length} result{filtered.length !== 1 ? "s" : ""}
        </Text>
      </div>

      {/* Table */}
      <Card className={styles.tableCard}>
        {filtered.length === 0 ? (
          <EmptyState
            title="No bids found"
            description="Try adjusting your search or filters."
          />
        ) : (
          <Table aria-label="Bid register" sortable>
            <TableHeader>
              <TableRow>
                <TableHeaderCell>Reference</TableHeaderCell>
                <TableHeaderCell>Title</TableHeaderCell>
                <TableHeaderCell>Customer</TableHeaderCell>
                <TableHeaderCell>Type</TableHeaderCell>
                <TableHeaderCell>Stage</TableHeaderCell>
                <TableHeaderCell>Source</TableHeaderCell>
                <TableHeaderCell>Status</TableHeaderCell>
                <TableHeaderCell>Deadline</TableHeaderCell>
                <TableHeaderCell>Value</TableHeaderCell>
                <TableHeaderCell className={styles.actionCell} />
              </TableRow>
            </TableHeader>
            <TableBody>
              {filtered.map((bid) => (
                <TableRow
                  key={bid.id}
                  onClick={() => navigate(`/workspace?bidId=${bid.id}`)}
                  style={{ cursor: "pointer" }}
                >
                  <TableCell>
                    <TableCellLayout>
                      <Text className={styles.refCell}>
                        {bid.cr5ab_bidreferencenumber}
                      </Text>
                    </TableCellLayout>
                  </TableCell>
                  <TableCell>
                    <TableCellLayout>
                      <Text weight="semibold" size={300}>
                        {bid.cr5ab_title}
                      </Text>
                    </TableCellLayout>
                  </TableCell>
                  <TableCell>
                    <TableCellLayout>
                      <Text size={300}>{bid.cr5ab_customername}</Text>
                    </TableCellLayout>
                  </TableCell>
                  <TableCell>
                    <TableCellLayout>
                      <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                        {BidTypeLabel[bid.cr5ab_bidtypeid.cr5ab_code as BidTypeCode]}
                      </Text>
                    </TableCellLayout>
                  </TableCell>
                  <TableCell>
                    <TableCellLayout>
                      {bid.cr5ab_opportunitystage !== undefined ? (
                        <Badge appearance="filled" color={OpportunityStageColor[bid.cr5ab_opportunitystage]} size="small">
                          {OpportunityStageLabel[bid.cr5ab_opportunitystage]}
                        </Badge>
                      ) : <Text size={200} style={{ color: tokens.colorNeutralForeground4 }}>—</Text>}
                    </TableCellLayout>
                  </TableCell>
                  <TableCell>
                    <TableCellLayout>
                      <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                        {bid.cr5ab_source !== undefined
                          ? bid.cr5ab_sourceportalname
                            ? `${BidSourceLabel[bid.cr5ab_source]} — ${bid.cr5ab_sourceportalname}`
                            : BidSourceLabel[bid.cr5ab_source]
                          : "—"}
                      </Text>
                    </TableCellLayout>
                  </TableCell>
                  <TableCell>
                    <TableCellLayout>
                      <StatusBadge status={bid.cr5ab_status} />
                    </TableCellLayout>
                  </TableCell>
                  <TableCell>
                    <TableCellLayout>
                      <Text size={200} className={styles.deadlineCell}>
                        {formatDate(bid.cr5ab_submissiondeadline)}
                      </Text>
                    </TableCellLayout>
                  </TableCell>
                  <TableCell>
                    <TableCellLayout>
                      <Text size={200} className={styles.valueCell}>
                        {formatCurrency(bid.cr5ab_estimatedvalue)}
                      </Text>
                    </TableCellLayout>
                  </TableCell>
                  <TableCell className={styles.actionCell}>
                    <TableCellLayout>
                      <Button
                        appearance="subtle"
                        icon={<ArrowRightRegular />}
                        size="small"
                        aria-label={`Open ${bid.cr5ab_title}`}
                        onClick={(e) => {
                          e.stopPropagation();
                          navigate(`/workspace?bidId=${bid.id}`);
                        }}
                      />
                    </TableCellLayout>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        )}
      </Card>
    </div>
  );
}
