/**
 * New Bid — Intake Form
 *
 * Step 1: Select bid type → Step 2: Fill type-specific form → Step 3: Review & Submit
 *
 * On submission the bid is created in Dataverse and a Power Automate routing
 * flow is triggered.
 */

import React, { useState } from "react";
import { useNavigate } from "react-router-dom";
import {
  makeStyles,
  shorthands,
  tokens,
  Text,
  Button,
  Card,
  Input,
  Textarea,
  Field,
  Select,
  Spinner,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import {
  ArrowRightRegular,
  ArrowLeftRegular,
  CheckmarkRegular,
  DocumentRegular,
  PeopleRegular,
  BuildingRegular,
  ClipboardTaskRegular,
  CheckmarkCircleRegular,
  CheckmarkCircleFilled,
  GlobeRegular,
  PersonRegular,
  LightbulbRegular,
  MoreHorizontalRegular,
} from "@fluentui/react-icons";
import { PageHeader } from "../../components/common/PageHeader";
import { useDataverse } from "../../hooks/useDataverse";
import { dataverseClient } from "../../lib/dataverse";
import { useAuth } from "../../context/AuthContext";
import {
  BidTypeCode,
  BidTypeLabel,
  BidStatus,
  BidSource,
  BidSourceLabel,
  OpportunityStage,
} from "../../types/dataverse";

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

const useStyles = makeStyles({
  stepperWrap: {
    marginBottom: tokens.spacingVerticalXL,
  },
  stepper: {
    display: "flex",
    alignItems: "center",
    gap: "0",
  },
  stepItem: {
    display: "flex",
    alignItems: "center",
    flexShrink: 0,
  },
  stepCircle: {
    width: "32px",
    height: "32px",
    borderRadius: "50%",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    flexShrink: 0,
  },
  stepCircleActive: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  stepCircleComplete: {
    backgroundColor: tokens.colorPaletteGreenBackground2,
    color: tokens.colorPaletteGreenForeground1,
  },
  stepCircleDefault: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground3,
  },
  stepLabel: {
    marginLeft: tokens.spacingHorizontalS,
  },
  stepConnector: {
    flexGrow: 1,
    height: "1px",
    backgroundColor: tokens.colorNeutralStroke2,
    ...shorthands.margin("0", tokens.spacingHorizontalM),
    minWidth: "32px",
  },
  typeGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))",
    gap: tokens.spacingVerticalM,
    marginBottom: tokens.spacingVerticalXL,
  },
  typeCard: {
    cursor: "pointer",
    border: `2px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusLarge,
    transition: "border-color 0.12s ease, background-color 0.12s ease, box-shadow 0.12s ease",
    position: "relative",
    ":hover": {
      ...shorthands.borderColor(tokens.colorBrandStroke1),
      boxShadow: tokens.shadow4,
    },
  },
  typeCardSelected: {
    border: "2px solid transparent",
    backgroundImage: "linear-gradient(white, white), linear-gradient(135deg, #0078d4, #50a0f0)",
    backgroundOrigin: "border-box",
    backgroundClip: "padding-box, border-box",
    boxShadow: "0 0 0 2px rgba(0,120,212,0.25)",
    cursor: "default",
    color: "inherit",
  },
  typeCardInner: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    ...shorthands.padding(tokens.spacingVerticalM),
  },
  typeIcon: {
    fontSize: "32px",
    color: tokens.colorBrandForeground1,
    marginBottom: tokens.spacingVerticalXS,
  },
  typeIconSelected: {
    fontSize: "32px",
    color: tokens.colorBrandForeground1,
    marginBottom: tokens.spacingVerticalXS,
  },
  // Tick badge overlaid in top-right of selected card
  selectionTick: {
    position: "absolute",
    top: tokens.spacingVerticalS,
    right: tokens.spacingHorizontalM,
    fontSize: "20px",
    color: "#0078d4",
  },
  sourceGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(170px, 1fr))",
    gap: tokens.spacingVerticalM,
    marginBottom: tokens.spacingVerticalL,
  },
  sourceCard: {
    cursor: "pointer",
    border: `2px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusLarge,
    transition: "border-color 0.12s ease, background-color 0.12s ease, box-shadow 0.12s ease",
    position: "relative",
    ":hover": {
      ...shorthands.borderColor(tokens.colorBrandStroke1),
      boxShadow: tokens.shadow4,
    },
  },
  sourceCardSelected: {
    border: "2px solid transparent",
    backgroundImage: "linear-gradient(white, white), linear-gradient(135deg, #0078d4, #50a0f0)",
    backgroundOrigin: "border-box",
    backgroundClip: "padding-box, border-box",
    boxShadow: "0 0 0 2px rgba(0,120,212,0.25)",
    cursor: "default",
    color: "inherit",
  },
  sourceCardInner: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: tokens.spacingVerticalXS,
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalM),
    textAlign: "center" as const,
  },
  sourceIcon: {
    fontSize: "24px",
    color: tokens.colorBrandForeground1,
  },
  sourceIconSelected: {
    fontSize: "24px",
    color: tokens.colorBrandForeground1,
  },
  sectionDivider: {
    ...shorthands.margin(tokens.spacingVerticalXL, "0", tokens.spacingVerticalL, "0"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  selectedTypeBanner: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalL,
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalL),
    backgroundColor: tokens.colorBrandBackground2,
    borderRadius: tokens.borderRadiusLarge,
    border: `1.5px solid ${tokens.colorBrandStroke1}`,
    marginBottom: tokens.spacingVerticalL,
  },
  selectedTypeBannerIcon: {
    fontSize: "32px",
    color: tokens.colorBrandForeground1,
    flexShrink: 0,
  },
  selectedTypeBannerBody: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
    flexGrow: 1,
  },
  selectedTypeBannerMeta: {
    display: "flex",
    gap: tokens.spacingHorizontalM,
    flexWrap: "wrap",
  },
  selectedTypeBannerChange: {
    marginLeft: "auto",
    flexShrink: 0,
    alignSelf: "center",
  },
  formGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: tokens.spacingVerticalM,
    marginBottom: tokens.spacingVerticalL,
    "@media (max-width: 700px)": {
      gridTemplateColumns: "1fr",
    },
  },
  formFullWidth: {
    gridColumn: "1 / -1",
  },
  reviewSection: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
  },
  reviewRow: {
    display: "flex",
    gap: tokens.spacingHorizontalM,
    ...shorthands.padding(tokens.spacingVerticalS, "0"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ":last-child": { borderBottom: "none" },
  },
  reviewLabel: {
    minWidth: "180px",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  actions: {
    display: "flex",
    justifyContent: "space-between",
    marginTop: tokens.spacingVerticalXL,
  },
});

// ---------------------------------------------------------------------------
// Custom stepper component (StepperField not in Fluent UI v9.x public API)
// ---------------------------------------------------------------------------

const STEPS = ["Select Type", "Bid Details", "Review & Submit"];

function StepIndicator({ currentStep }: { currentStep: number }) {
  const styles = useStyles();
  return (
    <div className={styles.stepper} aria-label="New bid progress">
      {STEPS.map((label, i) => {
        const isComplete = i < currentStep;
        const isActive = i === currentStep;
        return (
          <React.Fragment key={label}>
            {i > 0 && <div className={styles.stepConnector} />}
            <div className={styles.stepItem}>
              <div
                className={`${styles.stepCircle} ${
                  isComplete
                    ? styles.stepCircleComplete
                    : isActive
                    ? styles.stepCircleActive
                    : styles.stepCircleDefault
                }`}
                aria-current={isActive ? "step" : undefined}
              >
                {isComplete ? <CheckmarkCircleRegular style={{ fontSize: "16px" }} /> : i + 1}
              </div>
              <Text
                className={styles.stepLabel}
                size={300}
                weight={isActive ? "semibold" : "regular"}
                style={{
                  color: isActive
                    ? tokens.colorBrandForeground1
                    : isComplete
                    ? tokens.colorNeutralForeground1
                    : tokens.colorNeutralForeground3,
                }}
              >
                {label}
              </Text>
            </div>
          </React.Fragment>
        );
      })}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Bid type metadata for the card selection
// ---------------------------------------------------------------------------

const BID_TYPE_META: Record<
  BidTypeCode,
  { icon: React.ReactNode; description: string; sla: string }
> = {
  [BidTypeCode.SupplierQuestionnaire]: {
    icon: <ClipboardTaskRegular />,
    description: "Pre-qualification questionnaire for supplier registration.",
    sla: "5 business days",
  },
  [BidTypeCode.SalesLed]: {
    icon: <PeopleRegular />,
    description: "RFP or tender driven and owned by the Sales team.",
    sla: "3 business days",
  },
  [BidTypeCode.SMEToQualify]: {
    icon: <BuildingRegular />,
    description: "Opportunity routed to a Subject Matter Expert for qualification.",
    sla: "7 business days",
  },
  [BidTypeCode.BidManagement]: {
    icon: <DocumentRegular />,
    description: "Full lifecycle bid managed by the Bid Management team.",
    sla: "2 business days",
  },
};

// ---------------------------------------------------------------------------
// Form state
// ---------------------------------------------------------------------------

interface BidFormData {
  bidTypeCode: BidTypeCode | null;
  source: BidSource | null;
  sourcePortalName: string;
  title: string;
  customerName: string;
  customerIndustry: string;
  estimatedValue: string;
  submissionDeadline: string;
  expectedAwardDate: string;
  contractDuration: string;
  description: string;
  scope: string;
  specialRequirements: string;
  incumbentVendor: string;
}

const INITIAL_FORM: BidFormData = {
  bidTypeCode: null,
  source: null,
  sourcePortalName: "",
  title: "",
  customerName: "",
  customerIndustry: "",
  estimatedValue: "",
  submissionDeadline: "",
  expectedAwardDate: "",
  contractDuration: "",
  description: "",
  scope: "",
  specialRequirements: "",
  incumbentVendor: "",
};

const SOURCE_META: Record<BidSource, { icon: React.ReactNode; description: string }> = {
  [BidSource.Portal]:         { icon: <GlobeRegular />,           description: "Contracts Finder, Find a Tender, or similar portal" },
  [BidSource.SalesSubmitted]: { icon: <PersonRegular />,          description: "Opportunity sent to us by the Sales team" },
  [BidSource.Proactive]:      { icon: <LightbulbRegular />,       description: "Proactively identified by Ricoh" },
  [BidSource.Other]:          { icon: <MoreHorizontalRegular />,  description: "Another channel not listed above" },
};

const INDUSTRIES = [
  "Government",
  "Healthcare",
  "Education",
  "Retail",
  "Financial Services",
  "Manufacturing",
  "Professional Services",
  "Technology",
  "Other",
];

// ---------------------------------------------------------------------------
// Page
// ---------------------------------------------------------------------------

export function NewBidPage() {
  const styles = useStyles();
  const navigate = useNavigate();
  const { user } = useAuth();

  const [step, setStep] = useState(0);
  const [form, setForm] = useState<BidFormData>(INITIAL_FORM);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submitError, setSubmitError] = useState<string | null>(null);

  const { data: bidTypes } = useDataverse(
    () => dataverseClient.getBidTypes(),
    []
  );

  // ------------------------------------------------------------------
  // Handlers
  // ------------------------------------------------------------------

  function update<K extends keyof BidFormData>(key: K, value: BidFormData[K]) {
    setForm((prev) => ({ ...prev, [key]: value }));
  }

  function canProceedStep0() {
    return form.bidTypeCode !== null && form.source !== null;
  }

  function canProceedStep1() {
    return (
      form.title.trim() !== "" &&
      form.customerName.trim() !== "" &&
      form.submissionDeadline !== "" &&
      form.description.trim() !== ""
    );
  }

  async function handleSubmit() {
    if (!form.bidTypeCode || !user) return;

    const selectedType = bidTypes?.find(
      (t) => t.cr5ab_code === form.bidTypeCode
    );
    if (!selectedType) return;

    setIsSubmitting(true);
    setSubmitError(null);

    try {
      // 1. Create the bid request record
      const newBid = await dataverseClient.createBidRequest({
        cr5ab_bidreferencenumber: `BID-${new Date().getFullYear()}-DRAFT`,
        cr5ab_title: form.title,
        cr5ab_bidtypeid: {
          id: selectedType.id,
          cr5ab_name: selectedType.cr5ab_name,
          cr5ab_code: selectedType.cr5ab_code,
        },
        cr5ab_status: BidStatus.Submitted,
        cr5ab_opportunitystage: OpportunityStage.PendingOAF,
        cr5ab_source: form.source ?? undefined,
        cr5ab_sourceportalname: form.source === BidSource.Portal ? form.sourcePortalName || undefined : undefined,
        cr5ab_customername: form.customerName,
        cr5ab_customerindustry: form.customerIndustry || undefined,
        cr5ab_estimatedvalue: form.estimatedValue
          ? Number(form.estimatedValue)
          : undefined,
        cr5ab_currency: "GBP",
        cr5ab_submissiondeadline: form.submissionDeadline,
        cr5ab_expectedawarddate: form.expectedAwardDate || undefined,
        cr5ab_contractduration: form.contractDuration
          ? Number(form.contractDuration)
          : undefined,
        cr5ab_description: form.description,
        cr5ab_scope: form.scope || undefined,
        cr5ab_specialrequirements: form.specialRequirements || undefined,
        cr5ab_incumbentvendor: form.incumbentVendor || undefined,
        cr5ab_submittedby: user,
        cr5ab_routedto: selectedType.cr5ab_routingteam,
      });

      // 2. For Bid Management type, immediately create a workspace
      if (form.bidTypeCode === BidTypeCode.BidManagement) {
        await dataverseClient.createBidWorkspace({
          cr5ab_title: `${form.title} — Workspace`,
          cr5ab_bidrequestid: {
            id: newBid.id,
            cr5ab_title: newBid.cr5ab_title,
            cr5ab_bidreferencenumber: newBid.cr5ab_bidreferencenumber,
          },
          cr5ab_status: BidStatus.Submitted,
          cr5ab_bidmanagerid: user,
          cr5ab_completionpercentage: 0,
        });
      }

      // 3. Trigger Power Automate routing flow
      await dataverseClient.triggerRouteBidFlow({
        bidRequestId: newBid.id,
        bidTypeCode: selectedType.cr5ab_code,
        submittedById: user.id,
      });

      navigate("/bid-register", { state: { newBidId: newBid.id } });
    } catch (err) {
      setSubmitError(err instanceof Error ? err.message : "Submission failed");
    } finally {
      setIsSubmitting(false);
    }
  }

  // ------------------------------------------------------------------
  // Step 0: Type selection
  // ------------------------------------------------------------------

  function renderStep0() {
    return (
      <>
        {/* ── Section 1: Bid Type ──────────────────────────────────── */}
        <div style={{ display: "flex", alignItems: "baseline", gap: tokens.spacingHorizontalM, marginBottom: tokens.spacingVerticalS }}>
          <div style={{
            width: "24px", height: "24px", borderRadius: "50%",
            backgroundColor: "#0078d4",
            color: "#ffffff",
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightBold,
            flexShrink: 0,
          }}>1</div>
          <Text size={500} weight="semibold">Select bid type</Text>
        </div>
        <Text size={300} style={{ marginBottom: tokens.spacingVerticalL, display: "block", color: tokens.colorNeutralForeground2, paddingLeft: "32px" }}>
          Choose the type that best describes this opportunity.
        </Text>

        <div className={styles.typeGrid}>
          {(
            [
              BidTypeCode.SupplierQuestionnaire,
              BidTypeCode.SalesLed,
              BidTypeCode.SMEToQualify,
              BidTypeCode.BidManagement,
            ] as BidTypeCode[]
          ).map((code) => {
            const meta = BID_TYPE_META[code];
            const isSelected = form.bidTypeCode === code;
            return (
              <Card
                key={code}
                className={`${styles.typeCard} ${isSelected ? styles.typeCardSelected : ""}`}
                onClick={() => !isSelected && update("bidTypeCode", code)}
                role="radio"
                aria-checked={isSelected}
                tabIndex={0}
                onKeyDown={(e) => e.key === "Enter" && update("bidTypeCode", code)}
              >
                {/* Tick badge — top right when selected */}
                {isSelected && (
                  <CheckmarkCircleFilled className={styles.selectionTick} />
                )}
                <div className={styles.typeCardInner}>
                  <div className={isSelected ? styles.typeIconSelected : styles.typeIcon}>
                    {meta.icon}
                  </div>
                  <Text
                    weight="semibold"
                    size={400}
                  >
                    {BidTypeLabel[code]}
                  </Text>
                  <Text
                    size={200}
                    style={{ color: tokens.colorNeutralForeground3 }}
                  >
                    {meta.description}
                  </Text>
                  <Text
                    size={200}
                    weight="semibold"
                    style={{ color: tokens.colorBrandForeground1 }}
                  >
                    SLA: {meta.sla}
                  </Text>
                  {isSelected && (
                    <div style={{
                      marginTop: tokens.spacingVerticalXS,
                      display: "inline-flex", alignItems: "center", gap: "4px",
                      backgroundColor: "rgba(0,120,212,0.1)",
                      borderRadius: tokens.borderRadiusCircular,
                      padding: "2px 10px",
                      fontSize: tokens.fontSizeBase100,
                      color: "#0078d4",
                      fontWeight: tokens.fontWeightSemibold,
                      letterSpacing: "0.04em",
                      textTransform: "uppercase",
                    }}>
                      Selected
                    </div>
                  )}
                </div>
              </Card>
            );
          })}
        </div>

        {/* ── Section divider ──────────────────────────────────────── */}
        <div className={styles.sectionDivider} />

        {/* ── Section 2: Source / Channel ──────────────────────────── */}
        <div style={{ display: "flex", alignItems: "baseline", gap: tokens.spacingHorizontalM, marginBottom: tokens.spacingVerticalS }}>
          <div style={{
            width: "24px", height: "24px", borderRadius: "50%",
            backgroundColor: form.bidTypeCode !== null ? "#0078d4" : tokens.colorNeutralBackground3,
            color: form.bidTypeCode !== null ? "#ffffff" : tokens.colorNeutralForeground3,
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightBold,
            flexShrink: 0,
            transition: "background-color 0.2s ease",
          }}>2</div>
          <Text
            size={500}
            weight="semibold"
            style={{ color: form.bidTypeCode !== null ? undefined : tokens.colorNeutralForeground3 }}
          >
            How did this come in?
          </Text>
        </div>
        <Text size={300} style={{ marginBottom: tokens.spacingVerticalL, display: "block", color: tokens.colorNeutralForeground2, paddingLeft: "32px" }}>
          Tell us where this opportunity originated.
        </Text>

        <div className={styles.sourceGrid} style={{ opacity: form.bidTypeCode !== null ? 1 : 0.4, pointerEvents: form.bidTypeCode !== null ? "auto" : "none" }}>
          {([BidSource.Portal, BidSource.SalesSubmitted, BidSource.Proactive, BidSource.Other] as BidSource[]).map((src) => {
            const meta = SOURCE_META[src];
            const isSelected = form.source === src;
            return (
              <Card
                key={src}
                className={`${styles.sourceCard} ${isSelected ? styles.sourceCardSelected : ""}`}
                onClick={() => !isSelected && update("source", src)}
                role="radio"
                aria-checked={isSelected}
                tabIndex={0}
                onKeyDown={(e) => e.key === "Enter" && update("source", src)}
              >
                {isSelected && (
                  <CheckmarkCircleFilled className={styles.selectionTick} />
                )}
                <div className={styles.sourceCardInner}>
                  <div className={isSelected ? styles.sourceIconSelected : styles.sourceIcon}>
                    {meta.icon}
                  </div>
                  <Text
                    weight="semibold"
                    size={300}
                  >
                    {BidSourceLabel[src]}
                  </Text>
                  <Text
                    size={200}
                    style={{ color: tokens.colorNeutralForeground3 }}
                  >
                    {meta.description}
                  </Text>
                  {isSelected && (
                    <div style={{
                      marginTop: tokens.spacingVerticalXS,
                      display: "inline-flex", alignItems: "center", gap: "4px",
                      backgroundColor: "rgba(0,120,212,0.1)",
                      borderRadius: tokens.borderRadiusCircular,
                      padding: "2px 10px",
                      fontSize: tokens.fontSizeBase100,
                      color: "#0078d4",
                      fontWeight: tokens.fontWeightSemibold,
                      letterSpacing: "0.04em",
                      textTransform: "uppercase",
                    }}>
                      Selected
                    </div>
                  )}
                </div>
              </Card>
            );
          })}
        </div>

        {/* Portal name input — only shows when Portal is selected */}
        {form.source === BidSource.Portal && (
          <div style={{ maxWidth: "400px", marginBottom: tokens.spacingVerticalL, paddingLeft: "32px" }}>
            <Field label="Which portal?" hint="e.g. Contracts Finder, Find a Tender, G-Cloud, DOS">
              <Input
                placeholder="Type the portal name..."
                value={form.sourcePortalName}
                onChange={(_, d) => update("sourcePortalName", d.value)}
                autoFocus
              />
            </Field>
          </div>
        )}
      </>
    );
  }

  // ------------------------------------------------------------------
  // Step 1: Bid details form
  // ------------------------------------------------------------------

  function renderStep1() {
    const typeMeta = form.bidTypeCode !== null ? BID_TYPE_META[form.bidTypeCode] : null;
    return (
      <>
        {/* Selected type + source summary banner */}
        {form.bidTypeCode !== null && typeMeta && (
          <div className={styles.selectedTypeBanner}>
            <div className={styles.selectedTypeBannerIcon}>{typeMeta.icon}</div>
            <div className={styles.selectedTypeBannerBody}>
              <div style={{ display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS, flexWrap: "wrap" }}>
                <Text weight="semibold" size={400} style={{ color: tokens.colorBrandForeground1 }}>
                  {BidTypeLabel[form.bidTypeCode]}
                </Text>
                <CheckmarkCircleFilled style={{ fontSize: "16px", color: tokens.colorBrandForeground1 }} />
              </div>
              <div className={styles.selectedTypeBannerMeta}>
                <Text size={200} style={{ color: tokens.colorNeutralForeground2 }}>
                  SLA: {typeMeta.sla}
                </Text>
                {form.source !== null && (
                  <>
                    <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>·</Text>
                    <Text size={200} style={{ color: tokens.colorNeutralForeground2 }}>
                      {BidSourceLabel[form.source]}
                      {form.source === BidSource.Portal && form.sourcePortalName ? ` — ${form.sourcePortalName}` : ""}
                    </Text>
                    <CheckmarkCircleFilled style={{ fontSize: "14px", color: tokens.colorBrandForeground1 }} />
                  </>
                )}
              </div>
            </div>
            <Button
              appearance="subtle"
              size="small"
              className={styles.selectedTypeBannerChange}
              onClick={() => setStep(0)}
            >
              Change
            </Button>
          </div>
        )}
      <div className={styles.formGrid}>
        <Field label="Bid Title" required className={styles.formFullWidth}>
          <Input
            placeholder="e.g. NHS Digital Printing Framework"
            value={form.title}
            onChange={(_, d) => update("title", d.value)}
          />
        </Field>

        <Field label="Customer / Organisation" required>
          <Input
            placeholder="e.g. NHS England"
            value={form.customerName}
            onChange={(_, d) => update("customerName", d.value)}
          />
        </Field>

        <Field label="Industry">
          <Select
            value={form.customerIndustry}
            onChange={(_, d) => update("customerIndustry", d.value)}
          >
            <option value="">Select industry</option>
            {INDUSTRIES.map((i) => (
              <option key={i} value={i}>
                {i}
              </option>
            ))}
          </Select>
        </Field>

        <Field label="Estimated Value (£)">
          <Input
            type="number"
            placeholder="e.g. 500000"
            value={form.estimatedValue}
            onChange={(_, d) => update("estimatedValue", d.value)}
            contentBefore={<Text size={200}>£</Text>}
          />
        </Field>

        <Field label="Submission Deadline" required>
          <Input
            type="date"
            value={form.submissionDeadline}
            onChange={(_, d) => update("submissionDeadline", d.value)}
          />
        </Field>

        <Field label="Expected Award Date">
          <Input
            type="date"
            value={form.expectedAwardDate}
            onChange={(_, d) => update("expectedAwardDate", d.value)}
          />
        </Field>

        <Field label="Contract Duration (months)">
          <Input
            type="number"
            placeholder="e.g. 36"
            value={form.contractDuration}
            onChange={(_, d) => update("contractDuration", d.value)}
          />
        </Field>

        <Field label="Incumbent Vendor">
          <Input
            placeholder="e.g. Xerox, Konica Minolta"
            value={form.incumbentVendor}
            onChange={(_, d) => update("incumbentVendor", d.value)}
          />
        </Field>

        <Field label="Description" required className={styles.formFullWidth}>
          <Textarea
            placeholder="Describe the opportunity, background and objectives..."
            value={form.description}
            onChange={(_, d) => update("description", d.value)}
            resize="vertical"
            rows={4}
          />
        </Field>

        <Field label="Scope of Work" className={styles.formFullWidth}>
          <Textarea
            placeholder="What products / services are in scope?"
            value={form.scope}
            onChange={(_, d) => update("scope", d.value)}
            resize="vertical"
            rows={3}
          />
        </Field>

        <Field label="Special Requirements" className={styles.formFullWidth}>
          <Textarea
            placeholder="Security clearance, certifications, compliance requirements..."
            value={form.specialRequirements}
            onChange={(_, d) => update("specialRequirements", d.value)}
            resize="vertical"
            rows={3}
          />
        </Field>
      </div>
      </>
    );
  }

  // ------------------------------------------------------------------
  // Step 2: Review
  // ------------------------------------------------------------------

  function renderStep2() {
    const typeMeta = form.bidTypeCode !== null ? BID_TYPE_META[form.bidTypeCode] : null;
    const sourceLabel = form.source !== null
      ? `${BidSourceLabel[form.source]}${form.source === BidSource.Portal && form.sourcePortalName ? ` — ${form.sourcePortalName}` : ""}`
      : "—";
    const rows: [string, string][] = [
      ["Source / Channel", sourceLabel],
      ["Title", form.title],
      ["Customer", form.customerName],
      ["Industry", form.customerIndustry || "—"],
      ["Estimated Value", form.estimatedValue ? `£${Number(form.estimatedValue).toLocaleString()}` : "—"],
      ["Submission Deadline", form.submissionDeadline ? new Date(form.submissionDeadline).toLocaleDateString("en-GB") : "—"],
      ["Expected Award", form.expectedAwardDate ? new Date(form.expectedAwardDate).toLocaleDateString("en-GB") : "—"],
      ["Contract Duration", form.contractDuration ? `${form.contractDuration} months` : "—"],
      ["Incumbent Vendor", form.incumbentVendor || "—"],
      ["Description", form.description],
    ];

    return (
      <div className={styles.reviewSection}>
        <Text size={400}>
          Please review your submission before sending.
        </Text>
        {/* Bid type callout — prominent at the top of the review */}
        {form.bidTypeCode !== null && typeMeta && (
          <div className={styles.selectedTypeBanner}>
            <div className={styles.selectedTypeBannerIcon}>{typeMeta.icon}</div>
            <div className={styles.selectedTypeBannerBody}>
              <Text weight="semibold" size={400} style={{ color: tokens.colorBrandForeground1 }}>
                {BidTypeLabel[form.bidTypeCode]}
              </Text>
              <div className={styles.selectedTypeBannerMeta}>
                <Text size={200} style={{ color: tokens.colorNeutralForeground2 }}>
                  {typeMeta.description}
                </Text>
                <Text size={200} style={{ color: tokens.colorBrandForeground1, fontWeight: tokens.fontWeightSemibold }}>
                  SLA: {typeMeta.sla}
                </Text>
              </div>
            </div>
          </div>
        )}
        <Card>
          {rows.map(([label, value]) => (
            <div key={label} className={styles.reviewRow}>
              <Text className={styles.reviewLabel}>{label}</Text>
              <Text size={300}>{value}</Text>
            </div>
          ))}
        </Card>
        {submitError && (
          <MessageBar intent="error">
            <MessageBarBody>{submitError}</MessageBarBody>
          </MessageBar>
        )}
      </div>
    );
  }

  // ------------------------------------------------------------------
  // Render
  // ------------------------------------------------------------------

  return (
    <div>
      <PageHeader
        title="New Bid"
        subtitle="Submit a new bid or opportunity for review"
      />

      {/* Stepper */}
      <div className={styles.stepperWrap}>
        <StepIndicator currentStep={step} />
      </div>

      {/* Step content */}
      {step === 0 && renderStep0()}
      {step === 1 && renderStep1()}
      {step === 2 && renderStep2()}

      {/* Navigation actions */}
      <div className={styles.actions}>
        <Button
          appearance="subtle"
          icon={<ArrowLeftRegular />}
          disabled={step === 0 || isSubmitting}
          onClick={() => setStep((s) => s - 1)}
        >
          Back
        </Button>

        {step < 2 ? (
          <Button
            appearance="primary"
            icon={<ArrowRightRegular />}
            iconPosition="after"
            disabled={step === 0 ? !canProceedStep0() : !canProceedStep1()}
            onClick={() => setStep((s) => s + 1)}
          >
            Next
          </Button>
        ) : (
          <Button
            appearance="primary"
            icon={isSubmitting ? <Spinner size="tiny" /> : <CheckmarkRegular />}
            disabled={isSubmitting}
            onClick={handleSubmit}
          >
            {isSubmitting ? "Submitting..." : "Submit Bid"}
          </Button>
        )}
      </div>
    </div>
  );
}
