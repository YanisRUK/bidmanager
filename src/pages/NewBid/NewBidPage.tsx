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
} from "@fluentui/react-icons";
import { PageHeader } from "../../components/common/PageHeader";
import { useDataverse } from "../../hooks/useDataverse";
import { dataverseClient } from "../../lib/dataverse";
import { useAuth } from "../../context/AuthContext";
import {
  BidTypeCode,
  BidTypeLabel,
  BidStatus,
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
    gridTemplateColumns: "repeat(auto-fill, minmax(240px, 1fr))",
    gap: tokens.spacingVerticalL,
    marginBottom: tokens.spacingVerticalXL,
  },
  typeCard: {
    cursor: "pointer",
    border: `2px solid ${tokens.colorNeutralStroke2}`,
    transition: "border-color 0.15s ease, box-shadow 0.15s ease",
    ":hover": {
      ...shorthands.borderColor(tokens.colorBrandStroke1),
    },
  },
  typeCardSelected: {
    ...shorthands.borderColor(tokens.colorBrandBackground),
    boxShadow: `0 0 0 1px ${tokens.colorBrandBackground}`,
  },
  typeCardInner: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  typeIcon: {
    fontSize: "28px",
    color: tokens.colorBrandForeground1,
    marginBottom: tokens.spacingVerticalXS,
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
    return form.bidTypeCode !== null;
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
      (t) => t.ricoh_code === form.bidTypeCode
    );
    if (!selectedType) return;

    setIsSubmitting(true);
    setSubmitError(null);

    try {
      const newBid = await dataverseClient.createBidRequest({
        ricoh_bidreferencenumber: `BID-${new Date().getFullYear()}-DRAFT`,
        ricoh_title: form.title,
        ricoh_bidtypeid: {
          id: selectedType.id,
          ricoh_name: selectedType.ricoh_name,
          ricoh_code: selectedType.ricoh_code,
        },
        ricoh_status: BidStatus.Submitted,
        ricoh_customername: form.customerName,
        ricoh_customerindustry: form.customerIndustry || undefined,
        ricoh_estimatedvalue: form.estimatedValue
          ? Number(form.estimatedValue)
          : undefined,
        ricoh_currency: "GBP",
        ricoh_submissiondeadline: form.submissionDeadline,
        ricoh_expectedawarddate: form.expectedAwardDate || undefined,
        ricoh_contractduration: form.contractDuration
          ? Number(form.contractDuration)
          : undefined,
        ricoh_description: form.description,
        ricoh_scope: form.scope || undefined,
        ricoh_specialrequirements: form.specialRequirements || undefined,
        ricoh_incumbentvendor: form.incumbentVendor || undefined,
        ricoh_submittedby: user,
        ricoh_routedto: selectedType.ricoh_routingTeam,
      });

      // TODO: trigger Power Automate routing flow
      // await flowClient.triggerRouteBid({ bidRequestId: newBid.id, ... });

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
        <Text size={400} style={{ marginBottom: tokens.spacingVerticalL, display: "block" }}>
          Choose the bid type that best describes this opportunity.
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
                onClick={() => update("bidTypeCode", code)}
                role="radio"
                aria-checked={isSelected}
                tabIndex={0}
                onKeyDown={(e) => e.key === "Enter" && update("bidTypeCode", code)}
              >
                <div className={styles.typeCardInner}>
                  <div className={styles.typeIcon}>{meta.icon}</div>
                  <Text weight="semibold" size={400}>
                    {BidTypeLabel[code]}
                  </Text>
                  <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                    {meta.description}
                  </Text>
                  <Text size={200} style={{ color: tokens.colorBrandForeground1 }}>
                    SLA: {meta.sla}
                  </Text>
                </div>
              </Card>
            );
          })}
        </div>
      </>
    );
  }

  // ------------------------------------------------------------------
  // Step 1: Bid details form
  // ------------------------------------------------------------------

  function renderStep1() {
    return (
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
    );
  }

  // ------------------------------------------------------------------
  // Step 2: Review
  // ------------------------------------------------------------------

  function renderStep2() {
    const rows: [string, string][] = [
      ["Bid Type", form.bidTypeCode !== null ? BidTypeLabel[form.bidTypeCode] : "—"],
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
