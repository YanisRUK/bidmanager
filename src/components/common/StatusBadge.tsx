
import { Badge } from "@fluentui/react-components";
import { BidStatus, BidStatusLabel, BidStatusColor } from "../../types/dataverse";

interface StatusBadgeProps {
  status: BidStatus;
}

export function StatusBadge({ status }: StatusBadgeProps) {
  return (
    <Badge
      appearance="filled"
      color={BidStatusColor[status]}
      size="medium"
    >
      {BidStatusLabel[status]}
    </Badge>
  );
}
