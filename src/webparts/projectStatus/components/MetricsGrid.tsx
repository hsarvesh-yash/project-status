import * as React from 'react';
import { TextField } from '@fluentui/react';
import styles from './ProjectStatus.module.scss';
import { IProjectItem, IProjectEditableField, IContracts } from '../services/ProjectService';

interface IMetricsGridProps {
  projectDetails: IProjectItem;
  contractDetails: IContracts | null;
  deployedHeadCount: number | null;
  billableHeadCount: number | null;
  projectEditableField: IProjectEditableField;
  onBillingValueChange: (value: string) => void;
}

const formatCurrency = (currency: string, raw: string | number): string => {
  const prefix = currency ? `${currency} ` : '';
  const num = typeof raw === 'number' ? raw : parseFloat(String(raw).replace(/[^0-9.-]+/g, '')) || 0;
  return num > 0 ? `${prefix}${num.toLocaleString()}` : (raw ? `${prefix}${raw}` : 'N/A');
};

export const MetricsGrid: React.FC<IMetricsGridProps> = ({
  projectDetails,
  contractDetails,
  deployedHeadCount,
  billableHeadCount,
  projectEditableField,
  onBillingValueChange,
}) => {
  const tcv = contractDetails?.TotalContractValue
    ?? (parseFloat(String(projectDetails.TotalSOWAmount || 0).replace(/[^0-9.-]+/g, '')) || 0);

  const metrics = [
    { label: 'Project ID', value: projectDetails.ProjectID },
    { label: 'Project Name', value: projectDetails.ProjectName },
    { label: 'Project Type', value: projectDetails.ProjectType },
    { label: 'Planned Start Date', value: projectDetails.PlannedStartDate },
    { label: 'Planned End Date', value: projectDetails.PlannedEndDate },
    { label: 'Project Manager', value: projectDetails.projectManager },
    { label: 'Delivery Manager', value: projectDetails.deliveryManager },
    { label: 'Geographics', value: projectDetails.ProjectGEOS },
    { label: 'Total Contract Value', value: formatCurrency(projectDetails.Currency, tcv) },
    { label: 'Deployed HeadCount', value: deployedHeadCount !== null ? String(deployedHeadCount) : 'N/A' },
    { label: 'Billable HeadCount', value: billableHeadCount !== null ? billableHeadCount.toFixed(2) : 'N/A' },
    { label: 'Currency SOW Amount', value: formatCurrency(projectDetails.Currency, projectDetails.TotalSOWAmount) },
  ];

  return (
    <div className={styles.metricsGrid}>
      {metrics.map(({ label, value }) => (
        <div key={label} className={styles.metricCard}>
          <div className={styles.metricLabel}>{label}</div>
          <div className={styles.metricValue}>{value}</div>
        </div>
      ))}
      <div className={styles.metricCard}>
        <div className={styles.metricLabel}>Monthly Billing Value (Editable)</div>
        <TextField
          value={projectEditableField.BillingValue}
          onChange={(_, newValue) => onBillingValueChange(newValue ?? '')}
          placeholder="Enter monthly billing..."
          prefix={projectDetails.Currency ? `${projectDetails.Currency} ` : ''}
          styles={{
            fieldGroup: { border: '1px solid #cbd5e1', borderRadius: '6px', background: '#ffffff' },
            field: { fontWeight: 600, color: '#0f172a' },
          }}
        />
      </div>
    </div>
  );
};
