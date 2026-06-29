import * as React from 'react';
import { TextField } from '@fluentui/react';
import styles from './ProjectStatus.module.scss';
import { IProjectItem, IProjectEditableField, IContracts, ISPWSROutputRecord, formatDateDisplay } from '../services/ProjectService';

const cssStyles = styles as unknown as Record<string, string>;

interface IMetricsGridProps {
  projectDetails: IProjectItem;
  contractDetails: IContracts | undefined;
  deployedHeadCount: number | undefined;
  billableHeadCount: number | undefined;
  projectEditableField: IProjectEditableField;
  lastSavedRecord: ISPWSROutputRecord | undefined;
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
  lastSavedRecord,
  onBillingValueChange,
}) => {
  const tcv = (contractDetails?.TotalContractValue && contractDetails.TotalContractValue > 0)
    ? contractDetails.TotalContractValue
    : ((projectDetails.TotalContractValue && projectDetails.TotalContractValue > 0)
      ? projectDetails.TotalContractValue
      : (parseFloat(String(projectDetails.TotalSOWAmount || 0).replace(/[^0-9.-]+/g, '')) || 0));

  const sowAmt = (projectDetails.TotalSOWAmount && parseFloat(String(projectDetails.TotalSOWAmount).replace(/[^0-9.-]+/g, '')) > 0)
    ? projectDetails.TotalSOWAmount
    : tcv;

  const isDiff = (val1: unknown, val2: unknown): boolean => {
    if (lastSavedRecord === undefined || val2 === undefined || val2 === null || val2 === '') return false;
    return String(val1).trim().toLowerCase() !== String(val2).trim().toLowerCase();
  };

  const metrics = [
    { label: 'Project ID', value: projectDetails.ProjectID },
    { label: 'Project Name', value: projectDetails.ProjectName },
    { label: 'Project Type', value: projectDetails.ProjectType, lastSaved: lastSavedRecord?.ProjectType },
    { label: 'Project Currency', value: projectDetails.Currency || 'USD', lastSaved: lastSavedRecord?.Currency },
    { label: 'Planned Start Date', value: projectDetails.PlannedStartDate, lastSaved: formatDateDisplay(lastSavedRecord?.PlannedStartDate) },
    { label: 'Planned End Date', value: projectDetails.PlannedEndDate, lastSaved: formatDateDisplay(lastSavedRecord?.PlannedEndDate) },
    { label: 'Project Manager', value: projectDetails.projectManager, lastSaved: lastSavedRecord?.ProjectManager },
    { label: 'Delivery Manager', value: projectDetails.deliveryManager, lastSaved: lastSavedRecord?.DeliveryManager },
    { label: 'Geographics', value: projectDetails.ProjectGEOS, lastSaved: lastSavedRecord?.ProjectGEOS },
    { 
      label: 'Total Contract Value', 
      value: formatCurrency(projectDetails.Currency, tcv), 
      lastSaved: lastSavedRecord?.TotalContractValue !== undefined ? formatCurrency(projectDetails.Currency, lastSavedRecord.TotalContractValue) : undefined 
    },
    { 
      label: 'Deployed HeadCount', 
      value: deployedHeadCount !== undefined ? String(deployedHeadCount) : 'N/A', 
      lastSaved: lastSavedRecord?.deployedHeadCount !== undefined ? String(lastSavedRecord.deployedHeadCount) : undefined 
    },
    { 
      label: 'Billable HeadCount', 
      value: billableHeadCount !== undefined ? billableHeadCount.toFixed(2) : 'N/A', 
      lastSaved: lastSavedRecord?.BillableHeadCount !== undefined ? Number(lastSavedRecord.BillableHeadCount).toFixed(2) : undefined 
    },
    { 
      label: 'Currency SOW Amount', 
      value: formatCurrency(projectDetails.Currency, sowAmt), 
      lastSaved: lastSavedRecord?.TotalSOWAmount !== undefined ? formatCurrency(projectDetails.Currency, lastSavedRecord.TotalSOWAmount) : undefined 
    },
  ];

  const billingChanged = isDiff(projectEditableField.BillingValue, lastSavedRecord?.BillingValue);

  return (
    <div className={styles.metricsGrid}>
      {metrics.map(({ label, value, lastSaved }) => {
        const changed = isDiff(value, lastSaved);
        return (
          <div key={label} className={`${styles.metricCard} ${changed && cssStyles.metricCardChanged ? cssStyles.metricCardChanged : ''}`}>
            <div className={styles.metricLabel}>{label}</div>
            <div className={styles.metricValue}>{value}</div>
            {changed && lastSaved !== undefined && cssStyles.changedBadge && (
              <div className={cssStyles.changedBadge}>Changed (Last Saved: {String(lastSaved)})</div>
            )}
          </div>
        );
      })}
      <div className={`${styles.metricCard} ${billingChanged && cssStyles.metricCardChanged ? cssStyles.metricCardChanged : ''}`}>
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
        {billingChanged && lastSavedRecord?.BillingValue !== undefined && cssStyles.changedBadge && (
          <div className={cssStyles.changedBadge}>Changed (Last Saved: {String(lastSavedRecord.BillingValue)})</div>
        )}
      </div>
    </div>
  );
};
