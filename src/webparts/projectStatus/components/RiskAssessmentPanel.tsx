import * as React from 'react';
import { TextField } from '@fluentui/react';
import styles from './ProjectStatus.module.scss';
import { IRiskAssessment } from '../services/ProjectService';
import { SegmentedControl } from './SegmentedControl';

interface IRiskField {
  key: keyof Pick<IRiskAssessment, 'overallStatus' | 'scope' | 'schedule' | 'deliverableQuality' | 'resourceAvailability'>;
  label: string;
  commentKey: keyof IRiskAssessment;
  commentLabel: string;
}

const RISK_FIELDS: IRiskField[] = [
  { key: 'overallStatus', label: 'Overall Status', commentKey: 'overallStatusComment', commentLabel: 'Please provide details for non-green overall status' },
  { key: 'scope', label: 'Scope', commentKey: 'scopeComment', commentLabel: 'Please provide details for scope risk' },
  { key: 'schedule', label: 'Schedule', commentKey: 'scheduleComment', commentLabel: 'Please provide details for schedule variance' },
  { key: 'deliverableQuality', label: 'Deliverable Quality', commentKey: 'deliverableQualityComment', commentLabel: 'Please provide details for quality issues' },
  { key: 'resourceAvailability', label: 'Resource Availability', commentKey: 'resourceAvailabilityComment', commentLabel: 'Please provide details for staffing/resource gaps' },
];

interface IRiskAssessmentPanelProps {
  riskAssessment: IRiskAssessment;
  onChange: (field: keyof IRiskAssessment, value: string) => void;
}

export const RiskAssessmentPanel: React.FC<IRiskAssessmentPanelProps> = ({ riskAssessment, onChange }) => (
  <div className={styles.riskGrid}>
    {RISK_FIELDS.map(({ key, label, commentKey, commentLabel }) => (
      <div key={key} className={styles.riskRowCard}>
        <div className={styles.riskHeaderBar}>
          <div className={styles.riskTitle}>{label}</div>
          <div className={styles.riskSelector}>
            <SegmentedControl value={riskAssessment[key]} onChange={val => onChange(key, val)} />
          </div>
        </div>
        {(riskAssessment[key] === 'amber' || riskAssessment[key] === 'red') && (
          <div className={styles.commentContainer}>
            <TextField
              label={commentLabel}
              value={riskAssessment[commentKey] as string}
              onChange={(_, newValue) => onChange(commentKey, newValue ?? '')}
              multiline
              rows={2}
            />
          </div>
        )}
      </div>
    ))}
  </div>
);
