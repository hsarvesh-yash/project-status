import * as React from 'react';
import styles from './ProjectStatus.module.scss';

interface ISegmentedControlProps {
  value: string;
  onChange: (value: string) => void;
}

const SEGMENTS = ['green', 'amber', 'red'] as const;

export const SegmentedControl: React.FC<ISegmentedControlProps> = ({ value, onChange }) => (
  <div className={styles.segmentedControl}>
    {SEGMENTS.map(segment => (
      <button
        key={segment}
        type="button"
        className={`${styles.segmentBtn} ${styles[segment]} ${value === segment ? styles.active : ''}`}
        onClick={() => onChange(segment)}
      >
        <span className={styles.statusDot} />
        {segment.charAt(0).toUpperCase() + segment.slice(1)}
      </button>
    ))}
  </div>
);
