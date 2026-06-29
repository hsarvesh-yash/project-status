import * as React from 'react';
import {
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  MessageBar,
  TextField,
  Icon,
  Spinner,
  SpinnerSize,
} from '@fluentui/react';
import styles from './ProjectStatus.module.scss';
import type { IProjectStatusProps } from './IProjectStatusProps';
import { MetricsGrid } from './MetricsGrid';
import { RiskAssessmentPanel } from './RiskAssessmentPanel';
import { useProjectForm } from '../hooks/useProjectForm';

export default function ProjectStatus(props: IProjectStatusProps): React.ReactElement {
  const {
    listTypeOptions,
    selectedProject,
    projectDetails,
    contractDetails,
    deployedHeadCount,
    billableHeadCount,
    riskAssessment,
    projectEditableField,
    saveMessage,
    isLoading,
    dismissMessage,
    handleSelectProject,
    handleRiskChange,
    handleEditableFieldChange,
    handleLoadMostRecent,
    handleClearForm,
    handleExportToPPT,
  } = useProjectForm(props.context, props.backgroundImageUrl);

  return (
    <section className={`${styles.projectStatus} ${props.hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.mainContainer}>

        <div className={styles.headerHero}>
          <div className={styles.headerTitleGroup}>
            <h2>Project Status &amp; PowerPoint Generator</h2>
            <p>Executive WSR tracking, RAG status risk assessment, and slide deck exporter</p>
          </div>
        </div>

        {saveMessage && (
          <MessageBar
            messageBarType={saveMessage.type}
            isMultiline={false}
            onDismiss={dismissMessage}
            dismissButtonAriaLabel="Close"
          >
            {saveMessage.text}
          </MessageBar>
        )}

        <div className={styles.controlsCard}>
          <div className={styles.dropdownGroup}>
            <span className={styles.dropdownLabel}>Select Active Project</span>
            <Dropdown
              selectedKey={selectedProject}
              options={listTypeOptions}
              onChange={(_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
                if (option) handleSelectProject(option.key as string);
              }}
              placeholder={listTypeOptions.length === 0 ? 'No projects assigned to you' : 'Choose a project to manage...'}
              disabled={listTypeOptions.length === 0}
              styles={{ dropdown: { width: 320 } }}
            />
          </div>
          {selectedProject && (
            <div className={styles.actionButtonsGroup}>
              <PrimaryButton
                text="Load Most Recent WSR"
                onClick={() => { void handleLoadMostRecent(); }}
                iconProps={{ iconName: 'Download' }}
              />
              <PrimaryButton
                text="Clear Form"
                onClick={() => { void handleClearForm(); }}
                iconProps={{ iconName: 'Refresh' }}
              />
            </div>
          )}
        </div>

        {isLoading && <Spinner size={SpinnerSize.large} label="Loading project data..." />}

        {!isLoading && projectDetails && (
          <>
            <div className={styles.detailsSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="ViewDashboard" style={{ fontSize: 20, color: '#0f172a' }} />
                <h3>Project Overview &amp; Key Metrics</h3>
              </div>
              <MetricsGrid
                projectDetails={projectDetails}
                contractDetails={contractDetails}
                deployedHeadCount={deployedHeadCount}
                billableHeadCount={billableHeadCount}
                projectEditableField={projectEditableField}
                onBillingValueChange={value => handleEditableFieldChange('BillingValue', value)}
              />
              <div className={styles.remarksSection}>
                <div className={styles.remarkCard}>
                  <div className={styles.remarkLabel}>Project Description</div>
                  <TextField
                    value={projectEditableField.ProjectDescription}
                    onChange={(_, v) => handleEditableFieldChange('ProjectDescription', v ?? '')}
                    multiline
                    rows={3}
                    placeholder="Enter project scope & background..."
                  />
                </div>
                <div className={styles.remarksGridTwoCol}>
                  <div className={styles.remarkCard}>
                    <div className={styles.remarkLabel}>Key Updates (Remarks)</div>
                    <TextField
                      value={projectEditableField.KeyUpdates}
                      onChange={(_, v) => handleEditableFieldChange('KeyUpdates', v ?? '')}
                      multiline
                      rows={3}
                      placeholder="Enter key status updates..."
                    />
                  </div>
                  <div className={styles.remarkCard}>
                    <div className={styles.remarkLabel}>Next Steps / Other Updates</div>
                    <TextField
                      value={projectEditableField.NextStepUpdates}
                      onChange={(_, v) => handleEditableFieldChange('NextStepUpdates', v ?? '')}
                      multiline
                      rows={3}
                      placeholder="Enter planned upcoming milestones..."
                    />
                  </div>
                </div>
              </div>
            </div>

            <div className={styles.riskSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="Warning" style={{ fontSize: 20, color: '#0f172a' }} />
                <h3>Project RAG Status Assessment</h3>
              </div>
              <RiskAssessmentPanel riskAssessment={riskAssessment} onChange={handleRiskChange} />
              <div className={styles.ctaBar}>
                <PrimaryButton
                  text="Export to PowerPoint"
                  iconProps={{ iconName: 'PowerPointDocument' }}
                  onClick={() => { void handleExportToPPT(); }}
                />
              </div>
            </div>
          </>
        )}
      </div>
    </section>
  );
}
