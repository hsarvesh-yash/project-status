import { useState, useEffect, useRef } from 'react';
import { IDropdownOption, MessageBarType } from '@fluentui/react';
import { sp } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IProjectItem,
  IProjectEditableField,
  IRiskAssessment,
  IContracts,
  IWSROutputItem,
  ISPWSROutputRecord,
  getProjectList,
  getProjectDetails,
  getResourceAllocation,
  getContractData,
  checkTodayWSR,
  getAllWSRForProject,
  saveWSROutput,
} from '../services/ProjectService';
import { exportToPpt } from '../services/PptExportService';
import { getCurrentDate, calculateMonthlyBilling } from '../utils/dateUtils';

const DEFAULT_RISK: IRiskAssessment = {
  overallStatus: 'green',
  scope: 'green',
  schedule: 'green',
  deliverableQuality: 'green',
  resourceAvailability: 'green',
  overallStatusComment: '',
  scopeComment: '',
  scheduleComment: '',
  deliverableQualityComment: '',
  resourceAvailabilityComment: '',
};

const DEFAULT_EDITABLE: IProjectEditableField = {
  KeyUpdates: '',
  ProjectDescription: '',
  NextStepUpdates: '',
  BillingValue: '',
};

// eslint-disable-next-line @typescript-eslint/explicit-function-return-type
export const useProjectForm = (context: WebPartContext, backgroundImageUrl: string) => {
  const [listTypeOptions, setListTypeOptions] = useState<IDropdownOption[]>([]);
  const [selectedProject, setSelectedProject] = useState<string | undefined>(undefined);
  const [projectDetails, setProjectDetails] = useState<IProjectItem | undefined>(undefined);
  const [contractDetails, setContractDetails] = useState<IContracts | undefined>(undefined);
  const [deployedHeadCount, setDeployedHeadCount] = useState<number | undefined>(undefined);
  const [billableHeadCount, setBillableHeadCount] = useState<number | undefined>(undefined);
  const [riskAssessment, setRiskAssessment] = useState<IRiskAssessment>({ ...DEFAULT_RISK });
  const [projectEditableField, setProjectEditableField] = useState<IProjectEditableField>({ ...DEFAULT_EDITABLE });
  const [saveMessage, setSaveMessage] = useState<{ type: MessageBarType; text: string } | undefined>(undefined);
  const [isLoading, setIsLoading] = useState(false);
  const spInitialized = useRef(false);

  const showMessage = (type: MessageBarType, text: string, duration = 5000): void => {
    setSaveMessage({ type, text });
    setTimeout(() => setSaveMessage(undefined), duration);
  };

  const dismissMessage = (): void => setSaveMessage(undefined);

  const applyWSRData = (data: ISPWSROutputRecord): void => {
    setRiskAssessment({
      overallStatus: data.overallStatus || 'green',
      scope: data.scope || 'green',
      schedule: data.schedule || 'green',
      deliverableQuality: data.deliverableQuality || 'green',
      resourceAvailability: data.resourceAvailability || 'green',
      overallStatusComment: data.overallStatusComment || '',
      scopeComment: data.scopeComment || '',
      scheduleComment: data.scheduleComment || '',
      deliverableQualityComment: data.deliverableQualityComment || '',
      resourceAvailabilityComment: data.resourceAvailabilityComment || '',
    });
    setProjectEditableField({
      KeyUpdates: data.KeyUpdates || '',
      ProjectDescription: data.ProjectDescription || '',
      NextStepUpdates: data.NextStepUpdates || '',
      BillingValue: data.BillingValue ? String(data.BillingValue) : '',
    });
    if (data.deployedHeadCount !== undefined) setDeployedHeadCount(data.deployedHeadCount);
    if (data.BillableHeadCount !== undefined) setBillableHeadCount(data.BillableHeadCount);
    if (data.TotalContractValue !== undefined) {
      setContractDetails({ TotalContractValue: data.TotalContractValue, BillingValue: data.BillingValue || 0 });
    }
  };

  // Initialize PnP SP once and load project list
  useEffect(() => {
    if (!spInitialized.current && context?.pageContext) {
      sp.setup({ sp: { baseUrl: context.pageContext.web.absoluteUrl } });
      spInitialized.current = true;
    }
    const loadProjects = async (): Promise<void> => {
      try {
        const options = await getProjectList(context.pageContext.user.email);
        setListTypeOptions(options);
      } catch (error) {
        console.error('Error loading project list:', error);
        showMessage(MessageBarType.error, 'Error loading projects. Please refresh.', 8000);
      }
    };
    loadProjects().catch((err: unknown) => console.error('Error loading projects:', err));
  }, []);

  // Load all project data when selection changes — sequential to fix race condition
  useEffect(() => {
    if (!selectedProject) return;

    const loadProjectData = async (): Promise<void> => {
      setIsLoading(true);
      try {
        // Step 1: project details first (its name is needed to build the WSR report title)
        const details = await getProjectDetails(selectedProject);
        setProjectDetails(details);

        // Step 2: contracts + resources in parallel
        const [contracts, resources] = await Promise.all([
          getContractData(selectedProject),
          getResourceAllocation(selectedProject),
        ]);
        setContractDetails(contracts);
        setDeployedHeadCount(resources.deployed);
        setBillableHeadCount(resources.billable);

        // Step 3: check today's WSR (requires project name from step 1)
        if (details) {
          const reportTitle = `${details.ProjectName} - WSR - ${getCurrentDate()}`;
          const todayWSR = await checkTodayWSR(selectedProject, reportTitle);
          if (todayWSR) {
            applyWSRData(todayWSR);
            showMessage(MessageBarType.info, "Existing WSR data found and loaded for today's date!");
          } else {
            setRiskAssessment({ ...DEFAULT_RISK });
            setProjectEditableField({ ...DEFAULT_EDITABLE });
          }
        }
      } catch (error) {
        console.error('Error loading project data:', error);
        showMessage(MessageBarType.error, `Error loading project data: ${(error as Error).message}`);
      } finally {
        setIsLoading(false);
      }
    };

    loadProjectData().catch((err: unknown) => console.error('Error loading project data:', err));
  }, [selectedProject]);

  // Auto-populate billing when project or contract data arrives
  useEffect(() => {
    if (!projectDetails) return;
    setProjectEditableField(prev => {
      if (prev.BillingValue) return prev;
      const tcv = contractDetails?.TotalContractValue
        ?? (parseFloat(String(projectDetails.TotalSOWAmount || 0).replace(/[^0-9.-]+/g, '')) || 0);
      const contractBilling = contractDetails?.BillingValue || 0;
      const autoBilling = contractBilling > 0
        ? String(contractBilling)
        : calculateMonthlyBilling(tcv, projectDetails.PlannedStartDate, projectDetails.PlannedEndDate);
      return { ...prev, BillingValue: autoBilling };
    });
  }, [projectDetails, contractDetails]);

  const handleSelectProject = (projectId: string): void => setSelectedProject(projectId);

  const handleRiskChange = (field: keyof IRiskAssessment, value: string): void =>
    setRiskAssessment(prev => ({ ...prev, [field]: value }));

  const handleEditableFieldChange = (field: keyof IProjectEditableField, value: string): void =>
    setProjectEditableField(prev => ({ ...prev, [field]: value }));

  const handleLoadMostRecent = async (): Promise<void> => {
    if (!selectedProject) {
      showMessage(MessageBarType.error, 'Please select a project first!', 3000);
      return;
    }
    try {
      const allData = await getAllWSRForProject(selectedProject);
      if (allData.length > 0) {
        applyWSRData(allData[0]);
        const createdDate = new Date(allData[0].Created).toLocaleDateString();
        showMessage(MessageBarType.success, `Most recent WSR data loaded (from ${createdDate})!`);
      } else {
        showMessage(MessageBarType.info, 'No existing WSR data found for this project.');
      }
    } catch (error) {
      console.error('Error loading most recent WSR:', error);
      showMessage(MessageBarType.error, 'Error loading existing WSR data.');
    }
  };

  const handleClearForm = async (): Promise<void> => {
    setRiskAssessment({ ...DEFAULT_RISK });
    setProjectEditableField({ ...DEFAULT_EDITABLE });
    if (selectedProject) {
      try {
        const resources = await getResourceAllocation(selectedProject);
        setDeployedHeadCount(resources.deployed);
        setBillableHeadCount(resources.billable);
      } catch (error) {
        console.error('Error refreshing resource data:', error);
      }
    }
    showMessage(MessageBarType.info, 'Form reset with fresh data!', 3000);
  };

  const saveToWSROutput = async (): Promise<boolean> => {
    if (!projectDetails) return false;
    try {
      const numBilling = parseFloat(projectEditableField.BillingValue.replace(/[^0-9.-]+/g, ''))
        || (contractDetails?.BillingValue || 0);
      const wsrData: Partial<IWSROutputItem> = {
        Title: `${projectDetails.ProjectName} - WSR - ${getCurrentDate()}`,
        ProjectID: projectDetails.ProjectID,
        ProjectName: projectDetails.ProjectName,
        ProjectType: projectDetails.ProjectType,
        PlannedStartDate: projectDetails.PlannedStartDate,
        PlannedEndDate: projectDetails.PlannedEndDate,
        ProjectManager: projectDetails.projectManager,
        DeliveryManager: projectDetails.deliveryManager,
        TotalSOWAmount: projectDetails.TotalSOWAmount,
        Currency: projectDetails.Currency,
        ProjectGEOS: projectDetails.ProjectGEOS,
        QualityEngineer: projectDetails.QualityEngineer,
        deployedHeadCount: deployedHeadCount || 0,
        BillableHeadCount: billableHeadCount || 0,
        ...riskAssessment,
        KeyUpdates: projectEditableField.KeyUpdates,
        ProjectDescription: projectEditableField.ProjectDescription,
        NextStepUpdates: projectEditableField.NextStepUpdates,
        TotalContractValue: contractDetails?.TotalContractValue || 0,
        BillingValue: numBilling,
      };
      await saveWSROutput(wsrData);
      return true;
    } catch (error) {
      console.error('Error saving WSR output:', error);
      return false;
    }
  };

  const handleExportToPPT = async (): Promise<void> => {
    if (!projectDetails) {
      showMessage(MessageBarType.error, 'Please select a project before exporting');
      return;
    }
    const numBilling = parseFloat(projectEditableField.BillingValue.replace(/[^0-9.-]+/g, '')) || 0;
    if (!numBilling || numBilling <= 0) {
      showMessage(MessageBarType.error, 'Please enter a valid Monthly Billing Value before exporting to PowerPoint!', 6000);
      return;
    }
    try {
      const saveSuccess = await saveToWSROutput();
      await exportToPpt({
        projectDetails,
        projectEditableField,
        riskAssessment,
        contractDetails,
        deployedHeadCount,
        billableHeadCount,
        numBilling,
        backgroundImageUrl,
      });
      showMessage(
        saveSuccess ? MessageBarType.success : MessageBarType.warning,
        saveSuccess
          ? 'PowerPoint exported successfully and data saved to WSROutput list!'
          : 'PowerPoint exported successfully, but there was an issue saving to WSROutput list.'
      );
    } catch (error) {
      console.error('Error exporting to PowerPoint:', error);
      showMessage(MessageBarType.error, `Error exporting to PowerPoint: ${(error as Error).message}`);
    }
  };

  return {
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
  };
};
