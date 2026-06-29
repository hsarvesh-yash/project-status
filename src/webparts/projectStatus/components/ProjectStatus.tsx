import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './ProjectStatus.module.scss';
import type { IProjectStatusProps } from './IProjectStatusProps';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import pptxgen from 'pptxgenjs';
import { Dropdown, IDropdownOption, PrimaryButton, TextField, MessageBar, MessageBarType, Icon } from '@fluentui/react';

interface IProjectItem {
  ProjectID: string;
  ProjectName: string;
  ProjectType: string;
  PlannedStartDate: string;
  PlannedEndDate: string;
  projectManager: string;
  deliveryManager: string;
  TotalSOWAmount: string;
  Currency: string;
  ProjectGEOS: string;
  QualityEngineer: string;
}

interface IProjcetEditableField {
  KeyUpdates: string;
  ProjectDescription: string;
  NextStepUpdates: string;
  BillingValue: string;
}

interface IRiskAssessment {
  overallStatus: string;
  scope: string;
  schedule: string;
  deliverableQuality: string;
  resourceAvailability: string;
  overallStatusComment: string;
  scopeComment: string;
  scheduleComment: string;
  deliverableQualityComment: string;
  resourceAvailabilityComment: string;
}

interface IContracts {
  TotalContractValue: number;
  BillingValue: number;
}

interface IWSROutputItem {
  Title: string;
  ProjectID: string;
  ProjectName: string;
  ProjectType: string;
  PlannedStartDate: string;
  PlannedEndDate: string;
  ProjectManager: string;
  DeliveryManager: string;
  TotalSOWAmount: string;
  Currency: string;
  ProjectGEOS: string;
  QualityEngineer: string;
  deployedHeadCount: number;
  BillableHeadCount: number;
  overallStatus: string;
  scope: string;
  schedule: string;
  deliverableQuality: string;
  resourceAvailability: string;
  overallStatusComment: string;
  scopeComment: string;
  scheduleComment: string;
  deliverableQualityComment: string;
  resourceAvailabilityComment: string;
  KeyUpdates: string;
  ProjectDescription: string;
  NextStepUpdates: string;
  TotalContractValue: number;
  BillingValue: number;
}

interface ISPProjectItem {
  ProjectID: string;
  CustomerName?: string;
  ProjectName: string;
  ProjectType?: string;
  PlannedStartDate?: string;
  PlannedEndDate?: string;
  TotalSOWAmount: string;
  Currency: string;
  ProjectGEOS: string;
  ProjectManager?: { Title?: string; EMail?: string };
  DeliveryManager?: { Title?: string; EMail?: string };
  QualityEngineer?: { Title?: string; EMail?: string };
}

interface ISPResourceMapping {
  Id: number;
  ResourceId: string;
  ProjectId: string;
  AllocationID?: string;
}

interface ISPResourceAllocation {
  FTEValue?: number | string;
  Allocation?: string;
  AllocationID?: string;
  Billable?: boolean;
}

interface ISPContractItem extends Record<string, unknown> {
  TotalContractValue?: number;
  ContractValue?: number;
  TotalValue?: number;
  Amount?: number;
  BillingValue?: number;
  BilledValue?: number;
  BillingAmount?: number;
  InvoicedAmount?: number;
  TotalInvoiced?: number;
  InvoiceAmount?: number;
  TotalBilled?: number;
  ProjectID?: unknown;
  ProjectId?: unknown;
}

interface ISPWSROutputRecord {
  Created: string;
  overallStatus?: string;
  scope?: string;
  schedule?: string;
  deliverableQuality?: string;
  resourceAvailability?: string;
  overallStatusComment?: string;
  scopeComment?: string;
  scheduleComment?: string;
  deliverableQualityComment?: string;
  resourceAvailabilityComment?: string;
  KeyUpdates?: string;
  ProjectDescription?: string;
  NextStepUpdates?: string;
  deployedHeadCount?: number;
  BillableHeadCount?: number;
  TotalContractValue?: number;
  BillingValue?: number;
}

export default function ProjectStatus(props: IProjectStatusProps): React.ReactElement<IProjectStatusProps> {
  const [selectedProject, setSelectedProject] = useState<string | undefined>(undefined);
  const [listTypeOptions, setListTypeOptions] = useState<IDropdownOption[]>([]);
  const [projectDetails, setProjectDetails] = useState<IProjectItem | null>(null);
  const [contractDetails, setcontractDetails] = useState<IContracts | null>(null);
  const [deployedHeadCount, setdeployedHeadCount] = useState<number | null>(null);
  const [billableHeadCount, setbillableHeadCount] = useState<number | null>(null);
  const [saveMessage, setSaveMessage] = useState<{ type: MessageBarType; text: string } | null>(null);

  const [riskAssessment, setRiskAssessment] = useState<IRiskAssessment>({
    overallStatus: 'green',
    scope: 'green',
    schedule: 'green',
    deliverableQuality: 'green',
    resourceAvailability: 'green',
    overallStatusComment: '',
    scopeComment: '',
    scheduleComment: '',
    deliverableQualityComment: '',
    resourceAvailabilityComment: ''
  });
  const [projectEditableField, setprojectEditableField] = useState<IProjcetEditableField>({
    KeyUpdates: '',
    ProjectDescription: '',
    NextStepUpdates: '',
    BillingValue: '',
  });

  const calculateMonthlyBilling = (tcv: number, startDateStr?: string, endDateStr?: string): string => {
    if (!tcv || tcv <= 0 || !startDateStr || !endDateStr) return '';
    try {
      const start = new Date(startDateStr);
      const end = new Date(endDateStr);
      if (isNaN(start.getTime()) || isNaN(end.getTime()) || end <= start) return '';

      let months = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth());
      if (end.getDate() >= start.getDate()) months += 1;
      if (months <= 0) months = 1;

      const monthly = tcv / months;
      return monthly.toFixed(2);
    } catch {
      return '';
    }
  };

  const executeAsync = (promise: Promise<unknown>): void => {
    promise.catch((err: unknown) => {
      console.error(err);
    });
  };

  const getCurrentDate = (): string => {
    const date = new Date();
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return `${date.getDate()}-${months[date.getMonth()]}-${date.getFullYear()}`;
  };

  const getOrdinalSuffix = (day: number): string => {
    if (day > 3 && day < 21) return 'th';
    switch (day % 10) {
      case 1: return 'st';
      case 2: return 'nd';
      case 3: return 'rd';
      default: return 'th';
    }
  };

  const formatDate = (dateString: string | undefined): string => {
    if (!dateString) return '';
    try {
      const date = new Date(dateString);
      return date.toLocaleDateString('en-GB', {
        day: '2-digit',
        month: 'short',
        year: 'numeric'
      }).replace(/ /g, '-');
    } catch {
      return dateString;
    }
  };

  const shouldShowComment = (status: string): boolean => {
    return status === 'amber' || status === 'red';
  };

  const getRiskColor = (status: string): string => {
    switch (status.toLowerCase()) {
      case 'red':
        return 'EF4444';
      case 'amber':
        return 'F59E0B';
      case 'green':
        return '10B981';
      default:
        return '10B981';
    }
  };

  const GetMaindata = async (): Promise<void> => {
    try {
      if (!props.context || !props.context.pageContext) return;
      const user = props.context.pageContext.user.email;
      const userEmailLower = user ? user.toLowerCase() : '';

      const projectDetailsData: ISPProjectItem[] = await sp.web.lists
        .getByTitle("IMSProjects")
        .items
        .select('*,ProjectID,CustomerName,ProjectName,ProjectType,TotalSOWAmount,Currency,AdditionalPmOrDMId,AdditionalPmOrDM/Title,AdditionalPmOrDM/EMail,DeliveryManagerId,DeliveryManager/Title,DeliveryManager/EMail,ProjectManagerId,ProjectManager/Title,ProjectManager/EMail,QualityEngineerId,QualityEngineer/Title,QualityLeadId,QualityLead/Title,QualityManagerId,QualityManager/Title,QualityLead/EMail,QualityManager/EMail,QualityEngineer/EMail,AccountManager/Title,AccountManager/EMail,ProjectGEOS')
        .expand("AdditionalPmOrDM,DeliveryManager,ProjectManager,QualityEngineer,QualityLead,QualityManager,AccountManager")
        .filter(`(ProjectManager/EMail eq '${userEmailLower}' or DeliveryManager/EMail eq '${userEmailLower}' or AdditionalPmOrDM/EMail eq '${userEmailLower}' or AccountManager/EMail eq '${userEmailLower}' or QualityLead/EMail eq '${userEmailLower}' or QualityManager/EMail eq '${userEmailLower}' or QualityEngineer/EMail eq '${userEmailLower}') and IsProcess eq 'Yes'`)();

      const options: IDropdownOption[] = projectDetailsData.map(item => ({
        key: item.ProjectID,
        text: item.ProjectName
      }));

      setListTypeOptions(options);
    } catch (error) {
      console.error("Error in GetMaindata:", error);
    }
  };

  const getResourceAllocationData = async (projectId: string): Promise<void> => {
    try {
      const projectMappings: ISPResourceMapping[] = await sp.web.lists
        .getByTitle("ProjectResourceMapping")
        .items
        .top(4999)
        .select('*,Id,ResourceId,ProjectId,AllocationID')
        .filter(`ProjectId eq '${projectId}'`)();

      if (projectMappings.length > 0) {
        const allocationIds: string[] = [];
        projectMappings.forEach(item => {
          if (item.AllocationID && allocationIds.indexOf(item.AllocationID) === -1) {
            allocationIds.push(item.AllocationID);
          }
        });

        if (allocationIds.length > 0) {
          const allocationFilters = allocationIds.map(id => `AllocationID eq '${id}'`).join(' or ');
          
          const resourceAllocationData: ISPResourceAllocation[] = await sp.web.lists
            .getByTitle("ResourceAllocation")
            .items
            .select('FTEValue,Allocation,AllocationID,Billable')
            .filter(allocationFilters)
            .top(5000)();

          setdeployedHeadCount(resourceAllocationData.length);

          const billableCount = resourceAllocationData.reduce((sum, item) => {
            const fteValue = Number(item.FTEValue) || 0;
            return sum + fteValue;
          }, 0);

          setbillableHeadCount(billableCount);
        } else {
          setdeployedHeadCount(0);
          setbillableHeadCount(0);
        }
      } else {
        setdeployedHeadCount(0);
        setbillableHeadCount(0);
      }
    } catch (error) {
      console.error("Error fetching resource allocation data:", error);
      setdeployedHeadCount(0);
      setbillableHeadCount(0);
    }
  };

  const getTotalContractValueData = async (projectId: string): Promise<void> => {
    try {
      console.log(`[getTotalContractValueData] Fetching contract details for ProjectID: '${projectId}'`);
      let contractItems: ISPContractItem[] = [];

      // Strategy 1: Direct field filter (ProjectID eq '...' or ProjectId eq '...')
      try {
        contractItems = await sp.web.lists.getByTitle("Contracts").items
          .top(5000)
          .filter(`ProjectID eq '${projectId}' or ProjectId eq '${projectId}'`)();
      } catch (e1) {
        console.warn("[getTotalContractValueData] Strategy 1 (Direct field filter) failed:", e1);
      }

      // Strategy 2: Lookup expand query (ProjectID/ProjectID or ProjectID/Title)
      if (!contractItems || contractItems.length === 0) {
        try {
          contractItems = await sp.web.lists.getByTitle("Contracts").items
            .top(5000)
            .select('*,ProjectID/ProjectID,ProjectID/Title,TotalContractValue')
            .expand('ProjectID')
            .filter(`ProjectID/ProjectID eq '${projectId}' or ProjectID/Title eq '${projectId}'`)();
        } catch (e2) {
          console.warn("[getTotalContractValueData] Strategy 2 (Lookup expand) failed:", e2);
        }
      }

      // Strategy 3: Fallback query all items and filter in-memory
      if (!contractItems || contractItems.length === 0) {
        try {
          const allContracts: ISPContractItem[] = await sp.web.lists.getByTitle("Contracts").items
            .top(2000)();
          contractItems = allContracts.filter(item => {
            const itemPid = (item.ProjectID as Record<string, unknown>)?.ProjectID || (item.ProjectID as Record<string, unknown>)?.Title || item.ProjectID || item.ProjectId;
            return itemPid && String(itemPid).trim().toLowerCase() === String(projectId).trim().toLowerCase();
          });
        } catch (e3) {
          console.warn("[getTotalContractValueData] Strategy 3 (In-memory fallback) failed:", e3);
        }
      }

      console.log("[getTotalContractValueData] Final matched contract items:", contractItems);

      if (contractItems && contractItems.length > 0) {
        const item = contractItems[0];
        const rawVal = item.TotalContractValue ?? item.ContractValue ?? item.TotalValue ?? item.Amount ?? 0;
        const numericVal = typeof rawVal === 'number' ? rawVal : parseFloat(String(rawVal).replace(/[^0-9.-]+/g, '')) || 0;
        
        const rawBilling = item.BillingValue ?? item.BilledValue ?? item.BillingAmount ?? item.InvoicedAmount ?? item.TotalInvoiced ?? item.InvoiceAmount ?? item.TotalBilled ?? 0;
        const numericBilling = typeof rawBilling === 'number' ? rawBilling : parseFloat(String(rawBilling).replace(/[^0-9.-]+/g, '')) || 0;

        const Contracts: IContracts = {
          TotalContractValue: numericVal,
          BillingValue: numericBilling
        };
        setcontractDetails(Contracts);

        // Auto-populate monthly billing input field if currently empty
        setprojectEditableField(prev => {
          if (!prev.BillingValue) {
            const autoBilling = numericBilling > 0 
              ? numericBilling.toString() 
              : calculateMonthlyBilling(numericVal, projectDetails?.PlannedStartDate, projectDetails?.PlannedEndDate);
            return { ...prev, BillingValue: autoBilling };
          }
          return prev;
        });
      } else {
        console.log("[getTotalContractValueData] No matching contract found for ProjectID:", projectId);
        setcontractDetails(null);
      }
    } catch (error) {
      console.error("[getTotalContractValueData] Critical error fetching Contract data:", error);
      setcontractDetails(null);
    }
  };

  const getProjectDetails = async (projectId: string): Promise<void> => {
    try {
      const projectData: ISPProjectItem[] = await sp.web.lists
        .getByTitle("IMSProjects")
        .items
        .filter(`ProjectID eq '${projectId}' and IsProcess eq 'Yes'`)
        .select('*,ProjectID,CustomerName,ProjectName,ProjectType,PlannedStartDate,PlannedEndDate,DeliveryManager/Title,ProjectManager/Title,Currency,TotalSOWAmount,ProjectGEOS,QualityEngineer/Title')
        .expand("DeliveryManager,ProjectManager,QualityEngineer")();
      if (projectData && projectData.length > 0) {
        const itemObj = projectData[0] as unknown as Record<string, unknown>;
        const rawSow = itemObj.TotalSOWAmount ?? itemObj.SOWAmount ?? itemObj.TotalSOW ?? itemObj.SOWValue ?? itemObj.ContractValue ?? '';
        const projectItem: IProjectItem = {
          ProjectID: projectData[0].ProjectID,
          ProjectName: projectData[0].ProjectName,
          ProjectType: projectData[0].ProjectType || 'Not specified',
          PlannedStartDate: formatDate(projectData[0].PlannedStartDate) || 'Not specified',
          PlannedEndDate: formatDate(projectData[0].PlannedEndDate) || 'Not specified',
          projectManager: projectData[0].ProjectManager?.Title || '',
          deliveryManager: projectData[0].DeliveryManager?.Title || '',
          TotalSOWAmount: String(rawSow),
          Currency: projectData[0].Currency || '',
          ProjectGEOS: projectData[0].ProjectGEOS || '',
          QualityEngineer: projectData[0].QualityEngineer?.Title || '',
        };

        setProjectDetails(projectItem);
      } else {
        setProjectDetails(null);
      }
    } catch (error) {
      console.error("Error fetching project details:", error);
      setProjectDetails(null);
    }
  };

  const checkAndFetchWSROutputData = async (projectId: string): Promise<boolean> => {
    try {
      const reportTitle = `${projectDetails?.ProjectName || ''} - WSR - ${getCurrentDate()}`;

      const existingItems: ISPWSROutputRecord[] = await sp.web.lists
        .getByTitle("WSROutput")
        .items
        .filter(`ProjectID eq '${projectId}' and Title eq '${reportTitle}'`)
        .select('*')
        .get();

      if (existingItems.length > 0) {
        const existingData = existingItems[0];

        setRiskAssessment({
          overallStatus: existingData.overallStatus || 'green',
          scope: existingData.scope || 'green',
          schedule: existingData.schedule || 'green',
          deliverableQuality: existingData.deliverableQuality || 'green',
          resourceAvailability: existingData.resourceAvailability || 'green',
          overallStatusComment: existingData.overallStatusComment || '',
          scopeComment: existingData.scopeComment || '',
          scheduleComment: existingData.scheduleComment || '',
          deliverableQualityComment: existingData.deliverableQualityComment || '',
          resourceAvailabilityComment: existingData.resourceAvailabilityComment || ''
        });

        setprojectEditableField({
          KeyUpdates: existingData.KeyUpdates || '',
          ProjectDescription: existingData.ProjectDescription || '',
          NextStepUpdates: existingData.NextStepUpdates || '',
          BillingValue: existingData.BillingValue ? existingData.BillingValue.toString() : ''
        });

        if (existingData.deployedHeadCount !== null && existingData.deployedHeadCount !== undefined) {
          setdeployedHeadCount(existingData.deployedHeadCount);
        }
        if (existingData.BillableHeadCount !== null && existingData.BillableHeadCount !== undefined) {
          setbillableHeadCount(existingData.BillableHeadCount);
        }
        if (existingData.TotalContractValue !== null && existingData.TotalContractValue !== undefined) {
          setcontractDetails({
            TotalContractValue: existingData.TotalContractValue,
            BillingValue: existingData.BillingValue || 0
          });
        }

        setSaveMessage({
          type: MessageBarType.info,
          text: "Existing WSR data found and loaded for today's date!"
        });
        setTimeout(() => setSaveMessage(null), 5000);

        return true;
      } else {
        return false;
      }
    } catch (error) {
      console.error("Error checking WSROutput data:", error);
      return false;
    }
  };

  const checkAllExistingWSROutputData = async (projectId: string): Promise<ISPWSROutputRecord[]> => {
    try {
      const existingItems: ISPWSROutputRecord[] = await sp.web.lists
        .getByTitle("WSROutput")
        .items
        .filter(`ProjectID eq '${projectId}'`)
        .select('*')
        .orderBy('Created', false)
        .get();

      return existingItems;
    } catch (error) {
      console.error("Error checking all WSROutput data:", error);
      return [];
    }
  };

  const loadMostRecentWSRData = async (): Promise<void> => {
    if (!selectedProject) {
      setSaveMessage({
        type: MessageBarType.error,
        text: "Please select a project first!"
      });
      setTimeout(() => setSaveMessage(null), 3000);
      return;
    }

    try {
      const allExistingData = await checkAllExistingWSROutputData(selectedProject);

      if (allExistingData.length > 0) {
        const mostRecentData = allExistingData[0];

        setRiskAssessment({
          overallStatus: mostRecentData.overallStatus || 'green',
          scope: mostRecentData.scope || 'green',
          schedule: mostRecentData.schedule || 'green',
          deliverableQuality: mostRecentData.deliverableQuality || 'green',
          resourceAvailability: mostRecentData.resourceAvailability || 'green',
          overallStatusComment: mostRecentData.overallStatusComment || '',
          scopeComment: mostRecentData.scopeComment || '',
          scheduleComment: mostRecentData.scheduleComment || '',
          deliverableQualityComment: mostRecentData.deliverableQualityComment || '',
          resourceAvailabilityComment: mostRecentData.resourceAvailabilityComment || ''
        });

        setprojectEditableField({
          KeyUpdates: mostRecentData.KeyUpdates || '',
          ProjectDescription: mostRecentData.ProjectDescription || '',
          NextStepUpdates: mostRecentData.NextStepUpdates || '',
          BillingValue: mostRecentData.BillingValue ? mostRecentData.BillingValue.toString() : ''
        });

        if (mostRecentData.deployedHeadCount !== null && mostRecentData.deployedHeadCount !== undefined) {
          setdeployedHeadCount(mostRecentData.deployedHeadCount);
        }
        if (mostRecentData.BillableHeadCount !== null && mostRecentData.BillableHeadCount !== undefined) {
          setbillableHeadCount(mostRecentData.BillableHeadCount);
        }
        if (mostRecentData.TotalContractValue !== null && mostRecentData.TotalContractValue !== undefined) {
          setcontractDetails({
            TotalContractValue: mostRecentData.TotalContractValue,
            BillingValue: mostRecentData.BillingValue || 0
          });
        }

        const createdDate = new Date(mostRecentData.Created).toLocaleDateString();
        setSaveMessage({
          type: MessageBarType.success,
          text: `Most recent WSR data loaded (from ${createdDate})!`
        });
        setTimeout(() => setSaveMessage(null), 5000);
      } else {
        setSaveMessage({
          type: MessageBarType.info,
          text: "No existing WSR data found for this project."
        });
        setTimeout(() => setSaveMessage(null), 5000);
      }
    } catch (error) {
      console.error("Error loading most recent WSR data:", error);
      setSaveMessage({
        type: MessageBarType.error,
        text: "Error loading existing WSR data."
      });
      setTimeout(() => setSaveMessage(null), 5000);
    }
  };

  const saveToWSROutput = async (): Promise<boolean> => {
    if (!projectDetails) {
      return false;
    }

    try {
      const numBilling = parseFloat(projectEditableField.BillingValue.replace(/[^0-9.-]+/g, '')) || (contractDetails?.BillingValue || 0);
      const reportTitle = `${projectDetails.ProjectName} - WSR - ${getCurrentDate()}`;
      const wsrData: Partial<IWSROutputItem> = {
        Title: reportTitle,
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
        overallStatus: riskAssessment.overallStatus,
        scope: riskAssessment.scope,
        schedule: riskAssessment.schedule,
        deliverableQuality: riskAssessment.deliverableQuality,
        resourceAvailability: riskAssessment.resourceAvailability,
        overallStatusComment: riskAssessment.overallStatusComment,
        scopeComment: riskAssessment.scopeComment,
        scheduleComment: riskAssessment.scheduleComment,
        deliverableQualityComment: riskAssessment.deliverableQualityComment,
        resourceAvailabilityComment: riskAssessment.resourceAvailabilityComment,
        KeyUpdates: projectEditableField.KeyUpdates,
        ProjectDescription: projectEditableField.ProjectDescription,
        NextStepUpdates: projectEditableField.NextStepUpdates,
        TotalContractValue: (contractDetails && contractDetails.TotalContractValue) || 0,
        BillingValue: numBilling,
      };

      await sp.web.lists.getByTitle("WSROutput").items.add(wsrData);

      setSaveMessage({
        type: MessageBarType.success,
        text: "Data successfully saved to WSROutput list!"
      });

      setTimeout(() => setSaveMessage(null), 5000);
      return true;
    } catch (error) {
      console.error("Error saving to WSROutput list:", error);
      setSaveMessage({
        type: MessageBarType.error,
        text: `Error saving to WSROutput list: ${(error as Error).message}`
      });
      setTimeout(() => setSaveMessage(null), 5000);
      return false;
    }
  };

  const handleExportToPPT = async (): Promise<void> => {
    if (!projectDetails) {
      setSaveMessage({ type: MessageBarType.error, text: "Please select a project before exporting" });
      setTimeout(() => setSaveMessage(null), 5000);
      return;
    }

    const numBilling = parseFloat(projectEditableField.BillingValue.replace(/[^0-9.-]+/g, '')) || 0;
    if (!numBilling || numBilling <= 0) {
      setSaveMessage({
        type: MessageBarType.error,
        text: "Please enter a valid Monthly Billing Value on the form before exporting to PowerPoint!"
      });
      setTimeout(() => setSaveMessage(null), 6000);
      return;
    }

    try {
      const saveSuccess = await saveToWSROutput();

      const pptx = new pptxgen();
      pptx.layout = 'LAYOUT_16x9';

      const currentDate = new Date();
      const formattedDate = `${currentDate.getDate()}${getOrdinalSuffix(currentDate.getDate())} ${['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][currentDate.getMonth()]
        }'${String(currentDate.getFullYear()).substring(2)}`;

      const slide1 = pptx.addSlide();
      slide1.background = {
        path: "https://ytpl.sharepoint.com/:i:/r/sites/imsyashsharepoint/Shared%20Documents/SPFX%20Icons/DataTemplatePPT.png"
      };

      slide1.addText(`Project Owner: ${projectDetails.projectManager}`, {
        x: 0.4,
        y: 0.28,
        w: 5.0,
        h: 0.3,
        color: '0F172A',
        fontSize: 15,
        bold: true
      });

      slide1.addText(`Updated as on ${formattedDate}`, {
        x: 5.5,
        y: 0.28,
        w: 3.8,
        h: 0.3,
        color: '059669',
        fontSize: 12,
        bold: true,
        align: 'right'
      });

      // Accent vertical bar
      slide1.addShape(pptx.ShapeType.rect, {
        x: 0.2,
        y: 0.72,
        w: 0.06,
        h: 0.9,
        fill: { color: '0073CF' },
        line: { color: '0073CF', width: 0 }
      });

      // Project Description Banner Box
      slide1.addShape(pptx.ShapeType.rect, {
        x: 0.26,
        y: 0.72,
        w: 9.34,
        h: 0.9,
        fill: { color: 'F8FAFC' },
        line: { color: 'CBD5E1', width: 0.5 }
      });

      slide1.addText([
        { text: "Project Description: ", options: { bold: true, color: '0F172A', fontSize: 11 } },
        { text: projectEditableField.ProjectDescription || 'No description provided.', options: { color: '334155', fontSize: 10.5 } }
      ], {
        x: 0.38,
        y: 0.76,
        w: 9.1,
        h: 0.8,
        valign: 'top',
        wrap: true
      });

      // Left Section Header
      slide1.addText("Week Ending Status Report: " + getCurrentDate(), {
        x: 0.2,
        y: 1.72,
        w: 3.4,
        h: 0.3,
        fill: { color: '0F172A' },
        color: 'FFFFFF',
        fontSize: 11,
        bold: true,
        valign: 'middle',
        align: 'center'
      });

      const tableData = [
        [
          { text: 'Type', options: { fill: { color: '1E293B' }, color: 'FFFFFF', bold: true, valign: 'middle' } },
          { text: projectDetails.ProjectType, options: { fill: { color: 'F8FAFC' }, color: '1E293B', valign: 'middle' } }
        ],
        [
          { text: 'Start Date', options: { fill: { color: '1E293B' }, color: 'FFFFFF', bold: true, valign: 'middle' } },
          { text: projectDetails.PlannedStartDate, options: { fill: { color: 'FFFFFF' }, color: '1E293B', valign: 'middle' } }
        ],
        [
          { text: 'End Date', options: { fill: { color: '1E293B' }, color: 'FFFFFF', bold: true, valign: 'middle' } },
          { text: projectDetails.PlannedEndDate, options: { fill: { color: 'FFFFFF' }, color: '1E293B', valign: 'middle' } }
        ],
        [
          { text: 'Overall Status', options: { fill: { color: '1E293B' }, color: 'FFFFFF', bold: true, valign: 'middle' } },
          {
            text: riskAssessment.overallStatus.toUpperCase(),
            options: {
              fill: { color: getRiskColor(riskAssessment.overallStatus) },
              color: 'FFFFFF',
              bold: true,
              align: 'center',
              valign: 'middle'
            }
          }
        ],
        [
          { text: 'Scope', options: { fill: { color: '1E293B' }, color: 'FFFFFF', bold: true, valign: 'middle' } },
          {
            text: riskAssessment.scope.toUpperCase(),
            options: {
              fill: { color: getRiskColor(riskAssessment.scope) },
              color: 'FFFFFF',
              bold: true,
              align: 'center',
              valign: 'middle'
            }
          }
        ],
        [
          { text: 'Schedule', options: { fill: { color: '1E293B' }, color: 'FFFFFF', bold: true, valign: 'middle' } },
          {
            text: riskAssessment.schedule.toUpperCase(),
            options: {
              fill: { color: getRiskColor(riskAssessment.schedule) },
              color: 'FFFFFF',
              bold: true,
              align: 'center',
              valign: 'middle'
            }
          }
        ],
        [
          { text: 'Deliverable Quality', options: { fill: { color: '1E293B' }, color: 'FFFFFF', bold: true, valign: 'middle' } },
          {
            text: riskAssessment.deliverableQuality.toUpperCase(),
            options: {
              fill: { color: getRiskColor(riskAssessment.deliverableQuality) },
              color: 'FFFFFF',
              bold: true,
              align: 'center',
              valign: 'middle'
            }
          }
        ],
        [
          { text: 'Resource Availability', options: { fill: { color: '1E293B' }, color: 'FFFFFF', bold: true, valign: 'middle' } },
          {
            text: riskAssessment.resourceAvailability.toUpperCase(),
            options: {
              fill: { color: getRiskColor(riskAssessment.resourceAvailability) },
              color: 'FFFFFF',
              bold: true,
              align: 'center',
              valign: 'middle'
            }
          }
        ],
        [
          { text: 'Billable Head Count', options: { fill: { color: '1E293B' }, color: 'FFFFFF', bold: true, valign: 'middle' } },
          { text: billableHeadCount !== null ? billableHeadCount.toFixed(2) : 'N/A', options: { fill: { color: 'F8FAFC' }, color: '1E293B', valign: 'middle' } }
        ],
        [
          { text: 'Deployed Head Count', options: { fill: { color: '1E293B' }, color: 'FFFFFF', bold: true, valign: 'middle' } },
          { text: deployedHeadCount !== null ? deployedHeadCount.toString() : 'N/A', options: { fill: { color: 'FFFFFF' }, color: '1E293B', valign: 'middle' } }
        ],
        [
          { text: 'SOW Value', options: { fill: { color: '1E293B' }, color: 'FFFFFF', bold: true, valign: 'middle' } },
          { text: ((): string => {
              const sowNum = parseFloat(String(projectDetails.TotalSOWAmount || 0).replace(/[^0-9.-]+/g, '')) || 0;
              return sowNum > 0 ? `${projectDetails.Currency ? projectDetails.Currency + ' ' : ''}${sowNum.toLocaleString()}` : (projectDetails.TotalSOWAmount ? `${projectDetails.Currency ? projectDetails.Currency + ' ' : ''}${projectDetails.TotalSOWAmount}` : 'N/A');
            })(),
            options: { fill: { color: 'F8FAFC' }, color: '1E293B', bold: true, valign: 'middle' }
          }
        ],
        [
          { text: 'Billing Value', options: { fill: { color: '1E293B' }, color: 'FFFFFF', bold: true, valign: 'middle' } },
          { text: `${projectDetails.Currency ? projectDetails.Currency + ' ' : ''}${numBilling.toLocaleString()}`,
            options: { fill: { color: 'FFFFFF' }, color: '1E293B', bold: true, valign: 'middle' }
          }
        ],
        [
          { text: 'Total Contract Value', options: { fill: { color: '1E293B' }, color: 'FFFFFF', bold: true, valign: 'middle' } },
          { text: ((): string => {
              const val = (contractDetails && contractDetails.TotalContractValue) ? contractDetails.TotalContractValue : (parseFloat(String(projectDetails.TotalSOWAmount || 0).replace(/[^0-9.-]+/g, '')) || 0);
              return `${projectDetails.Currency ? projectDetails.Currency + ' ' : ''}${val.toLocaleString()}`;
            })(),
            options: { fill: { color: 'F8FAFC' }, color: '1E293B', bold: true, valign: 'middle' }
          }
        ],
      ];

      slide1.addTable(tableData as unknown as Parameters<typeof slide1.addTable>[0], {
        x: 0.2,
        y: 2.05,
        w: 3.4,
        colW: [1.4, 2.0],
        border: { type: 'solid', color: 'CBD5E1', pt: 0.5 },
        fontSize: 9.5
      });

      const trimText = (text: string): string => {
        return text.trim().replace(/\s+/g, ' ');
      };

      const riskFields = ['overallStatus', 'scope', 'schedule', 'deliverableQuality', 'resourceAvailability'] as const;
      const riskComments = riskFields
        .filter(field => riskAssessment[field] === 'red' || riskAssessment[field] === 'amber')
        .map(field => {
          const commentField = `${field}Comment` as keyof IRiskAssessment;
          const comment = riskAssessment[commentField];
          if (comment) {
            return `${field.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase())}: ${trimText(comment)}`;
          }
          return null;
        })
        .filter(comment => comment !== null);

      const riskText = riskComments.length > 0 ? riskComments.join('\n') : 'No Major Updates.';

      const formatKeyUpdates = (updates: string): string[] => {
        return updates.split('\n').map(update => trimText(update)).filter(update => update !== '');
      };

      const keyUpdates = formatKeyUpdates(projectEditableField.KeyUpdates);
      const keyUpdatesBullets = keyUpdates.length > 0 ? keyUpdates.join('\n') : '';

      const nextSteps = formatKeyUpdates(projectEditableField.NextStepUpdates);
      const nextStepsBullets = nextSteps.length > 0 ? nextSteps.join('\n') : '';

      const additionalInfoTable = [
        [
          { text: 'Description for Red & Amber Status', options: { fill: { color: '0F172A' }, color: 'FFFFFF', bold: true, valign: 'middle', align: 'left' } },
          { text: trimText(riskText), options: { fill: { color: 'F8FAFC' }, color: '334155', valign: 'top', align: 'left' } }
        ],
        [
          { text: 'Key Updates', options: { fill: { color: '0F172A' }, color: 'FFFFFF', bold: true, valign: 'middle', align: 'left' } },
          { text: keyUpdatesBullets || 'No major updates.', options: { fill: { color: 'FFFFFF' }, color: '334155', valign: 'top', align: 'left' } }
        ],
        [
          { text: 'Next Steps / Other Updates', options: { fill: { color: '0F172A' }, color: 'FFFFFF', bold: true, valign: 'middle', align: 'left' } },
          { text: nextStepsBullets || 'No upcoming updates.', options: { fill: { color: 'F8FAFC' }, color: '334155', valign: 'top', align: 'left' } }
        ],
        [
          { text: 'Risks / Challenges', options: { fill: { color: '0F172A' }, color: 'FFFFFF', bold: true, valign: 'middle', align: 'left' } },
          { text: riskText !== 'No Major Updates.' ? riskText : 'No Major Risks Identified.', options: { fill: { color: 'FFFFFF' }, color: '334155', valign: 'top', align: 'left' } }
        ]
      ];

      slide1.addTable(additionalInfoTable as unknown as Parameters<typeof slide1.addTable>[0], {
        x: 3.8,
        y: 1.72,
        w: 5.8,
        h: 3.6,
        colW: [1.6, 4.2],
        border: { type: 'solid', color: 'CBD5E1', pt: 0.5 },
        fontSize: 10
      });

      await pptx.writeFile({
        fileName: `${projectDetails.ProjectName} Status Report ${getCurrentDate()}`
      });

      if (saveSuccess) {
        setSaveMessage({ type: MessageBarType.success, text: "PowerPoint exported successfully and data saved to WSROutput list!" });
      } else {
        setSaveMessage({ type: MessageBarType.warning, text: "PowerPoint exported successfully, but there was an issue saving to WSROutput list." });
      }

      setTimeout(() => setSaveMessage(null), 5000);
    } catch (error) {
      console.error("Error exporting to PowerPoint:", error);
      setSaveMessage({ type: MessageBarType.error, text: `Error exporting to PowerPoint: ${(error as Error).message}` });
      setTimeout(() => setSaveMessage(null), 5000);
    }
  };

  useEffect(() => {
    if (props.context && props.context.pageContext) {
      sp.setup({
        sp: {
          baseUrl: props.context.pageContext.web.absoluteUrl,
        }
      });
    }

    executeAsync(GetMaindata());
  }, []);

  useEffect(() => {
    if (selectedProject) {
      executeAsync(getProjectDetails(selectedProject));
      executeAsync(getTotalContractValueData(selectedProject));
      executeAsync(getResourceAllocationData(selectedProject));

      executeAsync(checkAndFetchWSROutputData(selectedProject).then(dataFound => {
        if (!dataFound) {
          setRiskAssessment({
            overallStatus: 'green',
            scope: 'green',
            schedule: 'green',
            deliverableQuality: 'green',
            resourceAvailability: 'green',
            overallStatusComment: '',
            scopeComment: '',
            scheduleComment: '',
            deliverableQualityComment: '',
            resourceAvailabilityComment: ''
          });

          setprojectEditableField({
            KeyUpdates: '',
            ProjectDescription: '',
            NextStepUpdates: '',
            BillingValue: ''
          });
        }
      }));
    }
  }, [selectedProject]);

  useEffect(() => {
    if (projectDetails) {
      setprojectEditableField(prev => {
        if (!prev.BillingValue) {
          const tcv = (contractDetails && contractDetails.TotalContractValue) 
            ? contractDetails.TotalContractValue 
            : (parseFloat(String(projectDetails.TotalSOWAmount || 0).replace(/[^0-9.-]+/g, '')) || 0);
          const contractBilling = contractDetails?.BillingValue || 0;
          const autoBilling = contractBilling > 0 
            ? contractBilling.toString() 
            : calculateMonthlyBilling(tcv, projectDetails.PlannedStartDate, projectDetails.PlannedEndDate);
          return { ...prev, BillingValue: autoBilling };
        }
        return prev;
      });
    }
  }, [projectDetails, contractDetails]);

  const handleListTypeChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      setSelectedProject(option.key as string);
    }
  };

  const handleRiskChange = (field: keyof IRiskAssessment, val: string): void => {
    setRiskAssessment(prev => ({
      ...prev,
      [field]: val
    }));
  };

  const renderSegmentedControl = (field: keyof IRiskAssessment): React.ReactElement => {
    const currentVal = riskAssessment[field];
    return (
      <div className={styles.segmentedControl}>
        <button
          type="button"
          className={`${styles.segmentBtn} ${styles.green} ${currentVal === 'green' ? styles.active : ''}`}
          onClick={() => handleRiskChange(field, 'green')}
        >
          <span className={styles.statusDot} />
          Green
        </button>
        <button
          type="button"
          className={`${styles.segmentBtn} ${styles.amber} ${currentVal === 'amber' ? styles.active : ''}`}
          onClick={() => handleRiskChange(field, 'amber')}
        >
          <span className={styles.statusDot} />
          Amber
        </button>
        <button
          type="button"
          className={`${styles.segmentBtn} ${styles.red} ${currentVal === 'red' ? styles.active : ''}`}
          onClick={() => handleRiskChange(field, 'red')}
        >
          <span className={styles.statusDot} />
          Red
        </button>
      </div>
    );
  };

  const handleCommentChange = (field: keyof IRiskAssessment, ev?: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    if (newValue !== undefined) {
      setRiskAssessment({
        ...riskAssessment,
        [field]: newValue
      });
    }
  };

  const handleRemarksChange = (field: keyof IProjcetEditableField, ev?: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    if (newValue !== undefined) {
      setprojectEditableField({
        ...projectEditableField,
        [field]: newValue
      });
    }
  };

  return (
    <section className={`${styles.projectStatus} ${props.hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.mainContainer}>
        {/* Hero Banner */}
        <div className={styles.headerHero}>
          <div className={styles.headerTitleGroup}>
            <h2>Project Status & PowerPoint Generator</h2>
            <p>Executive WSR tracking, RAG status risk assessment, and slide deck exporter</p>
          </div>
        </div>

        {/* System Messages */}
        {saveMessage && (
          <MessageBar
            messageBarType={saveMessage.type}
            isMultiline={false}
            onDismiss={() => setSaveMessage(null)}
            dismissButtonAriaLabel="Close"
          >
            {saveMessage.text}
          </MessageBar>
        )}

        {/* Control Toolbar */}
        <div className={styles.controlsCard}>
          <div className={styles.dropdownGroup}>
            <span className={styles.dropdownLabel}>Select Active Project</span>
            <Dropdown
              selectedKey={selectedProject}
              options={listTypeOptions}
              onChange={handleListTypeChange}
              placeholder="Choose a project to manage..."
              styles={{ dropdown: { width: 320 } }}
            />
          </div>
          {selectedProject && (
            <div className={styles.actionButtonsGroup}>
              <PrimaryButton
                text="Load Most Recent WSR"
                onClick={() => { executeAsync(loadMostRecentWSRData()); }}
                iconProps={{ iconName: 'Download' }}
              />
              <PrimaryButton
                text="Clear Form"
                onClick={() => {
                  setRiskAssessment({
                    overallStatus: 'green',
                    scope: 'green',
                    schedule: 'green',
                    deliverableQuality: 'green',
                    resourceAvailability: 'green',
                    overallStatusComment: '',
                    scopeComment: '',
                    scheduleComment: '',
                    deliverableQualityComment: '',
                    resourceAvailabilityComment: ''
                  });

                  setprojectEditableField({
                    KeyUpdates: '',
                    ProjectDescription: '',
                    NextStepUpdates: '',
                    BillingValue: ''
                  });

                  executeAsync(getResourceAllocationData(selectedProject));

                  setSaveMessage({
                    type: MessageBarType.info,
                    text: "Form reset with fresh data!"
                  });
                  setTimeout(() => setSaveMessage(null), 3000);
                }}
                iconProps={{ iconName: 'Refresh' }}
              />
            </div>
          )}
        </div>

        {/* Project Details Section */}
        {projectDetails && (
          <>
            <div className={styles.detailsSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="ViewDashboard" style={{ fontSize: 20, color: '#0f172a' }} />
                <h3>Project Overview & Key Metrics</h3>
              </div>

              <div className={styles.metricsGrid}>
                <div className={styles.metricCard}>
                  <div className={styles.metricLabel}>Project ID</div>
                  <div className={styles.metricValue}>{projectDetails.ProjectID}</div>
                </div>
                <div className={styles.metricCard}>
                  <div className={styles.metricLabel}>Project Name</div>
                  <div className={styles.metricValue}>{projectDetails.ProjectName}</div>
                </div>
                <div className={styles.metricCard}>
                  <div className={styles.metricLabel}>Project Type</div>
                  <div className={styles.metricValue}>{projectDetails.ProjectType}</div>
                </div>
                <div className={styles.metricCard}>
                  <div className={styles.metricLabel}>Planned Start Date</div>
                  <div className={styles.metricValue}>{projectDetails.PlannedStartDate}</div>
                </div>
                <div className={styles.metricCard}>
                  <div className={styles.metricLabel}>Planned End Date</div>
                  <div className={styles.metricValue}>{projectDetails.PlannedEndDate}</div>
                </div>
                <div className={styles.metricCard}>
                  <div className={styles.metricLabel}>Project Manager</div>
                  <div className={styles.metricValue}>{projectDetails.projectManager}</div>
                </div>
                <div className={styles.metricCard}>
                  <div className={styles.metricLabel}>Delivery Manager</div>
                  <div className={styles.metricValue}>{projectDetails.deliveryManager}</div>
                </div>
                <div className={styles.metricCard}>
                  <div className={styles.metricLabel}>Geographics</div>
                  <div className={styles.metricValue}>{projectDetails.ProjectGEOS}</div>
                </div>
                <div className={styles.metricCard}>
                  <div className={styles.metricLabel}>Total Contract Value</div>
                  <div className={styles.metricValue}>{((): string => {
                    const val = (contractDetails && contractDetails.TotalContractValue) ? contractDetails.TotalContractValue : (parseFloat(String(projectDetails.TotalSOWAmount || 0).replace(/[^0-9.-]+/g, '')) || 0);
                    return `${projectDetails.Currency ? projectDetails.Currency + ' ' : ''}${val.toLocaleString()}`;
                  })()}</div>
                </div>
                <div className={styles.metricCard}>
                  <div className={styles.metricLabel}>Monthly Billing Value (Editable)</div>
                  <TextField
                    value={projectEditableField.BillingValue}
                    onChange={(ev, newValue) => handleRemarksChange('BillingValue', ev, newValue)}
                    placeholder="Enter monthly billing..."
                    prefix={projectDetails.Currency ? `${projectDetails.Currency} ` : ''}
                    styles={{
                      fieldGroup: { border: '1px solid #cbd5e1', borderRadius: '6px', background: '#ffffff' },
                      field: { fontWeight: 600, color: '#0f172a' }
                    }}
                  />
                </div>
                <div className={styles.metricCard}>
                  <div className={styles.metricLabel}>Deployed HeadCount</div>
                  <div className={styles.metricValue}>{deployedHeadCount}</div>
                </div>
                <div className={styles.metricCard}>
                  <div className={styles.metricLabel}>Billable HeadCount</div>
                  <div className={styles.metricValue}>{billableHeadCount !== null ? billableHeadCount.toFixed(2) : 'N/A'}</div>
                </div>
                <div className={styles.metricCard}>
                  <div className={styles.metricLabel}>Currency SOW Amount</div>
                  <div className={styles.metricValue}>{((): string => {
                    const sowNum = parseFloat(String(projectDetails.TotalSOWAmount || 0).replace(/[^0-9.-]+/g, '')) || 0;
                    return sowNum > 0 ? `${projectDetails.Currency ? projectDetails.Currency + ' ' : ''}${sowNum.toLocaleString()}` : (projectDetails.TotalSOWAmount ? `${projectDetails.Currency ? projectDetails.Currency + ' ' : ''}${projectDetails.TotalSOWAmount}` : 'N/A');
                  })()}</div>
                </div>
              </div>

              <div className={styles.remarksSection}>
                <div className={styles.remarkCard}>
                  <div className={styles.remarkLabel}>Project Description</div>
                  <TextField
                    value={projectEditableField.ProjectDescription}
                    onChange={(ev, newValue) => handleRemarksChange('ProjectDescription', ev, newValue)}
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
                      onChange={(ev, newValue) => handleRemarksChange('KeyUpdates', ev, newValue)}
                      multiline
                      rows={3}
                      placeholder="Enter key status updates..."
                    />
                  </div>
                  <div className={styles.remarkCard}>
                    <div className={styles.remarkLabel}>Next Steps / Other Updates</div>
                    <TextField
                      value={projectEditableField.NextStepUpdates}
                      onChange={(ev, newValue) => handleRemarksChange('NextStepUpdates', ev, newValue)}
                      multiline
                      rows={3}
                      placeholder="Enter planned upcoming milestones..."
                    />
                  </div>
                </div>
              </div>
            </div>

            {/* Risk Assessment Section */}
            <div className={styles.riskSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="Warning" style={{ fontSize: 20, color: '#0f172a' }} />
                <h3>Project RAG Status Assessment</h3>
              </div>

              <div className={styles.riskGrid}>
                {/* Overall Status */}
                <div className={styles.riskRowCard}>
                  <div className={styles.riskHeaderBar}>
                    <div className={styles.riskTitle}>Overall Status</div>
                    <div className={styles.riskSelector}>
                      {renderSegmentedControl('overallStatus')}
                    </div>
                  </div>
                  {shouldShowComment(riskAssessment.overallStatus) && (
                    <div className={styles.commentContainer}>
                      <TextField
                        label="Please provide details for non-green overall status"
                        value={riskAssessment.overallStatusComment}
                        onChange={(ev, newValue) => handleCommentChange('overallStatusComment', ev, newValue)}
                        multiline
                        rows={2}
                      />
                    </div>
                  )}
                </div>

                {/* Scope */}
                <div className={styles.riskRowCard}>
                  <div className={styles.riskHeaderBar}>
                    <div className={styles.riskTitle}>Scope</div>
                    <div className={styles.riskSelector}>
                      {renderSegmentedControl('scope')}
                    </div>
                  </div>
                  {shouldShowComment(riskAssessment.scope) && (
                    <div className={styles.commentContainer}>
                      <TextField
                        label="Please provide details for scope risk"
                        value={riskAssessment.scopeComment}
                        onChange={(ev, newValue) => handleCommentChange('scopeComment', ev, newValue)}
                        multiline
                        rows={2}
                      />
                    </div>
                  )}
                </div>

                {/* Schedule */}
                <div className={styles.riskRowCard}>
                  <div className={styles.riskHeaderBar}>
                    <div className={styles.riskTitle}>Schedule</div>
                    <div className={styles.riskSelector}>
                      {renderSegmentedControl('schedule')}
                    </div>
                  </div>
                  {shouldShowComment(riskAssessment.schedule) && (
                    <div className={styles.commentContainer}>
                      <TextField
                        label="Please provide details for schedule variance"
                        value={riskAssessment.scheduleComment}
                        onChange={(ev, newValue) => handleCommentChange('scheduleComment', ev, newValue)}
                        multiline
                        rows={2}
                      />
                    </div>
                  )}
                </div>

                {/* Deliverable Quality */}
                <div className={styles.riskRowCard}>
                  <div className={styles.riskHeaderBar}>
                    <div className={styles.riskTitle}>Deliverable Quality</div>
                    <div className={styles.riskSelector}>
                      {renderSegmentedControl('deliverableQuality')}
                    </div>
                  </div>
                  {shouldShowComment(riskAssessment.deliverableQuality) && (
                    <div className={styles.commentContainer}>
                      <TextField
                        label="Please provide details for quality issues"
                        value={riskAssessment.deliverableQualityComment}
                        onChange={(ev, newValue) => handleCommentChange('deliverableQualityComment', ev, newValue)}
                        multiline
                        rows={2}
                      />
                    </div>
                  )}
                </div>

                {/* Resource Availability */}
                <div className={styles.riskRowCard}>
                  <div className={styles.riskHeaderBar}>
                    <div className={styles.riskTitle}>Resource Availability</div>
                    <div className={styles.riskSelector}>
                      {renderSegmentedControl('resourceAvailability')}
                    </div>
                  </div>
                  {shouldShowComment(riskAssessment.resourceAvailability) && (
                    <div className={styles.commentContainer}>
                      <TextField
                        label="Please provide details for staffing/resource gaps"
                        value={riskAssessment.resourceAvailabilityComment}
                        onChange={(ev, newValue) => handleCommentChange('resourceAvailabilityComment', ev, newValue)}
                        multiline
                        rows={2}
                      />
                    </div>
                  )}
                </div>
              </div>

              {/* Action Footer */}
              <div className={styles.ctaBar}>
                <PrimaryButton
                  text="Export to PowerPoint"
                  iconProps={{ iconName: 'PowerPointDocument' }}
                  onClick={() => { executeAsync(handleExportToPPT()); }}
                />
              </div>
            </div>
          </>
        )}
      </div>
    </section>
  );
}
