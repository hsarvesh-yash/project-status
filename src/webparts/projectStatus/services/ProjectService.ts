import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IDropdownOption } from '@fluentui/react';

// ─── Public domain types ─────────────────────────────────────────────────────

export interface IProjectItem {
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

export interface IProjectEditableField {
  KeyUpdates: string;
  ProjectDescription: string;
  NextStepUpdates: string;
  BillingValue: string;
}

export interface IRiskAssessment {
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

export interface IContracts {
  TotalContractValue: number;
  BillingValue: number;
}

export interface IWSROutputItem {
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

export interface ISPWSROutputRecord {
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

// ─── Internal SP types ────────────────────────────────────────────────────────

interface ISPProjectItem {
  ProjectID: string;
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
  AllocationID?: string;
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

// ─── Helpers ──────────────────────────────────────────────────────────────────

/** Prevent OData filter injection by escaping single quotes. */
const escapeOData = (value: string): string => value.replace(/'/g, "''");

const ALLOCATION_BATCH_SIZE = 15;

const formatDateDisplay = (dateString: string | undefined): string => {
  if (!dateString) return '';
  try {
    return new Date(dateString)
      .toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })
      .replace(/ /g, '-');
  } catch {
    return dateString;
  }
};

// ─── Service functions ────────────────────────────────────────────────────────

export const getProjectList = async (userEmail: string): Promise<IDropdownOption[]> => {
  const safe = escapeOData(userEmail.toLowerCase());
  const data: ISPProjectItem[] = await sp.web.lists
    .getByTitle('IMSProjects')
    .items
    .select('*,ProjectID,CustomerName,ProjectName,ProjectType,TotalSOWAmount,Currency,AdditionalPmOrDMId,AdditionalPmOrDM/Title,AdditionalPmOrDM/EMail,DeliveryManagerId,DeliveryManager/Title,DeliveryManager/EMail,ProjectManagerId,ProjectManager/Title,ProjectManager/EMail,QualityEngineerId,QualityEngineer/Title,QualityLeadId,QualityLead/Title,QualityManagerId,QualityManager/Title,QualityLead/EMail,QualityManager/EMail,QualityEngineer/EMail,AccountManager/Title,AccountManager/EMail,ProjectGEOS')
    .expand('AdditionalPmOrDM,DeliveryManager,ProjectManager,QualityEngineer,QualityLead,QualityManager,AccountManager')
    .filter(`(ProjectManager/EMail eq '${safe}' or DeliveryManager/EMail eq '${safe}' or AdditionalPmOrDM/EMail eq '${safe}' or AccountManager/EMail eq '${safe}' or QualityLead/EMail eq '${safe}' or QualityManager/EMail eq '${safe}' or QualityEngineer/EMail eq '${safe}') and IsProcess eq 'Yes'`)();
  return data.map(item => ({ key: item.ProjectID, text: item.ProjectName }));
};

export const getProjectDetails = async (projectId: string): Promise<IProjectItem | null> => {
  const safe = escapeOData(projectId);
  const data: ISPProjectItem[] = await sp.web.lists
    .getByTitle('IMSProjects')
    .items
    .filter(`ProjectID eq '${safe}' and IsProcess eq 'Yes'`)
    .select('*,ProjectID,CustomerName,ProjectName,ProjectType,PlannedStartDate,PlannedEndDate,DeliveryManager/Title,ProjectManager/Title,Currency,TotalSOWAmount,ProjectGEOS,QualityEngineer/Title')
    .expand('DeliveryManager,ProjectManager,QualityEngineer')();

  if (!data || data.length === 0) return null;

  const raw = data[0] as unknown as Record<string, unknown>;
  const rawSow = raw.TotalSOWAmount ?? raw.SOWAmount ?? raw.TotalSOW ?? raw.SOWValue ?? raw.ContractValue ?? '';

  return {
    ProjectID: data[0].ProjectID,
    ProjectName: data[0].ProjectName,
    ProjectType: data[0].ProjectType || 'Not specified',
    PlannedStartDate: formatDateDisplay(data[0].PlannedStartDate) || 'Not specified',
    PlannedEndDate: formatDateDisplay(data[0].PlannedEndDate) || 'Not specified',
    projectManager: data[0].ProjectManager?.Title || '',
    deliveryManager: data[0].DeliveryManager?.Title || '',
    TotalSOWAmount: String(rawSow),
    Currency: data[0].Currency || '',
    ProjectGEOS: data[0].ProjectGEOS || '',
    QualityEngineer: data[0].QualityEngineer?.Title || '',
  };
};

export const getResourceAllocation = async (projectId: string): Promise<{ deployed: number; billable: number }> => {
  const safe = escapeOData(projectId);
  const mappings: ISPResourceMapping[] = await sp.web.lists
    .getByTitle('ProjectResourceMapping')
    .items
    .top(4999)
    .select('*,Id,ResourceId,ProjectId,AllocationID')
    .filter(`ProjectId eq '${safe}'`)();

  if (mappings.length === 0) return { deployed: 0, billable: 0 };

  const seen = new Set<string>();
  const allocationIds: string[] = [];
  mappings.forEach(m => {
    if (m.AllocationID && !seen.has(m.AllocationID)) {
      seen.add(m.AllocationID);
      allocationIds.push(m.AllocationID);
    }
  });
  if (allocationIds.length === 0) return { deployed: 0, billable: 0 };

  // Batch to avoid OData URL length limits
  const allAllocations: ISPResourceAllocation[] = [];
  for (let i = 0; i < allocationIds.length; i += ALLOCATION_BATCH_SIZE) {
    const chunk = allocationIds.slice(i, i + ALLOCATION_BATCH_SIZE);
    const filter = chunk.map(id => `AllocationID eq '${escapeOData(id)}'`).join(' or ');
    const batch: ISPResourceAllocation[] = await sp.web.lists
      .getByTitle('ResourceAllocation')
      .items
      .select('FTEValue,Allocation,AllocationID,Billable')
      .filter(filter)
      .top(5000)();
    allAllocations.push(...batch);
  }

  const billable = allAllocations.reduce((sum, item) => sum + (Number(item.FTEValue) || 0), 0);
  return { deployed: allAllocations.length, billable };
};

export const getContractData = async (projectId: string): Promise<IContracts | null> => {
  const safe = escapeOData(projectId);
  let contractItems: ISPContractItem[] = [];

  // Strategy 1: Direct field filter
  try {
    contractItems = await sp.web.lists.getByTitle('Contracts').items
      .top(5000)
      .filter(`ProjectID eq '${safe}' or ProjectId eq '${safe}'`)();
  } catch { /* fall through */ }

  // Strategy 2: Lookup expand
  if (contractItems.length === 0) {
    try {
      contractItems = await sp.web.lists.getByTitle('Contracts').items
        .top(5000)
        .select('*,ProjectID/ProjectID,ProjectID/Title,TotalContractValue')
        .expand('ProjectID')
        .filter(`ProjectID/ProjectID eq '${safe}' or ProjectID/Title eq '${safe}'`)();
    } catch { /* fall through */ }
  }

  // Strategy 3: In-memory fallback (last resort — logs a warning)
  if (contractItems.length === 0) {
    try {
      console.warn('[getContractData] Falling back to in-memory filter for ProjectID:', projectId);
      const all: ISPContractItem[] = await sp.web.lists.getByTitle('Contracts').items.top(2000)();
      contractItems = all.filter(item => {
        const pid = (item.ProjectID as Record<string, unknown>)?.ProjectID
          || (item.ProjectID as Record<string, unknown>)?.Title
          || item.ProjectID
          || item.ProjectId;
        return pid && String(pid).trim().toLowerCase() === projectId.trim().toLowerCase();
      });
    } catch { /* all strategies exhausted */ }
  }

  if (contractItems.length === 0) return null;

  const item = contractItems[0];
  const rawVal = item.TotalContractValue ?? item.ContractValue ?? item.TotalValue ?? item.Amount ?? 0;
  const numericVal = typeof rawVal === 'number' ? rawVal : parseFloat(String(rawVal).replace(/[^0-9.-]+/g, '')) || 0;

  const rawBilling = item.BillingValue ?? item.BilledValue ?? item.BillingAmount
    ?? item.InvoicedAmount ?? item.TotalInvoiced ?? item.InvoiceAmount ?? item.TotalBilled ?? 0;
  const numericBilling = typeof rawBilling === 'number' ? rawBilling : parseFloat(String(rawBilling).replace(/[^0-9.-]+/g, '')) || 0;

  return { TotalContractValue: numericVal, BillingValue: numericBilling };
};

export const checkTodayWSR = async (projectId: string, reportTitle: string): Promise<ISPWSROutputRecord | null> => {
  const safeId = escapeOData(projectId);
  const safeTitle = escapeOData(reportTitle);
  const items: ISPWSROutputRecord[] = await sp.web.lists
    .getByTitle('WSROutput')
    .items
    .filter(`ProjectID eq '${safeId}' and Title eq '${safeTitle}'`)
    .select('*')
    .get();
  return items.length > 0 ? items[0] : null;
};

export const getAllWSRForProject = async (projectId: string): Promise<ISPWSROutputRecord[]> => {
  const safe = escapeOData(projectId);
  return sp.web.lists
    .getByTitle('WSROutput')
    .items
    .filter(`ProjectID eq '${safe}'`)
    .select('*')
    .orderBy('Created', false)
    .get();
};

export const saveWSROutput = async (data: Partial<IWSROutputItem>): Promise<void> => {
  await sp.web.lists.getByTitle('WSROutput').items.add(data);
};
