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
  TotalContractValue?: number;
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
  Currency?: string;
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
  ProjectID?: string;
  ProjectName?: string;
  ProjectType?: string;
  PlannedStartDate?: string;
  PlannedEndDate?: string;
  ProjectManager?: string;
  DeliveryManager?: string;
  TotalSOWAmount?: string;
  Currency?: string;
  ProjectGEOS?: string;
  QualityEngineer?: string;
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

export const formatDateDisplay = (dateString: string | undefined): string => {
  if (!dateString) return '';
  try {
    const d = new Date(dateString);
    if (isNaN(d.getTime())) return dateString;
    return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/ /g, '-');
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

export const getProjectDetails = async (projectId: string): Promise<IProjectItem | undefined> => {
  const safe = escapeOData(projectId);
  const data: ISPProjectItem[] = await sp.web.lists
    .getByTitle('IMSProjects')
    .items
    .filter(`ProjectID eq '${safe}' and IsProcess eq 'Yes'`)
    .select('*,ProjectID,CustomerName,ProjectName,ProjectType,PlannedStartDate,PlannedEndDate,DeliveryManager/Title,ProjectManager/Title,Currency,TotalSOWAmount,ProjectGEOS,QualityEngineer/Title')
    .expand('DeliveryManager,ProjectManager,QualityEngineer')();

  if (!data || data.length === 0) return undefined;

  const raw = data[0] as unknown as Record<string, unknown>;
  const rawSow = raw.TotalSOWAmount ?? raw.SOWAmount ?? raw.TotalSOW ?? raw.SOWValue ?? raw.ContractValue ?? raw.TotalContractValue ?? '';
  const rawCurrency = raw.Currency ?? raw.Currecncy ?? raw.ProjectCurrency ?? raw.CurrencyCode ?? raw.SOWCurrency ?? raw.ContractCurrency ?? 'USD';
  const rawTcv = raw.TotalContractValue ?? raw.ContractValue ?? raw.TCV ?? raw.TotalContractAmount ?? rawSow ?? 0;
  const numTcv = typeof rawTcv === 'number' ? rawTcv : parseFloat(String(rawTcv).replace(/[^0-9.-]+/g, '')) || 0;

  return {
    ProjectID: data[0].ProjectID,
    ProjectName: data[0].ProjectName,
    ProjectType: data[0].ProjectType || 'Not specified',
    PlannedStartDate: formatDateDisplay(data[0].PlannedStartDate) || 'Not specified',
    PlannedEndDate: formatDateDisplay(data[0].PlannedEndDate) || 'Not specified',
    projectManager: data[0].ProjectManager?.Title || '',
    deliveryManager: data[0].DeliveryManager?.Title || '',
    TotalSOWAmount: String(rawSow),
    Currency: String(rawCurrency) || 'USD',
    ProjectGEOS: data[0].ProjectGEOS || '',
    QualityEngineer: data[0].QualityEngineer?.Title || '',
    TotalContractValue: numTcv,
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

export const getContractData = async (projectId: string): Promise<IContracts | undefined> => {
  const safe = escapeOData(projectId);
  console.log('[getContractData] Starting fetch for projectId:', projectId);
  let contractItems: Record<string, unknown>[] = [];

  try {
    const res: Record<string, unknown>[] = await sp.web.lists.getByTitle('Contracts').items
      .top(100)
      .filter(`ProjectId eq '${safe}'`)();
    console.log('[getContractData] Filter ProjectId result count:', res.length);
    if (res.length > 0) contractItems = res;
  } catch (err) {
    console.warn('[getContractData] Filter ProjectId failed:', err);
  }

  if (contractItems.length === 0) {
    try {
      const res: Record<string, unknown>[] = await sp.web.lists.getByTitle('Contracts').items
        .top(100)
        .filter(`ProjectID eq '${safe}'`)();
      console.log('[getContractData] Filter ProjectID result count:', res.length);
      if (res.length > 0) contractItems = res;
    } catch (err) {
      console.warn('[getContractData] Filter ProjectID failed:', err);
    }
  }

  if (contractItems.length === 0) {
    try {
      console.log('[getContractData] Fetching top 2000 items from Contracts for in-memory match...');
      const all: Record<string, unknown>[] = await sp.web.lists.getByTitle('Contracts').items.top(2000)();
      console.log('[getContractData] Total items fetched from Contracts:', all.length);
      if (all.length > 0) {
        console.log('[getContractData] Sample contract item keys:', Object.keys(all[0]));
      }
      contractItems = all.filter(item => {
        const p1 = String(item.ProjectId ?? '').trim().toLowerCase();
        const p2 = String(item.ProjectID ?? '').trim().toLowerCase();
        const target = projectId.trim().toLowerCase();
        return p1 === target || p2 === target;
      });
      console.log('[getContractData] In-memory matched items count:', contractItems.length);
    } catch (err) {
      console.error('[getContractData] In-memory fetch failed:', err);
    }
  }

  if (contractItems.length === 0) {
    console.warn('[getContractData] No contract items matched for:', projectId);
    return undefined;
  }

  const item = contractItems[0];
  console.log('[getContractData] Found matching contract item:', item);

  const rawVal = item.TotalContractValue 
    ?? item.TotalContractsValue 
    ?? item.Total_x0020_Contracts_x0020_value 
    ?? item.Total_x0020_Contract_x0020_Value 
    ?? item['Total Contracts value'] 
    ?? item['Total Contract Value'] 
    ?? item.ContractValue 
    ?? item.Amount 
    ?? 0;
  const numericVal = typeof rawVal === 'number' ? rawVal : parseFloat(String(rawVal).replace(/[^0-9.-]+/g, '')) || 0;

  const rawBilling = item.BillingValue ?? item.BilledValue ?? item.BillingAmount
    ?? item.InvoicedAmount ?? item.TotalInvoiced ?? item.InvoiceAmount ?? item.TotalBilled ?? 0;
  const numericBilling = typeof rawBilling === 'number' ? rawBilling : parseFloat(String(rawBilling).replace(/[^0-9.-]+/g, '')) || 0;

  const contractCurrency = String(item.Currecncy ?? item.Currency ?? item.ContractCurrency ?? item.ProjectCurrency ?? '').trim();

  console.log('[getContractData] Parsed result -> TotalContractValue:', numericVal, 'BillingValue:', numericBilling, 'Currency:', contractCurrency);

  return { TotalContractValue: numericVal, BillingValue: numericBilling, Currency: contractCurrency || undefined };
};

export const checkTodayWSR = async (projectId: string, reportTitle: string): Promise<ISPWSROutputRecord | undefined> => {
  const safeId = escapeOData(projectId);
  const safeTitle = escapeOData(reportTitle);
  const items: ISPWSROutputRecord[] = await sp.web.lists
    .getByTitle('WSROutput')
    .items
    .filter(`ProjectID eq '${safeId}' and Title eq '${safeTitle}'`)
    .select('*')
    .get();
  return items.length > 0 ? items[0] : undefined;
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

export const getInvoiceData = async (projectId: string): Promise<number | undefined> => {
  const safe = escapeOData(projectId);
  console.log('[getInvoiceData] Starting fetch for projectId:', projectId);

  // Focus directly on 'Invoices' as provided in user schema
  const listName = 'Invoices';
  let invoiceItems: Record<string, unknown>[] = [];

  try {
    const res: Record<string, unknown>[] = await sp.web.lists.getByTitle(listName).items
      .top(50)
      .filter(`ProjectId eq '${safe}'`)
      .orderBy('Created', false)();
    console.log('[getInvoiceData] Filter ProjectId result count:', res.length);
    if (res.length > 0) invoiceItems = res;
  } catch (err) {
    console.warn('[getInvoiceData] Filter ProjectId failed:', err);
  }

  if (invoiceItems.length === 0) {
    try {
      const res: Record<string, unknown>[] = await sp.web.lists.getByTitle(listName).items
        .top(50)
        .filter(`ProjectID eq '${safe}'`)
        .orderBy('Created', false)();
      console.log('[getInvoiceData] Filter ProjectID result count:', res.length);
      if (res.length > 0) invoiceItems = res;
    } catch (err) {
      console.warn('[getInvoiceData] Filter ProjectID failed:', err);
    }
  }

  if (invoiceItems.length === 0) {
    try {
      console.log('[getInvoiceData] Fetching top 1000 items from Invoices for in-memory match...');
      const all: Record<string, unknown>[] = await sp.web.lists.getByTitle(listName).items.top(1000).orderBy('Created', false)();
      console.log('[getInvoiceData] Total items fetched from Invoices:', all.length);
      if (all.length > 0) {
        console.log('[getInvoiceData] Sample invoice item keys:', Object.keys(all[0]));
      }
      invoiceItems = all.filter(item => {
        const p1 = String(item.ProjectId ?? '').trim().toLowerCase();
        const p2 = String(item.ProjectID ?? '').trim().toLowerCase();
        const target = projectId.trim().toLowerCase();
        return p1 === target || p2 === target;
      });
      console.log('[getInvoiceData] In-memory matched invoice items count:', invoiceItems.length);
    } catch (err) {
      console.error('[getInvoiceData] In-memory fetch failed:', err);
    }
  }

  if (invoiceItems.length > 0) {
    const item = invoiceItems[0];
    console.log('[getInvoiceData] Found matching invoice item:', item);
    const val = item.InvoiceAmount 
      ?? item.InvoiceAmount0 
      ?? item.InvoiceAmount1 
      ?? item.Invoice_x0020_Amount 
      ?? item['Invoice Amount'] 
      ?? item.BillingValue 
      ?? item.Amount 
      ?? 0;
    const num = typeof val === 'number' ? val : parseFloat(String(val).replace(/[^0-9.-]+/g, '')) || 0;
    console.log('[getInvoiceData] Parsed InvoiceAmount:', num);
    if (num > 0) return num;
  }

  console.warn('[getInvoiceData] No valid invoice amount found for:', projectId);
  return undefined;
};
