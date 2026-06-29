import pptxgen from 'pptxgenjs';
import { IProjectItem, IProjectEditableField, IRiskAssessment, IContracts } from './ProjectService';
import { getCurrentDate, getOrdinalSuffix } from '../utils/dateUtils';

export interface IPptExportData {
  projectDetails: IProjectItem;
  projectEditableField: IProjectEditableField;
  riskAssessment: IRiskAssessment;
  contractDetails: IContracts | null;
  deployedHeadCount: number | null;
  billableHeadCount: number | null;
  numBilling: number;
  backgroundImageUrl: string;
}

const getRiskColor = (status: string): string => {
  switch (status.toLowerCase()) {
    case 'red': return 'EF4444';
    case 'amber': return 'F59E0B';
    default: return '10B981';
  }
};

const trimText = (text: string): string => text.trim().replace(/\s+/g, ' ');

const formatKeyUpdates = (updates: string): string[] =>
  updates.split('\n').map(u => trimText(u)).filter(u => u !== '');

const formatCurrency = (currency: string, value: number | string): string => {
  const prefix = currency ? `${currency} ` : '';
  const num = typeof value === 'number' ? value : parseFloat(String(value).replace(/[^0-9.-]+/g, '')) || 0;
  return `${prefix}${num.toLocaleString()}`;
};

const RISK_STATUS_FIELDS = ['overallStatus', 'scope', 'schedule', 'deliverableQuality', 'resourceAvailability'] as const;

export const exportToPpt = async (data: IPptExportData): Promise<void> => {
  const {
    projectDetails, projectEditableField, riskAssessment, contractDetails,
    deployedHeadCount, billableHeadCount, numBilling, backgroundImageUrl,
  } = data;

  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';

  const now = new Date();
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const formattedDate = `${now.getDate()}${getOrdinalSuffix(now.getDate())} ${months[now.getMonth()]}'${String(now.getFullYear()).substring(2)}`;

  const tcv = contractDetails?.TotalContractValue
    ?? (parseFloat(String(projectDetails.TotalSOWAmount || 0).replace(/[^0-9.-]+/g, '')) || 0);

  const slide = pptx.addSlide();
  slide.background = { path: backgroundImageUrl };

  slide.addText(`Project Owner: ${projectDetails.projectManager}`, {
    x: 0.4, y: 0.28, w: 5.0, h: 0.3, color: '0F172A', fontSize: 15, bold: true,
  });
  slide.addText(`Updated as on ${formattedDate}`, {
    x: 5.5, y: 0.28, w: 3.8, h: 0.3, color: '059669', fontSize: 12, bold: true, align: 'right',
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.2, y: 0.72, w: 0.06, h: 0.9,
    fill: { color: '0073CF' }, line: { color: '0073CF', width: 0 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.26, y: 0.72, w: 9.34, h: 0.9,
    fill: { color: 'F8FAFC' }, line: { color: 'CBD5E1', width: 0.5 },
  });
  slide.addText([
    { text: 'Project Description: ', options: { bold: true, color: '0F172A', fontSize: 11 } },
    { text: projectEditableField.ProjectDescription || 'No description provided.', options: { color: '334155', fontSize: 10.5 } },
  ], { x: 0.38, y: 0.76, w: 9.1, h: 0.8, valign: 'top', wrap: true });

  slide.addText('Week Ending Status Report: ' + getCurrentDate(), {
    x: 0.2, y: 1.72, w: 3.4, h: 0.3,
    fill: { color: '0F172A' }, color: 'FFFFFF',
    fontSize: 11, bold: true, valign: 'middle', align: 'center',
  });

  const headerOpts = { fill: { color: '1E293B' }, color: 'FFFFFF', bold: true, valign: 'middle' };
  const altRow = { fill: { color: 'F8FAFC' }, color: '1E293B', valign: 'middle' };
  const plainRow = { fill: { color: 'FFFFFF' }, color: '1E293B', valign: 'middle' };

  const tableData = [
    [{ text: 'Type', options: headerOpts }, { text: projectDetails.ProjectType, options: altRow }],
    [{ text: 'Start Date', options: headerOpts }, { text: projectDetails.PlannedStartDate, options: plainRow }],
    [{ text: 'End Date', options: headerOpts }, { text: projectDetails.PlannedEndDate, options: altRow }],
    ...RISK_STATUS_FIELDS.map(field => [
      {
        text: field.replace(/([A-Z])/g, ' $1').replace(/^./, s => s.toUpperCase()),
        options: headerOpts,
      },
      {
        text: riskAssessment[field].toUpperCase(),
        options: {
          fill: { color: getRiskColor(riskAssessment[field]) },
          color: 'FFFFFF', bold: true, align: 'center', valign: 'middle',
        },
      },
    ]),
    [{ text: 'Billable Head Count', options: headerOpts }, { text: billableHeadCount !== null ? billableHeadCount.toFixed(2) : 'N/A', options: altRow }],
    [{ text: 'Deployed Head Count', options: headerOpts }, { text: deployedHeadCount !== null ? String(deployedHeadCount) : 'N/A', options: plainRow }],
    [{ text: 'SOW Value', options: headerOpts }, { text: formatCurrency(projectDetails.Currency, projectDetails.TotalSOWAmount), options: { ...altRow, bold: true } }],
    [{ text: 'Billing Value', options: headerOpts }, { text: formatCurrency(projectDetails.Currency, numBilling), options: { ...plainRow, bold: true } }],
    [{ text: 'Total Contract Value', options: headerOpts }, { text: formatCurrency(projectDetails.Currency, tcv), options: { ...altRow, bold: true } }],
  ];

  slide.addTable(tableData as unknown as Parameters<typeof slide.addTable>[0], {
    x: 0.2, y: 2.05, w: 3.4, colW: [1.4, 2.0],
    border: { type: 'solid', color: 'CBD5E1', pt: 0.5 }, fontSize: 9.5,
  });

  const riskComments = RISK_STATUS_FIELDS
    .filter(f => riskAssessment[f] === 'red' || riskAssessment[f] === 'amber')
    .map(f => {
      const comment = riskAssessment[`${f}Comment` as keyof IRiskAssessment] as string;
      return comment
        ? `${f.replace(/([A-Z])/g, ' $1').replace(/^./, s => s.toUpperCase())}: ${trimText(comment)}`
        : null;
    })
    .filter((c): c is string => c !== null);

  const riskText = riskComments.length > 0 ? riskComments.join('\n') : 'No Major Updates.';
  const keyUpdatesBullets = formatKeyUpdates(projectEditableField.KeyUpdates).join('\n') || 'No major updates.';
  const nextStepsBullets = formatKeyUpdates(projectEditableField.NextStepUpdates).join('\n') || 'No upcoming updates.';

  const infoHeaderOpts = { fill: { color: '0F172A' }, color: 'FFFFFF', bold: true, valign: 'middle', align: 'left' };
  const additionalInfoTable = [
    [{ text: 'Description for Red & Amber Status', options: infoHeaderOpts }, { text: trimText(riskText), options: { fill: { color: 'F8FAFC' }, color: '334155', valign: 'top', align: 'left' } }],
    [{ text: 'Key Updates', options: infoHeaderOpts }, { text: keyUpdatesBullets, options: { fill: { color: 'FFFFFF' }, color: '334155', valign: 'top', align: 'left' } }],
    [{ text: 'Next Steps / Other Updates', options: infoHeaderOpts }, { text: nextStepsBullets, options: { fill: { color: 'F8FAFC' }, color: '334155', valign: 'top', align: 'left' } }],
    [{ text: 'Risks / Challenges', options: infoHeaderOpts }, { text: riskText !== 'No Major Updates.' ? riskText : 'No Major Risks Identified.', options: { fill: { color: 'FFFFFF' }, color: '334155', valign: 'top', align: 'left' } }],
  ];

  slide.addTable(additionalInfoTable as unknown as Parameters<typeof slide.addTable>[0], {
    x: 3.8, y: 1.72, w: 5.8, h: 3.6, colW: [1.6, 4.2],
    border: { type: 'solid', color: 'CBD5E1', pt: 0.5 }, fontSize: 10,
  });

  await pptx.writeFile({ fileName: `${projectDetails.ProjectName} Status Report ${getCurrentDate()}` });
};
