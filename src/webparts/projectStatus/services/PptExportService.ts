import pptxgen from 'pptxgenjs';
import { IProjectItem, IProjectEditableField, IRiskAssessment, IContracts } from './ProjectService';
import { getCurrentDate, getOrdinalSuffix } from '../utils/dateUtils';

export interface IPptExportData {
  projectDetails: IProjectItem;
  projectEditableField: IProjectEditableField;
  riskAssessment: IRiskAssessment;
  contractDetails: IContracts | undefined;
  deployedHeadCount: number | undefined;
  billableHeadCount: number | undefined;
  monthlyBillingValue: number;
  backgroundImageUrl: string;
}

const COMPANY_TEMPLATE_URL = 'https://ytpl.sharepoint.com/:i:/r/sites/imsyashsharepoint/Shared%20Documents/SPFX%20Icons/DataTemplatePPT.png';

const getRiskColor = (status: string): string => {
  switch (status.toLowerCase()) {
    case 'red': return 'FF0000';
    case 'amber': return 'FFA500';
    default: return '00FF00';
  }
};

const trimText = (text: string): string => text.trim().replace(/\s+/g, ' ');

const formatKeyUpdates = (updates: string): string[] =>
  updates.split('\n').map(u => trimText(u)).filter(u => u !== '');

const RISK_STATUS_FIELDS = ['overallStatus', 'scope', 'schedule', 'deliverableQuality', 'resourceAvailability'] as const;

export const exportToPpt = async (data: IPptExportData): Promise<void> => {
  const {
    projectDetails, projectEditableField, riskAssessment, contractDetails,
    deployedHeadCount, billableHeadCount, monthlyBillingValue, backgroundImageUrl,
  } = data;

  // Total Contract Value from Contracts list; fall back to SOW amount
  const tcv = contractDetails?.TotalContractValue
    ?? (parseFloat(String(projectDetails.TotalSOWAmount || 0).replace(/[^0-9.-]+/g, '')) || 0);

  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';

  const now = new Date();
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const formattedDate = `${now.getDate()}${getOrdinalSuffix(now.getDate())} ${months[now.getMonth()]}'${String(now.getFullYear()).substring(2)}`;

  const slide = pptx.addSlide();
  // Use provided URL if set, otherwise fall back to the company template
  slide.background = { path: backgroundImageUrl || COMPANY_TEMPLATE_URL };

  // Project Name (above Project Owner — added per requirement)
  slide.addText(projectDetails.ProjectName, {
    x: 0.4, y: 0.05, w: 9.1, h: 0.2,
    color: '0073CF', fontSize: 18, bold: true,
  });

  // Project Owner
  slide.addText(`Project Owner: ${projectDetails.projectManager}`, {
    x: 0.4, y: 0.28, w: 5.0, h: 0.2,
    color: '0073CF', fontSize: 16, bold: false,
  });

  // Date stamp
  slide.addText(`Updated as on ${formattedDate}`, {
    x: 5.5, y: 0.28, w: 3.0, h: 0.2,
    color: '00CC00', fontSize: 12, bold: false,
  });

  // Description banner (white box)
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.2, y: 0.8, w: 9.4, h: 0.9,
    fill: { color: 'FFFFFF' }, line: { color: '000000', width: 1 },
  });
  slide.addText(`Project Description – ${projectEditableField.ProjectDescription || ''}`, {
    x: 0.3, y: 0.9, w: 9.2, h: 0.7,
    fontSize: 11, color: '000000',
  });

  // "Week Ending Status Report" header bar
  slide.addText('Week Ending Status Report: ' + getCurrentDate(), {
    x: 0.2, y: 1.8, w: 3.5, h: 0.2,
    fill: { color: '0073CF' }, color: 'FFFFFF',
    fontSize: 12, bold: false, valign: 'middle',
  });

  // ─── Left status / metrics table ──────────────────────────────────────────
  const hdr = { fill: { color: 'D3D3D3' }, color: '000000', bold: true, valign: 'middle' as const };
  const val = { fill: { color: 'F0F0F0' }, color: '000000', valign: 'middle' as const };

  const tableData = [
    [{ text: 'Type', options: hdr }, { text: projectDetails.ProjectType, options: val }],
    [{ text: 'Start Date', options: hdr }, { text: projectDetails.PlannedStartDate, options: val }],
    [{ text: 'End Date', options: hdr }, { text: projectDetails.PlannedEndDate, options: val }],
    ...RISK_STATUS_FIELDS.map(field => [
      {
        text: field.replace(/([A-Z])/g, ' $1').replace(/^./, s => s.toUpperCase()),
        options: hdr,
      },
      {
        text: riskAssessment[field].toUpperCase(),
        options: {
          fill: { color: getRiskColor(riskAssessment[field]) },
          color: riskAssessment[field] === 'green' ? '000000' : 'FFFFFF',
          valign: 'middle' as const,
        },
      },
    ]),
    [{ text: 'Billable Head Count', options: hdr }, { text: billableHeadCount !== undefined ? billableHeadCount.toFixed(2) : 'N/A', options: val }],
    [{ text: 'Deployed Head Count', options: hdr }, { text: deployedHeadCount !== undefined ? String(deployedHeadCount) : 'N/A', options: val }],
    [{ text: 'SOW Value', options: hdr }, { text: `${projectDetails.Currency} ${projectDetails.TotalSOWAmount}`, options: val }],
    [{ text: 'Total Contract Value', options: hdr }, { text: tcv > 0 ? `${projectDetails.Currency} ${tcv.toLocaleString()}` : 'N/A', options: val }],
    [{ text: 'Billing Value (Monthly)', options: { fill: { color: '00FF00' }, color: '000000', bold: true, valign: 'middle' as const } }, { text: monthlyBillingValue > 0 ? `${projectDetails.Currency} ${monthlyBillingValue.toLocaleString()}` : 'N/A', options: val }],
  ];

  slide.addTable(tableData as unknown as Parameters<typeof slide.addTable>[0], {
    x: 0.2, y: 2.0, w: 3.5,
    border: { type: 'solid', color: '000000', pt: 1 },
    fontSize: 10,
  });

  // ─── Right info table ──────────────────────────────────────────────────────
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
  const keyUpdatesBullets = formatKeyUpdates(projectEditableField.KeyUpdates).join('\n') || '';
  const nextStepsBullets = formatKeyUpdates(projectEditableField.NextStepUpdates).join('\n') || '';

  const infoHdr = { fill: { color: 'F0F0F0' }, color: '000000', bold: true, valign: 'middle' as const, align: 'top' as const };
  const infoVal = { fill: { color: 'FFFFFF' }, color: '000000', valign: 'top' as const, align: 'left' as const };

  const additionalInfoTable = [
    [{ text: 'Description for Red & Amber Status', options: infoHdr }, { text: trimText(riskText), options: infoVal }],
    [{ text: 'Key Updates', options: infoHdr }, { text: keyUpdatesBullets, options: infoVal }],
    [{ text: 'Next Steps / Other Updates', options: infoHdr }, { text: nextStepsBullets, options: infoVal }],
    [{ text: 'Risks / Challenge', options: infoHdr }, { text: 'No Major Updates', options: infoVal }],
  ];

  slide.addTable(additionalInfoTable as unknown as Parameters<typeof slide.addTable>[0], {
    x: 3.7, y: 1.8, w: 6.0, h: 3.6, colW: [1.5, 4.4],
    border: { type: 'solid', color: '000000', pt: 1 },
    fontSize: 10,
  });

  await pptx.writeFile({ fileName: `${projectDetails.ProjectName} Status Report ${getCurrentDate()}` });
};

