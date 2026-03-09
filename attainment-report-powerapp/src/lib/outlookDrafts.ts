import { Office365OutlookService } from '../generated';
import type { Office365OutlookModel } from '../generated';
import type { ManagerReportRecord } from './reportEngine';

export const DEFAULT_EMAIL_SUBJECT_TEMPLATE = '{fiscal_year} Attainment Report - {manager_name}';
export const DEFAULT_EMAIL_BODY_TEMPLATE = [
  'Hi {manager_name},',
  '',
  'Please find attached your {fiscal_year} Attainment Report.',
  '',
  'This report includes attainment data for your team with quarterly, half-year, and annual breakdowns.',
  '',
  'If you have any questions about the data, please reach out to the Sales Compensation team.',
  '',
  'Best regards,',
  'Sales Compensation',
].join('\n');

export interface DraftProgress {
  current: number;
  total: number;
  managerName: string;
}

export interface DraftCreationResult {
  created: number;
  failed: number;
  failures: string[];
  folderUsed: string;
}

export interface SendProgress {
  current: number;
  total: number;
  subject: string;
}

export interface SendDraftEmailsResult {
  totalFound: number;
  sent: number;
  failed: number;
  failures: string[];
  folderPath: string;
}

export class ManualDraftFolderRequiredError extends Error {
  readonly folderPath: string;

  constructor(folderPath: string, reason: string) {
    super(`${reason} Please create ${folderPath} in Outlook and click Create Outlook Drafts again.`);
    this.name = 'ManualDraftFolderRequiredError';
    this.folderPath = folderPath;
  }
}

interface DraftOptions {
  fiscalYear: string;
  subjectTemplate: string;
  bodyTemplate: string;
  targetFolderName?: string;
}

interface MailFolderItem {
  id?: string;
  displayName?: string;
}

interface DraftMessageItem {
  id: string;
  subject: string;
}

interface DraftSubFolderResolution {
  folderId: string | null;
  note: string | null;
}

const DEFAULT_DRAFT_FOLDER = 'Manager Report';
const DRAFT_FOLDER_PATH_PREFIX = 'Drafts';
const EXCEL_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

export async function createOutlookDrafts(
  reports: ManagerReportRecord[],
  options: DraftOptions,
  onProgress?: (progress: DraftProgress) => void,
): Promise<DraftCreationResult> {
  const targetFolderName = options.targetFolderName || DEFAULT_DRAFT_FOLDER;
  const targetFolderPath = `${DRAFT_FOLDER_PATH_PREFIX}/${targetFolderName}`;
  const folderResolution = await ensureDraftSubFolder(targetFolderName);
  const folderId = folderResolution.folderId;

  let created = 0;
  let failed = 0;
  const failures: string[] = [];

  for (let i = 0; i < reports.length; i += 1) {
    const report = reports[i];
    onProgress?.({
      current: i + 1,
      total: reports.length,
      managerName: report.displayName,
    });

    if (!report.email) {
      failed += 1;
      failures.push(`${report.displayName}: missing email`);
      continue;
    }

    try {
      const subject = resolveTemplate(options.subjectTemplate, report.displayName, options.fiscalYear);
      const body = resolveTemplate(options.bodyTemplate, report.displayName, options.fiscalYear);
      const htmlBody = plainTextToHtml(body);

      let messageId: string | null = null;
      const createErrors: string[] = [];

      try {
        messageId = await createDraftWithConnector({
          to: report.email,
          subject,
          htmlBody,
          attachmentName: report.fileName,
          attachmentBase64: report.attachmentBase64,
        });
      } catch (connectorError) {
        createErrors.push(`DraftEmail failed: ${toErrorMessage(connectorError)}`);
      }

      if (!messageId) {
        try {
          messageId = await createDraftWithAttachmentHttp({
            to: report.email,
            subject,
            htmlBody,
            attachmentName: report.fileName,
            attachmentBase64: report.attachmentBase64,
          });
        } catch (httpError) {
          createErrors.push(`Graph HttpRequest failed: ${toErrorMessage(httpError)}`);
        }
      }

      if (!messageId) {
        throw new Error(createErrors.join(' | ') || 'Unable to create draft message.');
      }

      let movedMessageId = messageId;
      try {
        movedMessageId = await moveDraftToFolder(
          messageId,
          folderId,
          targetFolderPath,
          targetFolderName,
          folderResolution.note,
        );
      } catch (moveError) {
        await deleteDraftIfExists(messageId);
        throw moveError;
      }

      await markMessageAsUnread(movedMessageId);
      created += 1;
    } catch (error) {
      if (isManualDraftFolderRequiredError(error)) {
        throw error;
      }

      failed += 1;
      failures.push(`${report.displayName}: ${toErrorMessage(error)}`);
    }

    if ((i + 1) % 3 === 0) {
      await yieldToUi();
    }
  }

  return {
    created,
    failed,
    failures,
    folderUsed: targetFolderPath,
  };
}

async function createDraftWithConnector(params: {
  to: string;
  subject: string;
  htmlBody: string;
  attachmentName: string;
  attachmentBase64: string;
}): Promise<string> {
  const draftMessage: Office365OutlookModel.ClientDraftHtmlMessage = {
    To: params.to,
    Subject: params.subject,
    Body: params.htmlBody,
    Attachments: [
      {
        Name: params.attachmentName,
        ContentBytes: params.attachmentBase64,
      },
    ],
  };

  const draftResult = await Office365OutlookService.DraftEmail(draftMessage);
  if (!draftResult.success) {
    throw new Error(toOperationErrorMessage(draftResult.error));
  }

  const messageId = asOptionalString((draftResult.data as { Id?: unknown } | undefined)?.Id);
  if (!messageId) {
    throw new Error('Draft created but message ID was not returned.');
  }

  return messageId;
}

async function createDraftWithAttachmentHttp(params: {
  to: string;
  subject: string;
  htmlBody: string;
  attachmentName: string;
  attachmentBase64: string;
}): Promise<string> {
  const createMessageResponse = await Office365OutlookService.HttpRequest(
    '/me/messages',
    'POST',
    JSON.stringify({
      subject: params.subject,
      body: {
        contentType: 'HTML',
        content: params.htmlBody,
      },
      toRecipients: [
        {
          emailAddress: {
            address: params.to,
          },
        },
      ],
    }),
    'application/json',
  );

  if (!createMessageResponse.success) {
    throw new Error(`Create draft via Graph failed: ${toOperationErrorMessage(createMessageResponse.error)}`);
  }

  const messageId = extractMessageIdFromUnknown(createMessageResponse.data);
  if (!messageId) {
    throw new Error('Draft created but message ID was missing in Graph response.');
  }

  const addAttachmentResponse = await Office365OutlookService.HttpRequest(
    `/me/messages/${encodeURIComponent(messageId)}/attachments`,
    'POST',
    JSON.stringify({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: params.attachmentName,
      contentType: EXCEL_MIME_TYPE,
      contentBytes: params.attachmentBase64,
    }),
    'application/json',
  );

  if (!addAttachmentResponse.success) {
    await deleteDraftIfExists(messageId);
    throw new Error(`Add attachment via Graph failed: ${toOperationErrorMessage(addAttachmentResponse.error)}`);
  }

  return messageId;
}

export async function sendAllDraftEmailsFromFolder(
  folderName = DEFAULT_DRAFT_FOLDER,
  onProgress?: (progress: SendProgress) => void,
): Promise<SendDraftEmailsResult> {
  const candidatePaths = [
    `${DRAFT_FOLDER_PATH_PREFIX}/${folderName}`,
    `${DRAFT_FOLDER_PATH_PREFIX}\\${folderName}`,
    folderName,
  ];

  let drafts: DraftMessageItem[] = [];
  let usedPath = candidatePaths[0];

  for (const path of candidatePaths) {
    const listResult = await listDraftMessages(path);
    if (listResult.success) {
      drafts = listResult.items;
      usedPath = path;
      break;
    }
  }

  if (drafts.length === 0) {
    return {
      totalFound: 0,
      sent: 0,
      failed: 0,
      failures: [],
      folderPath: usedPath,
    };
  }

  let sent = 0;
  let failed = 0;
  const failures: string[] = [];

  for (let i = 0; i < drafts.length; i += 1) {
    const draft = drafts[i];
    onProgress?.({
      current: i + 1,
      total: drafts.length,
      subject: draft.subject,
    });

    const sendResult = await Office365OutlookService.SendDraftEmail(draft.id);
    if (sendResult.success) {
      sent += 1;
    } else {
      failed += 1;
      failures.push(`${draft.subject}: ${toOperationErrorMessage(sendResult.error)}`);
    }

    if ((i + 1) % 5 === 0) {
      await yieldToUi();
    }
  }

  return {
    totalFound: drafts.length,
    sent,
    failed,
    failures,
    folderPath: usedPath,
  };
}

async function ensureDraftSubFolder(folderName: string): Promise<DraftSubFolderResolution> {
  const errors: string[] = [];
  const draftsFolder = await resolveDraftsFolder(errors);

  const childFolderUris = new Set<string>([
    '/me/mailFolders/drafts/childFolders',
    '/me/mailFolders/Drafts/childFolders',
    'https://graph.microsoft.com/v1.0/me/mailFolders/drafts/childFolders',
    'https://graph.microsoft.com/v1.0/me/mailFolders/Drafts/childFolders',
  ]);

  if (draftsFolder?.id) {
    const encodedId = encodeURIComponent(draftsFolder.id);
    childFolderUris.add(`/me/mailFolders/${encodedId}/childFolders`);
    childFolderUris.add(`https://graph.microsoft.com/v1.0/me/mailFolders/${encodedId}/childFolders`);
  }

  for (const uri of childFolderUris) {
    const existingId = await tryGetChildFolderId(uri, folderName, errors);
    if (existingId) {
      return { folderId: existingId, note: null };
    }

    const createdId = await tryCreateChildFolder(uri, folderName, errors);
    if (createdId) {
      return { folderId: createdId, note: null };
    }

    const existingAfterCreateId = await tryGetChildFolderId(uri, folderName, errors);
    if (existingAfterCreateId) {
      return { folderId: existingAfterCreateId, note: null };
    }
  }

  if (errors.length > 0 && errors.every((error) => isDlpBlockedText(error))) {
    return {
      folderId: null,
      note: `Automatic folder creation is blocked by tenant DLP policy. Create Drafts/${folderName} manually once, then rerun.`,
    };
  }

  throw new Error(
    `Unable to ensure Drafts/${folderName} exists. ${errors.join(' | ') || 'No folder APIs succeeded.'}`,
  );
}

async function resolveDraftsFolder(errors: string[]): Promise<MailFolderItem | null> {
  const candidateUris = [
    '/me/mailFolders/drafts',
    '/me/mailFolders/Drafts',
    'https://graph.microsoft.com/v1.0/me/mailFolders/drafts',
    'https://graph.microsoft.com/v1.0/me/mailFolders/Drafts',
  ];

  for (const uri of candidateUris) {
    const response = await Office365OutlookService.HttpRequest(uri, 'GET');
    if (!response.success) {
      errors.push(`Get Drafts folder failed for ${uri}: ${toOperationErrorMessage(response.error)}`);
      continue;
    }

    const folder = extractMailFolder(response.data);
    if (folder?.id) {
      return folder;
    }

    errors.push(`Get Drafts folder returned no folder id for ${uri}.`);
  }

  return null;
}

async function tryGetChildFolderId(uri: string, folderName: string, errors: string[]): Promise<string | null> {
  const response = await Office365OutlookService.HttpRequest(uri, 'GET');
  if (!response.success) {
    errors.push(`List child folders failed for ${uri}: ${toOperationErrorMessage(response.error)}`);
    return null;
  }

  const existing = extractMailFolderByName(response.data, folderName);
  if (existing?.id) {
    return existing.id;
  }

  return null;
}

async function tryCreateChildFolder(uri: string, folderName: string, errors: string[]): Promise<string | null> {
  const createBody = JSON.stringify({ displayName: folderName });
  const response = await Office365OutlookService.HttpRequest(uri, 'POST', createBody, 'application/json');
  if (!response.success) {
    errors.push(`Create child folder failed for ${uri}: ${toOperationErrorMessage(response.error)}`);
    return null;
  }

  const createdFolder = extractMailFolder(response.data);
  if (createdFolder?.id) {
    return createdFolder.id;
  }

  errors.push(`Create child folder succeeded for ${uri} but returned no folder id.`);
  return null;
}

async function listDraftMessages(
  folderPath: string,
): Promise<{ success: boolean; items: DraftMessageItem[] }> {
  const pageSize = 100;
  let skip = 0;
  const idSet = new Set<string>();
  const items: DraftMessageItem[] = [];

  while (true) {
    const response = await Office365OutlookService.GetEmails(
      folderPath,
      false,
      false,
      undefined,
      pageSize,
      skip,
    );

    if (!response.success) {
      return { success: false, items: [] };
    }

    const page = Array.isArray(response.data) ? response.data : [];
    if (page.length === 0) {
      break;
    }

    for (const message of page) {
      const id = asOptionalString((message as { Id?: unknown }).Id);
      if (!id || idSet.has(id)) {
        continue;
      }
      idSet.add(id);

      const subject = asOptionalString((message as { Subject?: unknown }).Subject) || '(No Subject)';
      items.push({ id, subject });
    }

    if (page.length < pageSize) {
      break;
    }

    skip += page.length;
  }

  return {
    success: true,
    items,
  };
}

async function moveDraftToFolder(
  messageId: string,
  folderId: string | null,
  targetFolderPath: string,
  targetFolderName: string,
  folderResolutionNote?: string | null,
): Promise<string> {
  const moveErrors: string[] = [];

  if (folderId) {
    const body = JSON.stringify({ destinationId: folderId });
    const response = await Office365OutlookService.HttpRequest(
      `/me/messages/${encodeURIComponent(messageId)}/move`,
      'POST',
      body,
      'application/json',
    );

    if (response.success) {
      const movedId = extractMessageIdFromUnknown(response.data);
      return movedId || messageId;
    }
    moveErrors.push(`Graph move by folderId failed: ${toOperationErrorMessage(response.error)}`);
  } else {
    if (folderResolutionNote) {
      moveErrors.push(folderResolutionNote);
    }
    moveErrors.push('Manager Report folder id could not be resolved.');
  }

  const fallbackPaths = [targetFolderPath, `${DRAFT_FOLDER_PATH_PREFIX}\\${targetFolderName}`, targetFolderName];
  for (const folderPath of fallbackPaths) {
    const moveV2Response = await Office365OutlookService.MoveV2(messageId, folderPath);
    if (moveV2Response.success) {
      const movedId = asOptionalString((moveV2Response.data as { id?: unknown }).id);
      return movedId || messageId;
    }
    moveErrors.push(`MoveV2(${folderPath}) failed: ${toOperationErrorMessage(moveV2Response.error)}`);

    const moveResponse = await Office365OutlookService.Move(messageId, folderPath);
    if (moveResponse.success) {
      const movedId = asOptionalString((moveResponse.data as { Id?: unknown }).Id);
      return movedId || messageId;
    }
    moveErrors.push(`Move(${folderPath}) failed: ${toOperationErrorMessage(moveResponse.error)}`);
  }

  const actionableMoveErrors = moveErrors.filter(
    (error) =>
      error !== 'Manager Report folder id could not be resolved.' &&
      error !== folderResolutionNote,
  );

  if (
    folderResolutionNote &&
    actionableMoveErrors.length > 0 &&
    actionableMoveErrors.every((error) => isMissingFolderText(error))
  ) {
    throw new ManualDraftFolderRequiredError(targetFolderPath, folderResolutionNote);
  }

  throw new Error(
    `Unable to move draft into ${targetFolderPath}. ${moveErrors.join(' | ')}`,
  );
}

async function markMessageAsUnread(messageId: string): Promise<void> {
  const markErrors: string[] = [];

  const markV2 = await Office365OutlookService.MarkAsRead_V2(messageId, undefined, { isRead: false });
  if (markV2.success) {
    return;
  }
  markErrors.push(`MarkAsRead_V2 failed: ${toOperationErrorMessage(markV2.error)}`);

  const markV3 = await Office365OutlookService.MarkAsRead_V3(messageId, undefined, { isRead: false });
  if (markV3.success) {
    return;
  }
  markErrors.push(`MarkAsRead_V3 failed: ${toOperationErrorMessage(markV3.error)}`);

  const patchResponse = await Office365OutlookService.HttpRequest(
    `/me/messages/${encodeURIComponent(messageId)}`,
    'PATCH',
    JSON.stringify({ isRead: false }),
    'application/json',
  );
  if (patchResponse.success) {
    return;
  }
  markErrors.push(`Graph PATCH failed: ${toOperationErrorMessage(patchResponse.error)}`);

  throw new Error(`Draft saved but could not mark as unread. ${markErrors.join(' | ')}`);
}

async function deleteDraftIfExists(messageId: string): Promise<void> {
  const deleteV2Response = await Office365OutlookService.DeleteEmail_V2(messageId);
  if (deleteV2Response.success) {
    return;
  }

  const deleteResponse = await Office365OutlookService.DeleteEmail(messageId);
  if (deleteResponse.success) {
    return;
  }
}

function extractMailFolderByName(data: unknown, folderName: string): MailFolderItem | null {
  if (!data || typeof data !== 'object') {
    return null;
  }

  const maybeValue = (data as { value?: unknown }).value;
  if (!Array.isArray(maybeValue)) {
    return null;
  }

  for (const item of maybeValue) {
    if (!item || typeof item !== 'object') {
      continue;
    }

    const folder = item as MailFolderItem;
    if (folder.displayName === folderName) {
      return folder;
    }
  }

  return null;
}

function extractMailFolder(data: unknown): MailFolderItem | null {
  if (!data || typeof data !== 'object') {
    return null;
  }

  const folder = data as MailFolderItem;
  if (asOptionalString(folder.id)) {
    return folder;
  }

  return null;
}

function resolveTemplate(template: string, managerName: string, fiscalYear: string): string {
  return template
    .replaceAll('{manager_name}', managerName)
    .replaceAll('{fiscal_year}', fiscalYear);
}

function plainTextToHtml(text: string): string {
  const escaped = escapeHtml(text);
  const paragraphs = escaped
    .split(/\n{2,}/)
    .map((block) => block.replace(/\n/g, '<br>'))
    .map((block) => `<p>${block}</p>`)
    .join('');

  return `
    <html>
      <body style="font-family: Calibri, Segoe UI, sans-serif; font-size: 11pt; color: #1f2937;">
        ${paragraphs}
      </body>
    </html>
  `;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function asOptionalString(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null;
  }
  const trimmed = value.trim();
  return trimmed || null;
}

function toErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  return String(error);
}

function toOperationErrorMessage(error: unknown): string {
  if (!error) {
    return 'unknown connector error';
  }
  if (error instanceof Error) {
    return error.message;
  }
  if (typeof error === 'object') {
    const asRecord = error as Record<string, unknown>;
    const message =
      asOptionalString(asRecord.message) ||
      asOptionalString(asRecord.Message) ||
      asOptionalString(asRecord.error) ||
      asOptionalString(asRecord.Error);
    if (message) {
      return message;
    }
    try {
      return JSON.stringify(error);
    } catch {
      // fall through to String conversion
    }
  }

  const text = String(error).trim();
  return text || 'unknown connector error';
}

function isDlpBlockedText(value: string): boolean {
  return value.includes('Request blocked due to DLP policies') && value.includes('HttpRequest');
}

function isMissingFolderText(value: string): boolean {
  return value.includes('Specified folder') && value.includes('does not exist');
}

export function isManualDraftFolderRequiredError(error: unknown): error is ManualDraftFolderRequiredError {
  return error instanceof ManualDraftFolderRequiredError;
}

function extractMessageIdFromUnknown(data: unknown): string | null {
  if (!data || typeof data !== 'object') {
    return null;
  }

  const asRecord = data as Record<string, unknown>;
  return asOptionalString(asRecord.id) || asOptionalString(asRecord.Id);
}

function yieldToUi(): Promise<void> {
  return new Promise((resolve) => {
    window.setTimeout(resolve, 0);
  });
}
