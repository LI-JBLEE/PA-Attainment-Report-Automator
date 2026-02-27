import { useMemo, useState } from 'react';
import './App.css';
import {
  buildReportsZip,
  cleanDisplayName,
  generateManagerReports,
  parseAttainmentWorkbook,
  parseSalesCompWorkbook,
  type GenerationResult,
  type ManagerReportRecord,
} from './lib/reportEngine';
import {
  createOutlookDrafts,
  DEFAULT_EMAIL_BODY_TEMPLATE,
  DEFAULT_EMAIL_SUBJECT_TEMPLATE,
  sendAllDraftEmailsFromFolder,
} from './lib/outlookDrafts';

type AppError = string | null;

interface ParsedAttainmentInfo {
  rows: import('./lib/reportEngine').AttainmentRow[];
  managerCount: number;
  fiscalYear: string;
  availableRegions: string[];
  fileName: string;
}

interface ParsedSalesCompInfo {
  emailMap: Record<string, string>;
  count: number;
  fileName: string;
}

type UploadZone = 'attainment' | 'sales';

function App() {
  const [attainmentInfo, setAttainmentInfo] = useState<ParsedAttainmentInfo | null>(null);
  const [salesCompInfo, setSalesCompInfo] = useState<ParsedSalesCompInfo | null>(null);
  const [selectedRegions, setSelectedRegions] = useState<string[]>([]);
  const [generationResult, setGenerationResult] = useState<GenerationResult | null>(null);
  const [zipBlob, setZipBlob] = useState<Blob | null>(null);

  const [isGenerating, setIsGenerating] = useState(false);
  const [isDrafting, setIsDrafting] = useState(false);
  const [isSendingDrafts, setIsSendingDrafts] = useState(false);
  const [showSendConfirmModal, setShowSendConfirmModal] = useState(false);
  const [progressText, setProgressText] = useState('');
  const [progressValue, setProgressValue] = useState(0);

  const [draftRegionFilter, setDraftRegionFilter] = useState<string[]>([]);
  const [selectedManagerIds, setSelectedManagerIds] = useState<string[]>([]);
  const [searchText, setSearchText] = useState('');

  const [subjectTemplate, setSubjectTemplate] = useState(DEFAULT_EMAIL_SUBJECT_TEMPLATE);
  const [bodyTemplate, setBodyTemplate] = useState(DEFAULT_EMAIL_BODY_TEMPLATE);
  const [operationMessage, setOperationMessage] = useState('');
  const [error, setError] = useState<AppError>(null);
  const [dragOverZone, setDragOverZone] = useState<UploadZone | null>(null);

  const filesReady = Boolean(attainmentInfo && salesCompInfo);
  const regionOptions = attainmentInfo?.availableRegions || [];
  const reportRegionOptions = useMemo(() => {
    if (!generationResult) {
      return [];
    }
    return Array.from(new Set(generationResult.reports.map((report) => report.region))).sort((a, b) =>
      a.localeCompare(b),
    );
  }, [generationResult]);

  const selectedIdSet = useMemo(() => new Set(selectedManagerIds), [selectedManagerIds]);
  const selectedDraftRegionSet = useMemo(() => new Set(draftRegionFilter), [draftRegionFilter]);

  const filteredReports = useMemo(() => {
    if (!generationResult) {
      return [];
    }

    const query = searchText.trim().toLowerCase();
    return generationResult.reports.filter((report) => {
      const inRegion = draftRegionFilter.length > 0 && selectedDraftRegionSet.has(report.region);
      if (!inRegion) {
        return false;
      }
      if (!query) {
        return true;
      }
      return (
        report.displayName.toLowerCase().includes(query) ||
        report.region.toLowerCase().includes(query) ||
        (report.email || '').toLowerCase().includes(query)
      );
    });
  }, [draftRegionFilter.length, generationResult, searchText, selectedDraftRegionSet]);

  const selectedReports = useMemo(() => {
    if (!generationResult) {
      return [];
    }
    return generationResult.reports.filter(
      (report) =>
        selectedIdSet.has(report.id) &&
        draftRegionFilter.length > 0 &&
        selectedDraftRegionSet.has(report.region),
    );
  }, [draftRegionFilter.length, generationResult, selectedIdSet, selectedDraftRegionSet]);

  const selectedWithEmail = useMemo(
    () => selectedReports.filter((report) => Boolean(report.email)),
    [selectedReports],
  );

  const selectedWithoutEmail = selectedReports.length - selectedWithEmail.length;
  const activeFiscalYear = generationResult?.fiscalYear || attainmentInfo?.fiscalYear || 'FY26';
  const uploadsDisabled = isGenerating || isDrafting;

  const resetGenerationState = () => {
    setGenerationResult(null);
    setZipBlob(null);
    setDraftRegionFilter([]);
    setSelectedManagerIds([]);
    setSearchText('');
    setProgressText('');
    setProgressValue(0);
  };

  const handleResetAll = () => {
    setAttainmentInfo(null);
    setSalesCompInfo(null);
    setSelectedRegions([]);
    resetGenerationState();
    setSubjectTemplate(DEFAULT_EMAIL_SUBJECT_TEMPLATE);
    setBodyTemplate(DEFAULT_EMAIL_BODY_TEMPLATE);
    setOperationMessage('');
    setError(null);
    setShowSendConfirmModal(false);
  };

  const processAttainmentFile = async (file: File) => {
    setError(null);
    setOperationMessage('');

    try {
      const parsed = await parseAttainmentWorkbook(file);
      setAttainmentInfo({
        ...parsed,
        fileName: file.name,
      });
      setSelectedRegions(parsed.availableRegions);
      resetGenerationState();
    } catch (uploadError) {
      setError(toErrorMessage(uploadError));
    }
  };

  const processSalesCompFile = async (file: File) => {
    setError(null);
    setOperationMessage('');

    try {
      const parsed = await parseSalesCompWorkbook(file);
      setSalesCompInfo({
        ...parsed,
        fileName: file.name,
      });
      resetGenerationState();
    } catch (uploadError) {
      setError(toErrorMessage(uploadError));
    }
  };

  const handleAttainmentUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    await processAttainmentFile(file);
    event.target.value = '';
  };

  const handleSalesCompUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    await processSalesCompFile(file);
    event.target.value = '';
  };

  const handleUploadDragOver = (event: React.DragEvent<HTMLDivElement>, zone: UploadZone) => {
    if (uploadsDisabled) {
      return;
    }
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy';
    if (dragOverZone !== zone) {
      setDragOverZone(zone);
    }
  };

  const handleUploadDragLeave = (event: React.DragEvent<HTMLDivElement>, zone: UploadZone) => {
    if (uploadsDisabled) {
      return;
    }
    const currentTarget = event.currentTarget;
    const nextTarget = event.relatedTarget as Node | null;
    if (nextTarget && currentTarget.contains(nextTarget)) {
      return;
    }
    if (dragOverZone === zone) {
      setDragOverZone(null);
    }
  };

  const handleUploadDrop = async (event: React.DragEvent<HTMLDivElement>, zone: UploadZone) => {
    event.preventDefault();
    setDragOverZone(null);

    if (uploadsDisabled) {
      return;
    }

    const file = event.dataTransfer.files?.[0];
    if (!file) {
      return;
    }

    if (!file.name.toLowerCase().endsWith('.xlsx')) {
      setError('Only .xlsx files are supported.');
      return;
    }

    if (zone === 'attainment') {
      await processAttainmentFile(file);
      return;
    }
    await processSalesCompFile(file);
  };

  const toggleRegion = (region: string) => {
    setSelectedRegions((current) =>
      current.includes(region) ? current.filter((item) => item !== region) : [...current, region],
    );
  };

  const handleGenerateReports = async () => {
    if (!attainmentInfo || !salesCompInfo) {
      return;
    }

    setError(null);
    setOperationMessage('');
    setIsGenerating(true);
    setProgressValue(0);
    setProgressText('Starting report generation...');

    try {
      const generated = await generateManagerReports({
        rows: attainmentInfo.rows,
        emailMap: salesCompInfo.emailMap,
        selectedRegions,
        fiscalYear: attainmentInfo.fiscalYear,
        onProgress: (current, total, message) => {
          setProgressText(`[${current}/${total}] ${message}`);
          setProgressValue(total > 0 ? current / total : 0);
        },
      });

      const zip = await buildReportsZip(generated);
      setGenerationResult(generated);
      setZipBlob(zip);
      setDraftRegionFilter(Array.from(new Set(generated.reports.map((report) => report.region))));
      setSelectedManagerIds(generated.reports.map((report) => report.id));
      setProgressText(`Completed: ${generated.total} reports generated`);
      setProgressValue(1);
    } catch (generationError) {
      setError(toErrorMessage(generationError));
    } finally {
      setIsGenerating(false);
    }
  };

  const handleDownloadZip = () => {
    if (!zipBlob || !generationResult) {
      return;
    }
    const fileName = `Manager_Reports_${generationResult.fiscalYear}_${formatDateTimeCompact(new Date())}.zip`;
    downloadBlob(zipBlob, fileName);
  };

  const toggleDraftRegion = (region: string) => {
    setDraftRegionFilter((current) =>
      current.includes(region) ? current.filter((item) => item !== region) : [...current, region],
    );
  };

  const toggleManager = (id: string) => {
    setSelectedManagerIds((current) =>
      current.includes(id) ? current.filter((item) => item !== id) : [...current, id],
    );
  };

  const selectVisibleManagers = () => {
    const visibleIds = filteredReports.map((report) => report.id);
    setSelectedManagerIds((current) => Array.from(new Set([...current, ...visibleIds])));
  };

  const deselectVisibleManagers = () => {
    const visibleIds = new Set(filteredReports.map((report) => report.id));
    setSelectedManagerIds((current) => current.filter((id) => !visibleIds.has(id)));
  };

  const handleCreateDrafts = async () => {
    if (!generationResult) {
      return;
    }

    const targets = selectedReports.filter((report) => report.email);
    if (targets.length === 0) {
      setError('No selected managers with email addresses.');
      return;
    }

    setError(null);
    setOperationMessage('');
    setIsDrafting(true);
    setProgressValue(0);
    setProgressText('Creating Outlook drafts...');

    try {
      const result = await createOutlookDrafts(
        targets,
        {
          fiscalYear: activeFiscalYear,
          subjectTemplate,
          bodyTemplate,
          targetFolderName: 'Manager Report',
        },
        ({ current, total, managerName }) => {
          setProgressText(`[${current}/${total}] Drafting ${managerName}`);
          setProgressValue(total > 0 ? current / total : 0);
        },
      );

      const summary = `Drafts complete: ${result.created} created, ${result.failed} failed (folder: ${result.folderUsed}).`;
      const failurePreview = result.failures.slice(0, 3).join(' | ');
      setOperationMessage(failurePreview ? `${summary} Examples: ${failurePreview}` : summary);
    } catch (draftError) {
      setError(toErrorMessage(draftError));
    } finally {
      setIsDrafting(false);
    }
  };

  const executeSendDraftEmails = async () => {
    setError(null);
    setOperationMessage('');
    setIsSendingDrafts(true);
    setProgressValue(0);
    setProgressText('Loading drafts from Drafts/Manager Report...');

    try {
      const result = await sendAllDraftEmailsFromFolder('Manager Report', ({ current, total, subject }) => {
        setProgressText(`[${current}/${total}] Sending ${subject}`);
        setProgressValue(total > 0 ? current / total : 0);
      });

      if (result.totalFound === 0) {
        setOperationMessage(`No draft emails were found in ${result.folderPath}.`);
      } else {
        const summary = `Send complete: ${result.sent} sent, ${result.failed} failed (from ${result.folderPath}).`;
        const failurePreview = result.failures.slice(0, 3).join(' | ');
        setOperationMessage(failurePreview ? `${summary} Examples: ${failurePreview}` : summary);
      }
    } catch (sendError) {
      setError(toErrorMessage(sendError));
    } finally {
      setIsSendingDrafts(false);
    }
  };

  const handleSendDraftEmails = () => {
    if (isSendingDrafts || isDrafting || isGenerating) {
      return;
    }
    setShowSendConfirmModal(true);
  };

  const confirmSendDraftEmails = async () => {
    setShowSendConfirmModal(false);
    await executeSendDraftEmails();
  };

  const cancelSendDraftEmails = () => {
    if (isSendingDrafts) {
      return;
    }
    setShowSendConfirmModal(false);
  };

  return (
    <div className="app-shell">
      <header className="hero">
        <div className="hero-title-block">
          <h1>Attainment Report Generator</h1>
          <p>Sales Compensation Attainment Reporting Tool</p>
        </div>
      </header>

      <main className="content-grid">
        <div className="top-actions">
          <button
            className="ghost-btn"
            type="button"
            onClick={handleResetAll}
            disabled={isGenerating || isDrafting || isSendingDrafts}
          >
            Reset
          </button>
        </div>

        <section className="panel">
          <h2>Step 1 - Upload Source Files</h2>
          <div className="two-col">
            <div className="upload-card">
              <h3>Global Attainment Report</h3>
              <div
                className={`file-picker-row ${dragOverZone === 'attainment' ? 'drag-active' : ''}`}
                onDragOver={(event) => handleUploadDragOver(event, 'attainment')}
                onDragLeave={(event) => handleUploadDragLeave(event, 'attainment')}
                onDrop={(event) => handleUploadDrop(event, 'attainment')}
              >
                <input
                  id="attainment-file-input"
                  className="file-input-hidden"
                  type="file"
                  accept=".xlsx"
                  onChange={handleAttainmentUpload}
                  disabled={uploadsDisabled}
                />
                <label
                  htmlFor="attainment-file-input"
                  className={`file-select-btn ${uploadsDisabled ? 'disabled' : ''}`}
                >
                  Choose File
                </label>
                <span className="file-load-status">
                  {attainmentInfo ? attainmentInfo.fileName : 'Drop .xlsx file here or choose file'}
                </span>
              </div>
            </div>

            <div className="upload-card">
              <h3>Sales Compensation Report</h3>
              <div
                className={`file-picker-row ${dragOverZone === 'sales' ? 'drag-active' : ''}`}
                onDragOver={(event) => handleUploadDragOver(event, 'sales')}
                onDragLeave={(event) => handleUploadDragLeave(event, 'sales')}
                onDrop={(event) => handleUploadDrop(event, 'sales')}
              >
                <input
                  id="sales-comp-file-input"
                  className="file-input-hidden"
                  type="file"
                  accept=".xlsx"
                  onChange={handleSalesCompUpload}
                  disabled={uploadsDisabled}
                />
                <label
                  htmlFor="sales-comp-file-input"
                  className={`file-select-btn ${uploadsDisabled ? 'disabled' : ''}`}
                >
                  Choose File
                </label>
                <span className="file-load-status">
                  {salesCompInfo ? salesCompInfo.fileName : 'Drop .xlsx file here or choose file'}
                </span>
              </div>
            </div>
          </div>
        </section>

        <section className="panel">
          <h2>Step 2 - Generate Manager Reports</h2>
          {!filesReady ? <p className="hint">Upload both source files to continue.</p> : null}

          {filesReady ? (
            <>
              <div className="controls-row">
                <button
                  className="ghost-btn"
                  type="button"
                  onClick={() => setSelectedRegions(regionOptions)}
                  disabled={regionOptions.length === 0 || isGenerating}
                >
                  Select All Regions
                </button>
                <button
                  className="ghost-btn"
                  type="button"
                  onClick={() => setSelectedRegions([])}
                  disabled={selectedRegions.length === 0 || isGenerating}
                >
                  Clear Region Selection
                </button>
              </div>

              <div className="pill-wrap step2-region-list">
                {regionOptions.map((region) => {
                  const active = selectedRegions.includes(region);
                  return (
                    <button
                      key={region}
                      className={active ? 'pill active' : 'pill'}
                      type="button"
                      onClick={() => toggleRegion(region)}
                      disabled={isGenerating}
                    >
                      {region}
                    </button>
                  );
                })}
              </div>

              <div className="controls-row">
                <button
                  className="primary-btn"
                  type="button"
                  onClick={handleGenerateReports}
                  disabled={isGenerating || selectedRegions.length === 0}
                >
                  {isGenerating ? 'Generating...' : 'Generate Reports'}
                </button>

                <button
                  className="ghost-btn"
                  type="button"
                  onClick={handleDownloadZip}
                  disabled={!zipBlob || isGenerating}
                >
                  Download Reports (.zip)
                </button>
              </div>
            </>
          ) : null}

          {(isGenerating || isDrafting || progressText) && (
            <div className="progress-shell" aria-live="polite">
              <div className="progress-label">{progressText}</div>
              <div className="progress-track">
                <div className="progress-fill" style={{ width: `${Math.round(progressValue * 100)}%` }} />
              </div>
            </div>
          )}

          {generationResult ? (
            <div className="stats-row">
              <Stat label="Fiscal Year" value={generationResult.fiscalYear} />
              <Stat label="Reports" value={String(generationResult.total)} />
              <Stat label="Regions" value={String(Object.keys(generationResult.regionCounts).length)} />
            </div>
          ) : null}
        </section>

        {generationResult ? (
          <section className="panel">
            <h2>Step 3 - Create Outlook Drafts</h2>

            <div className="three-col">
              <div className="filter-col">
                <h3>Region Filter</h3>
                <div className="pill-wrap compact">
                  {reportRegionOptions.map((region) => {
                    const active = draftRegionFilter.includes(region);
                    return (
                      <button
                        key={region}
                        className={active ? 'pill active' : 'pill'}
                        type="button"
                        onClick={() => toggleDraftRegion(region)}
                        disabled={isDrafting}
                      >
                        {region}
                      </button>
                    );
                  })}
                </div>

                <input
                  className="text-input"
                  type="search"
                  placeholder="Search manager, region, or email"
                  value={searchText}
                  onChange={(event) => setSearchText(event.target.value)}
                  disabled={isDrafting}
                />

                <div className="controls-row">
                  <button className="ghost-btn" type="button" onClick={selectVisibleManagers} disabled={isDrafting}>
                    Select Visible
                  </button>
                  <button
                    className="ghost-btn"
                    type="button"
                    onClick={deselectVisibleManagers}
                    disabled={isDrafting}
                  >
                    Deselect Visible
                  </button>
                </div>

                <p className="hint">
                  Selected in current region filter: {selectedReports.length} | With email:{' '}
                  {selectedWithEmail.length} | Without email: {selectedWithoutEmail}
                </p>
              </div>

              <div className="manager-col">
                <h3>Managers</h3>
                <div className="manager-list" role="list">
                  {filteredReports.map((report) => (
                    <ManagerRow
                      key={report.id}
                      report={report}
                      checked={selectedIdSet.has(report.id)}
                      onToggle={toggleManager}
                      disabled={isDrafting}
                    />
                  ))}
                </div>
              </div>

              <div className="template-col">
                <h3>Email Template</h3>
                <label className="field-label" htmlFor="subject-template">
                  Subject
                </label>
                <input
                  id="subject-template"
                  className="text-input"
                  type="text"
                  value={subjectTemplate}
                  onChange={(event) => setSubjectTemplate(event.target.value)}
                  disabled={isDrafting}
                />

                <label className="field-label" htmlFor="body-template">
                  Body (plain text)
                </label>
                <textarea
                  id="body-template"
                  className="text-area"
                  value={bodyTemplate}
                  onChange={(event) => setBodyTemplate(event.target.value)}
                  disabled={isDrafting}
                />

                <p className="hint">Placeholders: {'{manager_name}'}, {'{fiscal_year}'}</p>
                <div className="template-actions">
                  <button
                    className="primary-btn"
                    type="button"
                    onClick={handleCreateDrafts}
                    disabled={isDrafting || isSendingDrafts || selectedWithEmail.length === 0}
                  >
                    {isDrafting ? 'Creating Drafts...' : 'Create Outlook Drafts'}
                  </button>
                  <button
                    className="ghost-btn"
                    type="button"
                    onClick={handleSendDraftEmails}
                    disabled={isSendingDrafts || isDrafting || isGenerating}
                  >
                    {isSendingDrafts ? 'Sending Draft Emails...' : 'Send Draft Emails'}
                  </button>
                </div>
              </div>
            </div>
          </section>
        ) : null}

        {operationMessage ? <div className="message success">{operationMessage}</div> : null}
        {error ? <div className="message error">{error}</div> : null}
      </main>

      {showSendConfirmModal ? (
        <div className="modal-backdrop" role="dialog" aria-modal="true" aria-labelledby="send-drafts-title">
          <div className="confirm-modal">
            <h3 id="send-drafts-title">Send Draft Emails</h3>
            <p>Send all draft emails in Drafts/Manager Report folder now?</p>
            <p>This action immediately sends emails</p>
            <div className="confirm-actions">
              <button
                className="ghost-btn"
                type="button"
                onClick={cancelSendDraftEmails}
                disabled={isSendingDrafts}
              >
                Cancel
              </button>
              <button
                className="primary-btn"
                type="button"
                onClick={confirmSendDraftEmails}
                disabled={isSendingDrafts}
              >
                Send
              </button>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}

function Stat(props: { label: string; value: string }) {
  return (
    <article className="stat-card">
      <span>{props.label}</span>
      <strong>{props.value}</strong>
    </article>
  );
}

function ManagerRow(props: {
  report: ManagerReportRecord;
  checked: boolean;
  onToggle: (id: string) => void;
  disabled: boolean;
}) {
  return (
    <label className="manager-row" role="listitem">
      <input
        type="checkbox"
        checked={props.checked}
        onChange={() => props.onToggle(props.report.id)}
        disabled={props.disabled}
      />
      <div className="manager-row-content">
        <div className="manager-top">
          <span>{cleanDisplayName(props.report.fullName)}</span>
          <small>{props.report.region}</small>
        </div>
        <div className={props.report.email ? 'email ok' : 'email missing'}>
          {props.report.email || 'No email mapped'}
        </div>
      </div>
    </label>
  );
}

function formatDateTimeCompact(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hour = String(date.getHours()).padStart(2, '0');
  const minute = String(date.getMinutes()).padStart(2, '0');
  const second = String(date.getSeconds()).padStart(2, '0');
  return `${year}${month}${day}_${hour}${minute}${second}`;
}

function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function toErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  return String(error);
}

export default App;
