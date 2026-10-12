import { useCloudWorkspace } from '../../../lib/contexts/CloudWorkspaceContext';
import { useEffect, useMemo, useRef, useState } from 'react';
import { supabase, isSupabaseConfigured } from '../../../../lib/supabaseClient';
import { Dialog, DialogContent, DialogDescription, DialogHeader, DialogTitle } from '../../../components/ui/dialog';
import { HistorySaveConflictError, loadLatestVersion, loadTransactionCheckSource, replaceTransactionVersionsAtomically, saveVersion, type TransactionCheckSource, type TransactionVersion } from '../../../lib/transaction-history-store';
import { compareAccountsAcrossHistory, formatHistoryDate, selectPeriodRows, visibleHistoricalComparisons, type TransactionRow, type HistoricalAccountComparison } from '../../../lib/utils/transaction-history';
import { formatResolutionPeriods, type TransactionHistoryResolutionField } from '../../../lib/utils/transaction-history-resolution';
import { sameTransactionSnapshot } from '../../../lib/utils/transaction-snapshot';
import { conflictingLocalTransactionPeriods, type LocalTransactionMonths } from '../../../lib/utils/transaction-local-history';
import { createIdentityResolutionBuilder, IDENTITY_FIELDS, identityResolutionTargets, planIdentityResolution } from '../../../lib/utils/transaction-identity-resolution';
import { syncTransactionEmployeesToSupabase } from '../../../lib/utils/transaction-employee-sync';
import { TransactionHistoryTable } from './TransactionHistoryTable';

interface Props {
  rows: TransactionRow[];
  month: string;
  showReport: boolean;
  onOpenReport: () => void;
  onApplyVersions: (versions: TransactionVersion[]) => void;
  localMonths?: LocalTransactionMonths;
  hasPendingEdits: boolean;
  syncRevision?: number;
  onReportStateChange?: (hasExceptions: boolean, viewingSource: boolean) => void;
}
interface Report extends TransactionCheckSource {
  context: object;
  comparisons: HistoricalAccountComparison[];
}
interface IdentityDecision {
  context: object;
  rowIndex: number;
  field: TransactionHistoryResolutionField;
  optionKey: string;
  selectedVersionIds: string[];
}

function employeeSyncMessage(summary: Awaited<ReturnType<typeof syncTransactionEmployeesToSupabase>>): string {
  const skipped = summary.skipped
    ? ` Bỏ qua ID mâu thuẫn: ${summary.skippedDocumentIds.join(', ')}.`
    : '';
  return `Đã cập nhật nhan_vien: ${summary.synced} dòng (${summary.inserted} mới, ${summary.updated} cập nhật).${skipped}`;
}

const markedDocumentId = (value: string, syncNote: string) => (
  value ? `${value}${syncNote ? '!' : ''}` : '—'
);

export function TransactionHistoryPanel({ rows, month, showReport, onOpenReport, onApplyVersions, localMonths, hasPendingEdits, syncRevision = 0, onReportStateChange }: Props) {
  const cloud = useCloudWorkspace();
  const [userId, setUserId] = useState('');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loginOpen, setLoginOpen] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('');
  const [report, setReport] = useState<Report | null>(null);
  const [decision, setDecision] = useState<IdentityDecision | null>(null);
  const [page, setPage] = useState(1);
  const [defaultBank, setDefaultBank] = useState('VCB');
  const operation = useRef(false);
  const authIdentity = useRef('');
  const retry = useRef<{context: object; requestId: string} | null>(null);
  // AppData updates are immutable. Include retained months so a history edit
  // also invalidates reports and in-flight reads without serializing every row.
  const context = useMemo(() => ({month, rows, userId, hasPendingEdits, defaultBank, syncRevision, localMonths}), [month, rows, userId, hasPendingEdits, defaultBank, syncRevision, localMonths]);
  const currentContext = useRef(context);
  useEffect(() => { currentContext.current = context; }, [context]);
  useEffect(() => {
    const openSettings = () => setSettingsOpen(true);
    window.addEventListener('open-transaction-settings', openSettings);
    return () => window.removeEventListener('open-transaction-settings', openSettings);
  }, []);
  useEffect(() => {
    if (!isSupabaseConfigured()) return;
    let live = true;
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      const nextId = session?.user.id || '';
      if (live && nextId !== authIdentity.current) {
        authIdentity.current = nextId;
        setUserId(nextId);
        setReport(null);
        retry.current = null;
      }
    });
    return () => { live = false; subscription.unsubscribe(); };
  }, []);
  const visibleReport = report?.context === context ? report : null;
  const localMatchesCloud = useMemo(() => Boolean(visibleReport
    && sameTransactionSnapshot(rows, visibleReport.currentVersion.rows, month)), [visibleReport, rows, month]);
  const buildResolution = useMemo(() => visibleReport
    ? createIdentityResolutionBuilder(visibleReport.currentVersion, visibleReport.versions, defaultBank)
    : null, [visibleReport, defaultBank]);
  const visibleGroup = useMemo(() => decision?.context === context && buildResolution
    ? buildResolution(decision.rowIndex, decision.field) : null, [decision, context, buildResolution]);
  const targets = visibleGroup && decision ? identityResolutionTargets(visibleGroup, decision.optionKey) : [];
  const selectedTargets = targets.filter(target => decision?.selectedVersionIds.includes(target.versionId));
  const chosenOption = visibleGroup?.options.find(option => option.key === decision?.optionKey);
  const chosenField = IDENTITY_FIELDS.find(item => item.field === decision?.field);
  const comparisons = useMemo(() => visibleReport
    ? visibleHistoricalComparisons(visibleReport.comparisons)
    : [], [visibleReport]);
  const exceptions = useMemo(
    () => comparisons.filter(row => row.issues.length > 0),
    [comparisons],
  );
  const hasExceptions = exceptions.length > 0;
  useEffect(() => {
    onReportStateChange?.(showReport && hasExceptions, false);
  }, [showReport, hasExceptions, onReportStateChange]);
  const resolutionGroups = useMemo(() => new Map(exceptions.map(row => [row.currentRowIndex,
    IDENTITY_FIELDS.map(item => ({...item, group: buildResolution?.(row.currentRowIndex, item.field)}))
      .filter(item => item.group),
  ])), [exceptions, buildResolution]);
  const pageCount = Math.max(1, Math.ceil(exceptions.length / 25));
  const activePage = Math.min(page, pageCount);
  const buttonClass = 'rounded-full border border-primary/20 bg-primary/5 px-3 py-1.5 text-xs font-semibold whitespace-nowrap hover:bg-primary/10 disabled:opacity-50 active:scale-[0.98]';
  const resolutionButtonClass = 'rounded-full border border-primary/25 bg-primary/10 px-2.5 py-1 text-[10px] font-bold whitespace-nowrap text-foreground hover:bg-primary/20 disabled:opacity-50 active:scale-[0.98]';

  function requireUnchangedContext(started: object) {
    if (currentContext.current !== started) throw new HistorySaveConflictError('Transaction hoặc tháng đang chọn đã đổi. Hãy kiểm tra lại.');
  }

  async function refreshReport(started: object) {
    const source = await loadTransactionCheckSource(supabase, month);
    if (currentContext.current !== started) return false;
    const conflicts = conflictingLocalTransactionPeriods(source, rows, month, localMonths);
    if (conflicts.length) {
      throw new HistorySaveConflictError(
        `Batch Payment đã lưu trên máy khác dữ liệu Supabase ở tháng ${formatResolutionPeriods(conflicts)}. Mở từng tháng, kiểm tra bản đúng rồi bấm Lưu tháng để đưa bản trên máy lên Supabase, hoặc Cài đặt → Tải tháng để dùng bản Supabase. Sau đó Check STK & ID lại.`,
      );
    }
    setReport({
      context: started,
      ...source,
      comparisons: compareAccountsAcrossHistory(source.currentVersion.rows, source.versions, defaultBank),
    });
    setPage(1);
    return true;
  }

  async function run(action: 'save' | 'check' | 'load') {
    if (operation.current || hasPendingEdits) return;
    operation.current = true;
    setBusy(true);
    setMessage('');
    setReport(null);
    setDecision(null);
    const started = context;
    const operationOwner = authIdentity.current;
    let savedMonthId: string | null = null;
    const stillViewingMonth = () => currentContext.current.month === month && authIdentity.current === operationOwner;
    try {
      if (!isSupabaseConfigured()) throw new Error('Chưa cấu hình Supabase URL và publishable/anon key.');
      if (action === 'save') {
        const selected = selectPeriodRows(rows, month);
        if (retry.current?.context !== started) retry.current = {context: started, requestId: crypto.randomUUID()};
        const id = await saveVersion(supabase, month, selected, retry.current.requestId);
        retry.current = null;
        const latest = await loadLatestVersion(supabase, month);
        if (!latest || latest.id !== id) throw new HistorySaveConflictError('Tháng vừa có phiên bản mới hơn trên Supabase. Bấm Check STK & ID để lấy dữ liệu mới nhất.');
        if (!sameTransactionSnapshot(selected, latest.rows, month)) throw new HistorySaveConflictError('Dữ liệu đọc lại từ Supabase chưa khớp Transaction vừa lưu. Bấm Lưu tháng lại trước khi Check STK & ID.');
        requireUnchangedContext(started);
        savedMonthId = id;
        onApplyVersions([latest]);
        // Monthly Transaction and the complete working data must both be durable.
        if (cloud?.authenticated) await cloud.flush();
        let employeeMessage: string;
        try { employeeMessage = employeeSyncMessage(await syncTransactionEmployeesToSupabase(supabase, latest.rows)); }
        catch (error) { employeeMessage = `Tháng đã lưu; danh mục nhân viên chưa cập nhật: ${error instanceof Error ? error.message : 'Lỗi kết nối'}`; }
        if (stillViewingMonth()) {
          setReport(null);
          setMessage(`Đã lưu tháng ${month} lên Supabase: ${selected.length} dòng (#${id}). ${employeeMessage} Check STK & ID sẽ tải lại phiên bản mới nhất.`);
        }
      } else if (action === 'load') {
        const latest = await loadLatestVersion(supabase, month);
        if (!latest) throw new Error('Tháng này chưa có dữ liệu trên Supabase.');
        requireUnchangedContext(started);
        onApplyVersions([latest]);
        if (cloud?.authenticated) await cloud.flush();
        if (stillViewingMonth()) setMessage(`Đã tải tháng ${month} từ Supabase (#${latest.id}). Bấm Check STK & ID để đối chiếu.`);
      } else if (await refreshReport(started)) {
        onOpenReport();
      }
    } catch (error) {
      if (error instanceof HistorySaveConflictError) retry.current = null;
      if (stillViewingMonth()) setMessage(`${savedMonthId ? `Transaction tháng ${month} đã lưu (#${savedMonthId}); đồng bộ các bảng chưa hoàn tất. ` : ''}${error instanceof Error ? error.message : 'Không thể truy cập kho Transaction.'}`);
    } finally {
      operation.current = false;
      setBusy(false);
    }
  }

  async function login() {
    if (operation.current) return;
    operation.current = true;
    setBusy(true);
    try {
      if (!isSupabaseConfigured()) throw new Error('Chưa cấu hình Supabase.');
      const { error } = await supabase.auth.signInWithPassword({ email, password });
      if (error) throw error;
      setLoginOpen(false);
      setMessage('Đã đăng nhập kho.');
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'Đăng nhập thất bại.');
    } finally {
      setPassword('');
      operation.current = false;
      setBusy(false);
    }
  }

  async function resolveIdentity() {
    if (!visibleReport || !visibleGroup || !decision || !chosenOption
      || operation.current || hasPendingEdits || !localMatchesCloud || !selectedTargets.length) return;
    operation.current = true;
    setBusy(true);
    setMessage('');
    const started = context;
    let committed = false;
    const periods = selectedTargets.map(target => target.period);
    try {
      const plan = planIdentityResolution([visibleReport.currentVersion, ...visibleReport.versions],
        visibleGroup, decision.optionKey, decision.selectedVersionIds, new Date().toISOString());
      requireUnchangedContext(started);
      const saved = await replaceTransactionVersionsAtomically(supabase, visibleReport, plan);
      committed = true;
      requireUnchangedContext(started);
      const fresh = await loadTransactionCheckSource(supabase, month);
      requireUnchangedContext(started);
      const verifiedVersions: TransactionVersion[] = [];
      for (const change of plan) {
        const period = change.version.period.slice(0, 7);
        const version = [fresh.currentVersion, ...fresh.versions].find(item => item.period.slice(0, 7) === period);
        const receipt = saved.find(item => item.period.slice(0, 7) === period);
        if (!version || version.id !== receipt?.id || !sameTransactionSnapshot(version.rows, change.rows, period)) {
          throw new HistorySaveConflictError(`Tháng ${formatResolutionPeriods([period])} chưa đọc lại được hoặc vừa có thay đổi mới.`);
        }
        verifiedVersions.push(version);
      }
      let employeeMessage: string;
      try { employeeMessage = employeeSyncMessage(await syncTransactionEmployeesToSupabase(supabase, fresh.currentVersion.rows)); }
      catch (error) { employeeMessage = `Các tháng đã lưu; danh mục nhân viên chưa cập nhật: ${error instanceof Error ? error.message : 'Lỗi kết nối'}`; }
      requireUnchangedContext(started);
      // Every changed month must be reflected locally, otherwise switching back
      // can display/re-save an obsolete snapshot over the resolved cloud data.
      onApplyVersions(verifiedVersions);
      setReport(null);
      setDecision(null);
      setMessage(`Đã đồng bộ ${chosenField?.label}: ${chosenOption.value} theo ${formatResolutionPeriods([chosenOption.period])}. Đã lưu tháng ${formatResolutionPeriods(periods)} lên Supabase và cập nhật Batch Payment các tháng tương ứng. ${employeeMessage} Bấm Check STK & ID để đối chiếu lại.`);
    } catch (error) {
      if (currentContext.current === started) {
        setReport(null);
        setDecision(null);
        setMessage(`${committed ? `Đã lưu tháng ${formatResolutionPeriods(periods)}. ` : ''}${error instanceof Error ? error.message : 'Không thể xác nhận kết quả lưu.'} Bấm Check STK & ID để tải trạng thái mới nhất.`);
      }
    } finally {
      operation.current = false;
      setBusy(false);
    }
  }

  async function exportReport() {
    try {
      if (!visibleReport) return;
      const XLSX = await import('xlsx');
      const data = exceptions.map(row => ({
        'Tháng này': month,
        'Nguồn tháng hiện tại': `Supabase #${visibleReport.currentVersion.id} · ${visibleReport.currentVersion.created_at}`,
        'Ngân hàng kiểm tra': row.bankCheck?.bank || 'Chưa xác định',
        'Ngân hàng mặc định': row.bankCheck?.bankAssumed ? 'Có' : 'Không',
        'Phạm vi xác minh': 'Đối chiếu dữ liệu; chưa xác minh với ngân hàng',
        'Nguồn đối chiếu': row.sources.map(source => `${source.period} · #${source.versionId} · ${formatHistoryDate(source.createdAt)} · Document ID ${markedDocumentId(source.documentId, source.documentIdSyncNote)} · STK ${source.account} · ${source.name}${source.documentIdSyncNote ? ` · ! Đồng bộ theo ${source.documentIdSyncNote}` : ''}`).join('\n'),
        'Document ID lịch sử': markedDocumentId(row.previousDocumentId, row.sources.map(source => source.documentIdSyncNote).find(Boolean) || ''),
        'Document ID hiện tại': markedDocumentId(row.documentId, row.currentDocumentIdSyncNote),
        'STK lịch sử': row.previousAccount,
        'STK hiện tại': row.currentAccount, 'Tên lịch sử': row.previousName,
        'Tên hiện tại': row.currentName,
        'Ghi chú đồng bộ ID': [
          row.currentDocumentIdSyncNote ? `Hiện tại: ${row.currentDocumentIdSyncNote}` : '',
          ...row.sources.filter(source => source.documentIdSyncNote)
            .map(source => `${formatResolutionPeriods([source.period])}: ${source.documentIdSyncNote}`),
        ].filter(Boolean).join('\n'),
        'Cảnh báo': row.issues.join('; '),
      }));
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(data), 'Check STK-ID');
      XLSX.writeFile(workbook, `Check-STK-ID-${month.replace(/[^\d-]/g, '-')}.xlsx`);
    } catch { setMessage('Không thể xuất báo cáo Check STK & ID.'); }
  }

  const actionRef = useRef({run, userId});
  useEffect(() => { actionRef.current = {run, userId}; });
  useEffect(() => {
    const handleSave = () => {
      if (!actionRef.current.userId) {
        setSettingsOpen(true);
        setLoginOpen(true);
        return;
      }
      void actionRef.current.run('save');
    };
    const handleCheck = () => {
      if (!actionRef.current.userId) {
        setSettingsOpen(true);
        setLoginOpen(true);
        return;
      }
      void actionRef.current.run('check');
    };
    window.addEventListener('trigger-transaction-save', handleSave);
    window.addEventListener('trigger-transaction-check', handleCheck);
    return () => {
      window.removeEventListener('trigger-transaction-save', handleSave);
      window.removeEventListener('trigger-transaction-check', handleCheck);
    };
  }, []);

  const hasPanelContent = hasPendingEdits || message || (showReport && visibleReport) || busy;

  if (!hasPanelContent) {
    return (
      <>
        {/* Supabase Settings Dialog */}
        <Dialog open={settingsOpen} onOpenChange={setSettingsOpen}>
          <DialogContent className="!max-w-md !rounded-2xl !border !border-primary/20 !bg-card p-5 text-foreground shadow-2xl">
            <DialogHeader>
              <DialogTitle className="text-base font-bold normal-case not-italic tracking-tight">Cài đặt Kho Batch Payment</DialogTitle>
              <DialogDescription className="text-xs text-muted-foreground">Tùy chọn đối chiếu và kết nối Supabase.</DialogDescription>
            </DialogHeader>
            <div className="space-y-3 text-xs">
              <label className="block space-y-1.5">
                <span className="font-semibold">Ngân hàng mặc định khi dòng chưa ghi ngân hàng</span>
                <select aria-label="Ngân hàng mặc định khi dòng chưa ghi ngân hàng" disabled={busy} className="w-full rounded-lg border border-primary/20 bg-background px-3 py-2 text-foreground" value={defaultBank} onChange={event => setDefaultBank(event.target.value)}>
                  <option value="VCB">Vietcombank</option><option value="">Chưa xác định</option>
                </select>
              </label>
              <div className="rounded-xl border border-primary/15 bg-primary/5 p-3">
                <p className="font-semibold">Kho dữ liệu Supabase</p>
                <p className="mt-1 text-[10px] text-muted-foreground">Đăng nhập để lưu tháng và kiểm tra dữ liệu Transaction đã lưu.</p>
                <div className="mt-2 flex flex-wrap items-center gap-2">
                  {!userId ? <button type="button" className={buttonClass} disabled={busy} onClick={() => setLoginOpen(value => !value)}>Đăng nhập kho</button>
                    : <button type="button" className={buttonClass} disabled={busy} onClick={async () => {
                      const { error } = await supabase.auth.signOut();
                      setMessage(error ? error.message : 'Đã đăng xuất kho.');
                    }}>Đăng xuất kho</button>}
                  {userId && <span className="text-[10px] text-emerald-700">Đã kết nối</span>}
                  {userId && <button type="button" className={buttonClass} disabled={busy || hasPendingEdits} title="Thay Batch Payment tháng đang chọn bằng bản đã lưu trên Supabase" onClick={() => { setSettingsOpen(false); void run('load'); }}>Tải tháng</button>}
                </div>
                {loginOpen && !userId && <form className="mt-3 space-y-2" onSubmit={event => {event.preventDefault(); void login();}}>
                  <label className="block">Email<input type="email" required autoComplete="username" className="mt-1 block w-full rounded border p-2 text-foreground bg-background" value={email} onChange={event => setEmail(event.target.value)} /></label>
                  <label className="block">Mật khẩu<input type="password" required autoComplete="current-password" className="mt-1 block w-full rounded border p-2 text-foreground bg-background" value={password} onChange={event => setPassword(event.target.value)} /></label>
                  <button className={buttonClass} disabled={busy}>Đăng nhập</button>
                  <p className="text-[10px] text-muted-foreground">Cần tài khoản được quản trị viên cấp quyền kho payroll.</p>
                </form>}
              </div>
            </div>
          </DialogContent>
        </Dialog>
      </>
    );
  }

  return <section aria-label="Kho Batch Payment theo tháng" className="shrink-0 border-b border-primary/15 bg-card p-2 text-foreground" style={{fontFamily: 'var(--font-table, var(--font-main))'}}>
    {busy && <div role="status" className="text-xs font-semibold text-primary py-1">Đang xử lý kết nối Supabase…</div>}
    {hasPendingEdits && <p role="status" className="mt-1 text-xs text-primary">Có chỉnh sửa chưa lưu. Lưu sửa chỉ lưu dữ liệu trên máy; sau đó bấm Lưu tháng để cập nhật Supabase trước khi Check STK & ID.</p>}
    {message && <p role="status" className="text-xs mt-1">{message}</p>}
    {showReport && report && !visibleReport && <p className="text-xs mt-1">Dữ liệu đã đổi. Bấm Check STK & ID để kiểm tra lại.</p>}
    {showReport && visibleReport && <div className="mt-2">
      <div className="flex flex-wrap items-center gap-2 text-xs">
        <strong>Check STK & ID · {month}</strong>
        <span>{exceptions.length} cần kiểm tra · {comparisons.length - exceptions.length} khớp</span>
        <button type="button" className={buttonClass} disabled={!exceptions.length || busy} onClick={() => void exportReport()}>Xuất kết quả</button>
      </div>
      {!localMatchesCloud && <div role="status" className="mt-2 flex flex-wrap items-center gap-2 rounded-xl border border-primary/20 bg-primary/5 p-2 text-xs">
        <span>Batch Payment trên máy khác bản Supabase đang kiểm tra. Lưu sửa chỉ lưu trên máy; bấm Lưu tháng để cập nhật Supabase, hoặc tải bản đã lưu trước khi đồng bộ.</span>
        <button type="button" className={buttonClass} disabled={busy || hasPendingEdits} onClick={() => void run('load')}>Tải bản đã lưu</button>
      </div>}
      <details className="mt-2 rounded-xl border border-primary/15 bg-primary/5 px-3 py-2 text-xs">
        <summary className="cursor-pointer font-semibold">Quy tắc & nguồn · {visibleReport.versions.length} tháng</summary>
        <div className="mt-2 space-y-1.5 text-muted-foreground">
          <p>Chọn ID khi cùng tên + STK; chọn tên khi cùng ID + STK; chọn STK khi cùng ID + tên. Các dòng phải cùng ngân hàng và không mâu thuẫn trong cùng tháng.</p>
          <p>Ô trống có thể được bổ sung từ tháng có giá trị đúng. Trùng tên đơn thuần không dùng để đối chiếu. Chỉ trùng STK cần xác minh chủ tài khoản, không tự gộp nhân viên.</p>
          <p>Tên được so sánh sau khi bỏ khác biệt hoa/thường, dấu và khoảng trắng. ID và STK giữ nguyên số 0 đầu. Dữ liệu lỗi, số mũ hoặc nickname không được dùng làm nguồn STK.</p>
          <p>Hộp thoại cho chọn giá trị nguồn và các tháng cập nhật. Nếu thông tin thay đổi hợp lệ theo tháng, bỏ chọn tháng đó hoặc chọn Giữ nguyên. Chỉ trường đã chọn được đồng bộ.</p>
          <p>Mỗi lần Check tải dữ liệu Supabase mới nhất và đối chiếu các tháng lịch sử đã Lưu sửa trên máy với bản cloud. Lưu sửa không ghi Supabase; Lưu tháng mới cập nhật Supabase.</p>
          <p>Hiện tại: #{visibleReport.currentVersion.id} · {formatHistoryDate(visibleReport.currentVersion.created_at)}. {visibleReport.comparisons.length - comparisons.length} dòng chưa tìm thấy lịch sử và không có cảnh báo được ẩn.</p>
          <p>Nguồn: {visibleReport.versions.map(version => `${formatResolutionPeriods([version.period])} (#${version.id})`).join(' · ') || 'Chưa có tháng trước'}</p>
        </div>
      </details>
      {exceptions.length > 0 ? <>
        <div className="max-h-[55vh] overflow-auto mt-2 rounded-xl border border-primary/15">
          <TransactionHistoryTable rows={exceptions} month={month} page={activePage} renderActions={row => {
            const actions = resolutionGroups.get(row.currentRowIndex) || [];
            return <div className="flex flex-col items-start gap-1.5">
              {actions.map(action => <button key={action.field} type="button" className={resolutionButtonClass}
                disabled={busy || hasPendingEdits || !localMatchesCloud}
                title={action.rule}
                onClick={() => setDecision({context, rowIndex: row.currentRowIndex, field: action.field, optionKey: '', selectedVersionIds: []})}>
                Chọn {action.label}
              </button>)}
              {!actions.length && <span className="text-muted-foreground" title="Cần kiểm tra thông tin nhân viên hoặc bổ sung dữ liệu trước khi đồng bộ.">Cần xác minh</span>}
            </div>;
          }} />
        </div>
        <div className="flex items-center gap-2 mt-1 text-xs"><button type="button" className={buttonClass} disabled={activePage === 1} onClick={() => setPage(activePage - 1)}>Trước</button><span>{activePage}/{pageCount}</span><button type="button" className={buttonClass} disabled={activePage === pageCount} onClick={() => setPage(activePage + 1)}>Sau</button></div>
      </> : <p role="status" className="mt-3 text-xs">Không có chênh lệch ID, tên hoặc STK cần kiểm tra.</p>}
    </div>}
    <Dialog open={settingsOpen} onOpenChange={open => { if (!busy) { setSettingsOpen(open); if (!open) setLoginOpen(false); } }}>
      <DialogContent className="!max-w-md !rounded-2xl !border !border-primary/20 !bg-card p-5 text-foreground shadow-2xl">
        <DialogHeader>
          <DialogTitle className="text-base font-bold normal-case not-italic tracking-tight">Cài đặt Transaction</DialogTitle>
          <DialogDescription className="text-xs text-muted-foreground">Các tuỳ chọn kết nối và dữ liệu mặc định của riêng bảng Transaction.</DialogDescription>
        </DialogHeader>
        <div className="space-y-4 text-xs">
          <label className="block space-y-1.5">
            <span className="font-semibold">Ngân hàng mặc định khi dòng chưa ghi ngân hàng</span>
            <select aria-label="Ngân hàng mặc định khi dòng chưa ghi ngân hàng" disabled={busy} className="w-full rounded-lg border border-primary/20 bg-background px-3 py-2 text-foreground" value={defaultBank} onChange={event => setDefaultBank(event.target.value)}>
              <option value="VCB">Vietcombank</option><option value="">Chưa xác định</option>
            </select>
          </label>
          <div className="rounded-xl border border-primary/15 bg-primary/5 p-3">
            <p className="font-semibold">Kho dữ liệu Supabase</p>
            <p className="mt-1 text-[10px] text-muted-foreground">Đăng nhập để lưu tháng và kiểm tra dữ liệu Transaction đã lưu.</p>
            <div className="mt-2 flex flex-wrap items-center gap-2">
              {!userId ? <button type="button" className={buttonClass} disabled={busy} onClick={() => setLoginOpen(value => !value)}>Đăng nhập kho</button>
                : <button type="button" className={buttonClass} disabled={busy} onClick={async () => {
                  const { error } = await supabase.auth.signOut();
                  setMessage(error ? error.message : 'Đã đăng xuất kho.');
                }}>Đăng xuất kho</button>}
              {userId && <span className="text-[10px] text-emerald-700">Đã kết nối</span>}
              {userId && <button type="button" className={buttonClass} disabled={busy || hasPendingEdits} title="Thay Batch Payment tháng đang chọn bằng bản đã lưu trên Supabase" onClick={() => { setSettingsOpen(false); void run('load'); }}>Tải tháng</button>}
            </div>
            {loginOpen && !userId && <form className="mt-3 space-y-2" onSubmit={event => {event.preventDefault(); void login();}}>
              <label className="block">Email<input type="email" required autoComplete="username" className="mt-1 block w-full rounded border p-2 text-foreground bg-background" value={email} onChange={event => setEmail(event.target.value)} /></label>
              <label className="block">Mật khẩu<input type="password" required autoComplete="current-password" className="mt-1 block w-full rounded border p-2 text-foreground bg-background" value={password} onChange={event => setPassword(event.target.value)} /></label>
              <button className={buttonClass} disabled={busy}>Đăng nhập</button>
              <p className="text-[10px] text-muted-foreground">Cần tài khoản được quản trị viên cấp quyền kho payroll.</p>
            </form>}
          </div>
        </div>
      </DialogContent>
    </Dialog>
    <Dialog open={Boolean(visibleGroup)} onOpenChange={open => { if (!open && !busy) setDecision(null); }}>
      <DialogContent className="!max-w-xl max-h-[90vh] overflow-y-auto !rounded-2xl !border !border-primary/20 !bg-card p-5 text-foreground shadow-2xl">
        <DialogHeader>
          <DialogTitle className="text-base font-bold normal-case not-italic tracking-tight">Đồng bộ {chosenField?.label} theo tháng nào?</DialogTitle>
          <DialogDescription className="text-xs text-muted-foreground">{chosenField?.rule} Chọn giá trị đúng và các tháng cần sửa.</DialogDescription>
        </DialogHeader>
        {visibleGroup && decision && <div className="space-y-3 text-xs">
          <fieldset disabled={busy} className="max-h-52 space-y-2 overflow-auto">
            <legend className="mb-2 font-semibold">Lấy giá trị từ tháng</legend>
            {visibleGroup.options.map(option => <label key={option.key} className={`flex cursor-pointer items-start gap-2 rounded-xl border p-2.5 ${decision.optionKey === option.key ? 'border-primary bg-primary/10' : 'border-primary/20'}`}>
              <input type="radio" name="identity-source-month" value={option.key} checked={decision.optionKey === option.key}
                onChange={() => setDecision({...decision, optionKey: option.key, selectedVersionIds: identityResolutionTargets(visibleGroup, option.key).map(target => target.versionId)})} />
              <span><span className="block font-semibold">{formatResolutionPeriods([option.period])}{option.versionId === visibleReport?.currentVersion.id ? ' · hiện tại' : ''}</span><span className="tabular-nums">{option.value}</span></span>
            </label>)}
          </fieldset>
          {chosenOption && <fieldset disabled={busy} className="rounded-xl border border-primary/20 p-3">
            <legend className="px-1 font-semibold">Tháng sẽ cập nhật · {selectedTargets.length}</legend>
            <p className="mb-2 text-muted-foreground">Bỏ chọn những tháng có thay đổi hợp lệ cần giữ lại.</p>
            <div className="max-h-44 space-y-2 overflow-auto">{targets.map(target => <label key={target.versionId} className="flex cursor-pointer items-start gap-2">
              <input type="checkbox" checked={decision.selectedVersionIds.includes(target.versionId)} onChange={event => setDecision({...decision,
                selectedVersionIds: event.target.checked ? [...decision.selectedVersionIds, target.versionId]
                  : decision.selectedVersionIds.filter(id => id !== target.versionId),
              })} />
              <span><span className="font-semibold">{formatResolutionPeriods([target.period])} · {target.rowIndexes.length} dòng</span><span className="block tabular-nums">{target.fromValues.map(value => value || '(trống)').join(' / ')} → {chosenOption.value}</span></span>
            </label>)}</div>
          </fieldset>}
          <p className="text-muted-foreground">Các tháng đã chọn được lưu cùng lúc lên Supabase. Danh mục nhân viên được cập nhật theo Transaction của tháng đang xem.</p>
          <div className="flex justify-end gap-2">
            <button type="button" className={buttonClass} disabled={busy} onClick={() => setDecision(null)}>Giữ nguyên</button>
            <button type="button" className={buttonClass} disabled={busy || !selectedTargets.length || !localMatchesCloud || hasPendingEdits} onClick={() => void resolveIdentity()}>{busy ? 'Đang lưu…' : 'Lưu đồng bộ'}</button>
          </div>
        </div>}
      </DialogContent>
    </Dialog>
  </section>;
}
