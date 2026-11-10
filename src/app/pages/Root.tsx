import { useState, useEffect } from "react";
import { Outlet, useLocation } from "react-router";
import { Navbar } from "../components/layouts/Navbar";
import { UiSettingsModal } from "../components/UiSettingsModal";
import { ErrorBoundary } from "../components/shared/ErrorBoundary";
import { ConfirmDialog } from "../components/shared/ConfirmDialog";
import { useAppActions, useAppDataOnly } from "../lib/contexts/AppDataContext";
import { restoreWorkspaceMonth } from "../lib/utils/workspace-month-reset";
import { requestPivotCacheMonthReset } from "../lib/utils/pivot-utils";
import { toast } from "sonner";

// ── Root không dùng framer-motion để tránh layout thrashing trên shell layout ──
export function Root() {
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [resetMonth, setResetMonth] = useState<string | null>(null);
  const [resetRevision, setResetRevision] = useState(0);
  const location = useLocation();
  const { rawAppData, isStorageHydrating } = useAppDataOnly();
  const { updateAppData } = useAppActions();
  const month = rawAppData.globalMonth || "03.2026";

  function resetTables() {
    if (!resetMonth || resetMonth !== month || isStorageHydrating) return;
    const result = restoreWorkspaceMonth(rawAppData, resetMonth);
    updateAppData(previous => previous === rawAppData ? result.data : restoreWorkspaceMonth(previous, resetMonth).data, true, true);
    requestPivotCacheMonthReset(resetMonth);
    sessionStorage.removeItem("transaction_reference_return");
    setResetRevision(previous => previous + 1);
    if (result.batchRestored) toast.success(`Đã reset các bảng tháng ${resetMonth}.`);
    else toast.info(`Đã làm mới các bảng tháng ${resetMonth}. Batch Payment giữ dữ liệu hiện có vì chưa có bản khôi phục đúng tháng.`);
  }

  useEffect(() => {
    // AppDataProvider lives above RouterProvider, so publish route changes for
    // route-scoped data work without forcing a document reload.
    window.dispatchEvent(new Event("app-route-changed"));
  }, [location.pathname]);

  useEffect(() => {
    const handleOpenSettings = () => setIsSettingsOpen(true);
    window.addEventListener("open-ui-settings", handleOpenSettings);
    return () => window.removeEventListener("open-ui-settings", handleOpenSettings);
  }, []);

  return (
    <div className="app-viewport-shell flex overflow-hidden font-sans text-foreground bg-background justify-center">
      {/* Main Content Area */}
      <div 
        className="flex-1 flex flex-col overflow-hidden min-w-0 relative bg-background w-full"
      >
        <div className="bg-transparent relative z-40">
          <Navbar
            onOpenSettings={() => setIsSettingsOpen(true)}
            onResetTables={() => setResetMonth(month)}
            resetDisabled={isStorageHydrating}
          />
        </div>

        <main className="flex-1 flex flex-col min-h-0 relative">
          <ErrorBoundary key={`${location.pathname}:${resetRevision}`}>
            <div className="flex-1 flex flex-col min-h-0">
              <Outlet />
            </div>
          </ErrorBoundary>

          {/* Quick Shortcuts Exclamation Floating Dock - Removed as requested */}
        </main>
        
        <UiSettingsModal
          isOpen={isSettingsOpen}
          onClose={() => setIsSettingsOpen(false)}
        />
        <ConfirmDialog
          isOpen={resetMonth !== null}
          onClose={() => setResetMonth(null)}
          onConfirm={resetTables}
          title={`Reset các bảng tháng ${resetMonth ?? month}`}
          description="Khôi phục dữ liệu theo bản lưu hoặc bản gốc đúng tháng, tính lại các bảng liên quan và bỏ bộ lọc đang mở. Bao gồm Batch Payment, Reconcile và Analysis. Deduction đã khóa giữ số đã lưu."
          confirmText="Reset tất cả"
        />
      </div>
    </div>
  );
}
