import {
  createContext,
  useContext,
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  useMemo,
  type ReactNode,
} from "react";
import { Cloud, CloudOff, LoaderCircle } from "lucide-react";
import { supabase, isSupabaseConfigured } from "../../../lib/supabaseClient";
import { useAppActions, useAppDataOnly } from "./AppDataContext";
import { CloudRepository } from "../cloud/repository";
import { SyncEngine, OWNER_KEY, type SyncState } from "../cloud/engine";
import { saveSnapshot } from "../utils/snapshot-manager";
import { withDeadline } from "../cloud/deadline";
import {
  Dialog,
  DialogContent,
  DialogTitle,
  DialogDescription,
} from "../../components/ui/dialog";
interface CloudContextValue {
  flush: () => Promise<void>;
  open: () => void;
  state: SyncState;
  authenticated: boolean;
}
const Context = createContext<CloudContextValue | null>(null);
type WorkspaceUser = { id: string; email?: string };
// This hook deliberately shares the same context as the provider.
// eslint-disable-next-line react-refresh/only-export-components
export function useCloudWorkspace() {
  return useContext(Context);
}
export function CloudSyncButton() {
  const cloud = useContext(Context);
  if (!cloud) return null;
  const configured = isSupabaseConfigured();
  const label = !configured
    ? "Chưa kết nối Cloud"
    : !cloud.authenticated
      ? "Đăng nhập để đồng bộ"
    : cloud.state.status === "synced"
      ? "Đã đồng bộ"
      : cloud.state.status === "saving"
        ? "Đang đồng bộ"
        : cloud.state.status === "starting"
          ? "Đang tải dữ liệu"
          : cloud.state.status === "conflict"
            ? "Xung đột dữ liệu"
            : cloud.state.status === "offline"
              ? "Chờ kết nối"
              : "Lỗi đồng bộ";
  return (
    <button
      type="button"
      onClick={cloud.open}
      title={label}
      aria-label={label}
      className="inline-flex shrink-0 items-center gap-1 rounded-full border border-primary/20 px-2 py-1 text-[10px] text-primary whitespace-nowrap active:scale-[0.98]"
    >
      {cloud.authenticated && (cloud.state.status === "saving" || cloud.state.status === "starting") ? (
        <LoaderCircle className="h-3 w-3 animate-spin" />
      ) : cloud.state.status === "offline" ? (
        <CloudOff className="h-3 w-3" />
      ) : (
        <Cloud className="h-3 w-3" />
      )}
      <span className="max-md:hidden">{label}</span>
    </button>
  );
}
export function CloudWorkspaceProvider({ children }: { children: ReactNode }) {
  const { rawAppData, isStorageHydrating } = useAppDataOnly();
  const { applyCloudData, setCloudRestorePending } = useAppActions();
  const dataRef = useRef(rawAppData);
  useLayoutEffect(() => {
    dataRef.current = rawAppData;
  }, [rawAppData]);
  const engine = useRef<SyncEngine | null>(null);
  const [user, setUser] = useState<WorkspaceUser | null>(null);
  const [authReady, setAuthReady] = useState(() => !isSupabaseConfigured());
  const [state, setState] = useState<SyncState>(() =>
    isSupabaseConfigured()
      ? { status: "starting" }
      : { status: "error", message: "Chưa có cấu hình cloud." },
  );
  const [dialogOpen, setDialogOpen] = useState(false);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [busy, setBusy] = useState(false);
  // Identity is retained for token refresh, but a logout/login gets a fresh
  // object even for the same account and must restore cloud again.
  const [unlockedUser, setUnlockedUser] = useState<WorkspaceUser | null>(null);
  useEffect(() => {
    if (!isSupabaseConfigured()) return;
    let alive = true;
    let receivedAuthEvent = false;
    withDeadline(() => supabase.auth.getSession(), "Kiểm tra đăng nhập quá lâu. Tải lại trang để thử lại.").then(({ data, error }) => {
      if (alive && !receivedAuthEvent) {
        setUser(previous => previous?.id === data.session?.user.id ? previous : data.session?.user || null);
        setAuthReady(true);
        if (error) setLoginError(error.message);
      }
    }).catch(error => {
      if (!alive || receivedAuthEvent) return;
      setAuthReady(true);
      setState({ status: "error", message: error instanceof Error ? error.message : "Không thể kiểm tra đăng nhập." });
      setLoginError(error instanceof Error ? error.message : "Không thể kiểm tra đăng nhập.");
    });
    const { data } = supabase.auth.onAuthStateChange((_event, session) => {
      if (!alive) return;
      receivedAuthEvent = true;
      // The callback only updates React state; network calls run in the effect below.
      setUser((previous) =>
        previous?.id === session?.user.id ? previous : session?.user || null,
      );
      setAuthReady(true);
    });
    return () => {
      alive = false;
      data.subscription.unsubscribe();
    };
  }, []);
  useEffect(() => {
    if (!authReady || isStorageHydrating || !user) return;
    let alive = true;
    const instance = new SyncEngine(
      new CloudRepository(user.id),
      () => dataRef.current,
      applyCloudData,
      (next) => {
        if (alive) {
          setState(previous => previous.status === next.status && previous.message === next.message && previous.updatedAt === next.updatedAt && previous.revision === next.revision && previous.conflicts === next.conflicts ? previous : next);
          if (next.status === "synced") setUnlockedUser(user);
        }
      },
      async () => {
        await saveSnapshot(dataRef.current, {
          title: "Bản máy chưa đồng bộ trước khi tải cloud",
          trigger: "restore",
          isPinned: true,
        });
      },
    );
    engine.current = instance;
    void instance.start();
    const retry = () => {
      if (document.visibilityState !== "hidden") void instance.retry();
    };
    const interval = setInterval(retry, 30000);
    window.addEventListener("online", retry);
    window.addEventListener("focus", retry);
    return () => {
      alive = false;
      instance.stop();
      engine.current = null;
      clearInterval(interval);
      window.removeEventListener("online", retry);
      window.removeEventListener("focus", retry);
    };
  }, [authReady, isStorageHydrating, user, applyCloudData]);
  useEffect(() => {
    engine.current?.changed();
  }, [rawAppData]);
  const flush = useCallback(async () => {
    // React must publish Save's state change to dataRef before capturing the upload.
    await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
    const current = engine.current;
    if (!current)
      throw new Error("Đăng nhập để đồng bộ đầy đủ các bảng lên cloud.");
    await current.flush();
  }, []);
  const configured = isSupabaseConfigured();
  const owned = localStorage.getItem(OWNER_KEY);
  const gated =
    configured &&
    (!authReady || isStorageHydrating || (!!user && unlockedUser !== user) || (!!owned && !user));
  useLayoutEffect(() => {
    setCloudRestorePending(gated);
    return () => setCloudRestorePending(false);
  }, [gated, setCloudRestorePending]);
  const showDialog = dialogOpen || (authReady && gated);
  const [mode, setMode] = useState<"login" | "signup">("login");

  async function handleSubmit(event: React.FormEvent) {
    event.preventDefault();
    if (!isSupabaseConfigured()) {
      setLoginError("Chưa cấu hình Supabase Cloud (cần VITE_SUPABASE_URL & VITE_SUPABASE_ANON_KEY trong Settings/Biến môi trường).");
      return;
    }
    setBusy(true);
    setLoginError("");
    try {
      if (mode === "signup") {
        const { data: signUpData, error } = await supabase.auth.signUp({
          email: email.trim(),
          password,
        });
        if (error) throw error;
        if (signUpData.user) {
          if (!signUpData.session) {
            setLoginError("Đã gửi email xác nhận. Vui lòng kiểm tra hộp thư để xác nhận tài khoản nếu cần thiết.");
          } else {
            const signedIn = signUpData.user;
            setUser(previous => previous?.id === signedIn.id ? previous : signedIn);
            setDialogOpen(false);
          }
        }
      } else {
        const { data: loginData, error } = await supabase.auth.signInWithPassword(
          { email: email.trim(), password },
        );
        if (error) throw error;
        setUser(previous => previous?.id === loginData.user.id ? previous : loginData.user);
        setDialogOpen(false);
      }
    } catch (error) {
      setLoginError(
        error instanceof Error ? error.message : (mode === "signup" ? "Không thể đăng ký" : "Không thể đăng nhập"),
      );
    } finally {
      setPassword("");
      setBusy(false);
    }
  }
  const messages: Record<SyncState["status"], string> = {
    starting: "Đang mở dữ liệu cloud…",
    saving: "Đang lưu thay đổi lên cloud…",
    synced: "Đã lưu và đồng bộ dữ liệu giữa các máy.",
    offline: "Mất kết nối — thay đổi đang được giữ trên máy.",
    conflict: "Cần chọn bản dữ liệu",
    error: "Đồng bộ chưa hoàn tất",
  };
  const open = useCallback(() => setDialogOpen(true), []);
  const contextValue = useMemo(() => ({ flush, open, state, authenticated: !!user }), [flush, open, state, user]);
  return (
    <Context.Provider value={contextValue}>
      {gated ? (
        <div className="flex min-h-screen items-center justify-center text-sm text-muted-foreground">
          {!authReady ? "Đang kiểm tra đăng nhập…" : "Đang chờ bản cloud mới nhất…"}
        </div>
      ) : (
        children
      )}
      <Dialog open={showDialog} onOpenChange={setDialogOpen}>
        <DialogContent
          onKeyDown={(event) => event.stopPropagation()}
          onInteractOutside={(event) => {
            if (gated) event.preventDefault();
          }}
          onEscapeKeyDown={(event) => {
            if (gated) event.preventDefault();
          }}
          className="max-w-md"
        >
          <DialogTitle>Đồng bộ giữa các máy</DialogTitle>
          <DialogDescription>
            Dùng cùng tài khoản trên các máy. Mỗi lần đăng nhập, ứng dụng tải bản
            cloud mới nhất trước khi mở các bảng. Thay đổi trên máy chưa đồng bộ
            được giữ trong bản sao lưu riêng.
          </DialogDescription>
          {!user ? (
            <form onSubmit={handleSubmit} className="space-y-3">
              {!configured && (
                <div className="rounded-xl border border-amber-500/30 bg-amber-50/80 p-3 text-xs text-amber-900 space-y-1">
                  <p className="font-bold">⚠️ Chưa cấu hình Supabase Cloud</p>
                  <p className="leading-relaxed">
                    Để kích hoạt đồng bộ qua Cloud giữa các thiết bị, vui lòng cung cấp biến môi trường <strong>VITE_SUPABASE_URL</strong> và <strong>VITE_SUPABASE_ANON_KEY</strong> trong phần Settings của AI Studio hoặc file <code>.env</code>.
                  </p>
                  <p className="leading-relaxed text-[11px] text-amber-800/80">
                    Dữ liệu trên máy của bạn hiện tại vẫn đang được lưu trữ cục bộ đầy đủ và an toàn.
                  </p>
                </div>
              )}
              <label className="block text-sm">
                Email
                <input
                  type="email"
                  autoComplete="username"
                  required
                  value={email}
                  onChange={(event) => setEmail(event.target.value)}
                  className="mt-1 w-full rounded-lg border bg-background p-2"
                />
              </label>
              <label className="block text-sm">
                Mật khẩu
                <input
                  type="password"
                  autoComplete="current-password"
                  required
                  value={password}
                  onChange={(event) => setPassword(event.target.value)}
                  className="mt-1 w-full rounded-lg border bg-background p-2"
                />
              </label>
              {loginError && (
                <p role="alert" className="text-sm text-red-700">
                  {loginError}
                </p>
              )}
              <div className="flex items-center justify-between pt-1">
                <button
                  type="submit"
                  disabled={busy || !configured}
                  className="rounded-full bg-primary px-4 py-2 text-sm text-primary-foreground disabled:opacity-50 cursor-pointer"
                >
                  {busy ? "Đang xử lý…" : mode === "signup" ? "Đăng ký tài khoản" : "Đăng nhập"}
                </button>
                {configured && (
                  <button
                    type="button"
                    onClick={() => {
                      setMode((m) => (m === "login" ? "signup" : "login"));
                      setLoginError("");
                    }}
                    className="text-xs text-primary hover:underline cursor-pointer"
                  >
                    {mode === "login" ? "Chưa có tài khoản? Đăng ký" : "Đã có tài khoản? Đăng nhập"}
                  </button>
                )}
              </div>
            </form>
          ) : (
            <div className="space-y-3 text-sm">
              <p className="break-all text-muted-foreground">{user.email}</p>
              <p role="status">{messages[state.status]}</p>
              {state.message && (
                <p className="text-muted-foreground">{state.message}</p>
              )}
              {state.updatedAt && (
                <p className="text-xs text-muted-foreground">
                  Lần lưu cloud:{" "}
                  {new Date(state.updatedAt).toLocaleString("vi-VN")}
                </p>
              )}
              {state.revision !== undefined && (
                <p className="text-xs text-muted-foreground">
                  Phiên bản dữ liệu cloud: #{state.revision}
                </p>
              )}
              {gated && (state.status === "error" || state.status === "offline") && (
                <p className="text-xs text-muted-foreground">
                  Kết nối mạng rồi bấm Đồng bộ ngay để mở bản cloud mới nhất.
                </p>
              )}
              {state.status === "conflict" ? (
                <>
                  <p className="break-words text-xs">
                    {state.conflicts?.join(", ")}
                  </p>
                  <div className="flex gap-2">
                    <button
                      className="rounded-full bg-primary px-3 py-2 text-primary-foreground"
                      onClick={() => void engine.current?.resolve(false)}
                    >
                      Dùng bản cloud
                    </button>
                    <button
                      className="rounded-full border px-3 py-2"
                      onClick={() => void engine.current?.resolve(true)}
                    >
                      Giữ bản máy
                    </button>
                  </div>
                  <p className="text-xs text-muted-foreground">
                    Ứng dụng tạo bản sao trên máy trước khi xử lý xung đột.
                  </p>
                </>
              ) : (
                <button
                  className="rounded-full border px-3 py-2"
                  disabled={state.status === "starting" || state.status === "saving"}
                  onClick={() => void engine.current?.retry()}
                >
                  Đồng bộ ngay
                </button>
              )}
              <button
                className="block text-xs text-muted-foreground underline"
                onClick={() => void supabase.auth.signOut()}
              >
                Đăng xuất
              </button>
            </div>
          )}
          <p className="text-xs text-muted-foreground">
            Lưu cloud riêng tư trên gói miễn phí, tối đa 200 MB dữ liệu làm việc
            sau nén. Đồng bộ dữ liệu đang dùng và tháng đã lưu; các bản sao lịch
            sử đầy đủ vẫn nằm trên máy.
          </p>
        </DialogContent>
      </Dialog>
    </Context.Provider>
  );
}
