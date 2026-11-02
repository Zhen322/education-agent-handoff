"use client";
import { useCallback, useEffect, useRef, useState } from "react";
import { ArrowRight, BookOpen, Camera, Check, ChevronRight, CircleHelp, ClipboardList, GraduationCap, Home, Leaf, MessageCircle, ShieldCheck, Sparkles, Sprout, Star, Sun, Target, UsersRound, X } from "lucide-react";
import { act, ApiError, currentWorkspaceId, fetchWorkspace, workspaceUrl } from "@/lib/client";
import type { ApiResult, WorkspaceState } from "@/lib/types";
import type { AppProps, ChatEntryContext, View } from "@/lib/ui-types";
import { readChatEntryContext, writeChatEntryContext } from "@/lib/chat-navigation";
import StudentLearning from "@/components/student-learning";
import { AdminWorkspace } from "@/components/admin-workspace";
import CameraWorkspace from "@/components/camera-workspace";
import CompanionChat from "@/components/companion-chat";
import AccessWorkspace, { roleNames } from "@/components/access-workspace";
import EvidenceReviewWorkspace from "@/components/evidence-review-workspace";
import { effectiveAttempt, effectiveMistakes } from "@/lib/effective-evidence";
import SharedLearning from "@/components/shared-learning";
import HistoryDeletionWorkspace from "@/components/history-deletion-workspace";
import { pruneLearningCaches } from "@/lib/learning-cache-cleanup";

const navigation: { view: View; label: string; icon: typeof Home }[] = [
  { view: "home", label: "今天", icon: Home }, { view: "knowledge", label: "知识地图", icon: BookOpen },
  { view: "camera", label: "拍照学习", icon: Camera }, { view: "papers", label: "我的练习", icon: ClipboardList },
  { view: "mistakes", label: "错题与订正", icon: Target }, { view: "evidence-review", label: "记录复核", icon: ShieldCheck }, { view: "chat", label: "和小禾聊聊", icon: MessageCircle },
];
const validViews: View[] = ["home", "knowledge", "camera", "papers", "practice", "mistakes", "profile", "chat", "report", "access", "shared-help", "evidence-review", "history-delete"];
function belongsToRequest(incoming: WorkspaceState, requested: string | undefined): boolean {
  return currentWorkspaceId() === requested && (!requested || incoming.access?.workspaceId === requested);
}
export default function HomePage() {
  const [state, setState] = useState<WorkspaceState | null>(null);
  const [view, setView] = useState<View>("home");
  const [selectedId, setSelectedId] = useState<string>();
  const [evidenceRevision, setEvidenceRevision] = useState<number>();
  const [chatContext, setChatContext] = useState<ChatEntryContext>();
  const [chatEntryError, setChatEntryError] = useState("");
  const [admin, setAdmin] = useState(false);
  const [loading, setLoading] = useState(true);
  const [signedOut, setSignedOut] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [pending, setPending] = useState(0);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const acceptedState = useRef<WorkspaceState | null>(null);
  const mergeState = useCallback((incoming: WorkspaceState) => {
    const previous = acceptedState.current;
    if (previous?.access && incoming.access && previous.access.workspaceId !== incoming.access.workspaceId) return;
    if (previous && incoming.revision < previous.revision) return;
    acceptedState.current = incoming; pruneLearningCaches(incoming);
    setState(current => {
      if (current?.access && incoming.access && current.access.workspaceId !== incoming.access.workspaceId) return current;
      return !current || incoming.revision >= current.revision ? incoming : current;
    });
  }, []);
  useEffect(() => { if (state) pruneLearningCaches(state); }, [state]);
  const load = useCallback(async () => {
    const requested = currentWorkspaceId();
    try { const data = await fetchWorkspace(); if (!belongsToRequest(data.state, requested)) throw new Error("工作区已改变，请重新打开当前空间"); mergeState(data.state); setSignedOut(false); }
    catch (e) { if (currentWorkspaceId() !== requested) return; const text = e instanceof Error ? e.message : "暂时无法连接，请重试"; setSignedOut(e instanceof ApiError && e.status === 401); setError(text); }
    finally { if (currentWorkspaceId() === requested) setLoading(false); }
  }, [mergeState]);
  const refresh = useCallback(async () => {
    const requested = currentWorkspaceId();
    try { const data = await fetchWorkspace(); if (belongsToRequest(data.state, requested)) mergeState(data.state); }
    catch (e) { if (currentWorkspaceId() === requested && e instanceof ApiError && [401, 403, 404].includes(e.status)) { setState(null); setSignedOut(e.status === 401); setError(e.message); } }
  }, [mergeState]);
  useEffect(() => {
    let active = true;
    const requested = currentWorkspaceId();
    void fetchWorkspace().then(data => {
      if (!active || !belongsToRequest(data.state, requested)) return;
      mergeState(data.state); setSignedOut(false);
    }).catch(e => {
      if (!active || currentWorkspaceId() !== requested) return;
      setSignedOut(e instanceof ApiError && e.status === 401); setError(e instanceof Error ? e.message : "暂时无法连接，请重试");
    }).finally(() => { if (active && currentWorkspaceId() === requested) setLoading(false); });
    return () => { active = false; };
  }, [mergeState]);
  useEffect(() => {
    const refreshVisible = () => { if (document.visibilityState === "visible") void refresh(); };
    window.addEventListener("focus", refreshVisible); document.addEventListener("visibilitychange", refreshVisible);
    return () => { window.removeEventListener("focus", refreshVisible); document.removeEventListener("visibilitychange", refreshVisible); };
  }, [refresh]);
  useEffect(() => {
    const restore = () => { const params = new URLSearchParams(location.hash.slice(1)); const requested = params.get("view") as View; if (validViews.includes(requested)) { setView(requested); setSelectedId(params.get("id") || undefined); const revision = params.get("evidenceRevision"); setEvidenceRevision(revision !== null && /^\d+$/.test(revision) && Number.isSafeInteger(Number(revision)) ? Number(revision) : undefined); setChatContext(undefined); setChatEntryError(""); if (requested === "chat") { try { setChatContext(readChatEntryContext(params)); } catch (error) { setChatEntryError(error instanceof Error ? error.message : "学伴入口无效，请重新选择题目。"); } } } setAdmin(params.get("mode") === "admin"); };
    restore(); window.addEventListener("hashchange", restore); return () => window.removeEventListener("hashchange", restore);
  }, []);
  useEffect(() => () => { if (timer.current) clearTimeout(timer.current); }, []);
  const navigate = useCallback((next: View, id?: string, context?: ChatEntryContext, revision?: number) => {
    setView(next); setSelectedId(id); setEvidenceRevision(revision); setChatContext(next === "chat" ? context : undefined); setChatEntryError(""); setAdmin(false);
    const params = new URLSearchParams({ view: next }); if (id) params.set("id", id);
    if (next === "chat" && context) writeChatEntryContext(params, context);
    if (revision !== undefined && Number.isSafeInteger(revision) && revision >= 0) params.set("evidenceRevision", String(revision));
    location.hash = params.toString(); window.scrollTo({ top: 0, behavior: "smooth" });
  }, []);
  const openAdmin = () => { setAdmin(true); setView("home"); location.hash = "mode=admin"; };
  const mutate = useCallback(async (action: string, payload: Record<string, unknown> = {}): Promise<ApiResult | undefined> => {
    const requested = currentWorkspaceId();
    setPending(n => n + 1); setError("");
    try { const data = await act(action, payload); if (!belongsToRequest(data.state, requested)) return undefined; mergeState(data.state);
      if (!["attempt.save", "attempt.help", "chat.send", "chat.open", "chat.draft.save", "chat.delete.preview", "evidence.preview", "history.delete.preview"].includes(action)) { setNotice(typeof data.result?.message === "string" ? data.result.message : "已保存"); if (timer.current) clearTimeout(timer.current); timer.current = setTimeout(() => setNotice(""), 3500); } return data;
    } catch (e) { if (currentWorkspaceId() === requested) { setError(e instanceof Error ? e.message : "操作失败，请重试"); if (e instanceof ApiError && [401, 403, 404, 409].includes(e.status)) void refresh(); } if (action.startsWith("chat.") || action.startsWith("memory.preference.") || action === "help.create" || action.startsWith("evidence.") || action.startsWith("history.")) throw e; return undefined; }
    finally { setPending(n => n - 1); }
  }, [mergeState, refresh]);
  if (loading) return <main className="welcome"><div className="brand-mark"><Sprout size={30}/></div><h1>小禾学伴</h1><p className="muted">正在打开你的学习空间…</p><span className="spinner"/></main>;
  if (!state) return <main className="welcome"><div className="welcome-panel"><div className="brand"><span className="brand-mark"><Sprout/></span><strong>小禾学伴</strong></div><img className="welcome-art" src="/illustrations/xiaohe-companion.png" alt="小禾与学伴一起学习"/><span className="eyebrow">每一次小进步，都值得被看见</span><h1>把今天要学的，<br/>变成刚好适合的练习。</h1><p className="muted">从一道题开始，找到方法，慢慢成长。</p>{signedOut ? <a className="btn btn-primary" href={`/signin-with-chatgpt?return_to=${encodeURIComponent(workspaceUrl("/"))}`} target="_top">登录，打开学习空间<ArrowRight size={18}/></a> : <><p role="alert" className="error-text">{error}</p><button className="btn btn-primary" onClick={() => { setLoading(true); setError(""); void load(); }}>重新连接</button>{currentWorkspaceId() && <a className="btn btn-secondary" href={workspaceUrl("/#view=access", null)}>返回我的空间，查看邀请</a>}</>}<p className="welcome-note">四年级数学试用版 · 初始内容为原创演示课程</p></div></main>;
  const props: AppProps = { state, mutate, navigate, selectedId, evidenceRevision, chatContext, chatEntryError, reload: refresh, busy: pending > 0 };
  const role = state.access?.role ?? state.actor.role;
  const learningRole = role === "owner" || role === "student";
  const contentOnly = role === "editor" || role === "reviewer";
  const observer = role === "parent" || role === "teacher";
  const canAdmin = role === "owner" || contentOnly;
  const showAdmin = view !== "access" && canAdmin && (contentOnly || admin);
  const roleNavigation = learningRole ? navigation : observer ? [{ view: "home" as View, label: "学习报告", icon: ClipboardList }, ...(role === "teacher" ? [{ view: "shared-help" as View, label: "已共享求助", icon: MessageCircle }, { view: "evidence-review" as View, label: "证据复核", icon: ShieldCheck }] : [])] : [];
  const currentLabel = view === "access" ? "成员与共享" : view === "history-delete" ? "删除学习历史" : view === "evidence-review" ? "记录更正与复核" : showAdmin ? "教研工作台" : observer ? view === "shared-help" && role === "teacher" ? "已共享求助" : "学习报告" : view === "profile" ? "我的学习空间" : view === "practice" ? "专注练习" : view === "report" ? "学习报告" : navigation.find(n => n.view === view)?.label;
  const actorLabel = learningRole ? state.profile.name : state.actor.name;
  const desktopHome = () => contentOnly ? openAdmin() : navigate("home");
  const mobileItems = learningRole ? [{ view: "home" as View, label: "首页", icon: Home }, { view: "knowledge" as View, label: "知识", icon: BookOpen }, { view: "camera" as View, label: "拍照", icon: Camera }, { view: "papers" as View, label: "出题", icon: ClipboardList }, { view: "profile" as View, label: "我的", icon: Sprout }] : [...roleNavigation, { view: "access" as View, label: "成员与空间", icon: UsersRound }];
  return <div className={`app-shell ${showAdmin ? "admin-mode" : ""}`}>
    <aside className="sidebar"><button className="brand" onClick={desktopHome} aria-label="小禾学伴首页"><span className="brand-mark"><Sprout size={26}/></span><span><strong>小禾学伴</strong><small>陪你一起，慢慢长大</small></span></button><div className="sidebar-label">{learningRole ? "我的学习空间" : "当前授权空间"}</div><nav aria-label="主导航">{roleNavigation.map(({view: tab,label,icon: Icon}) => <button key={tab} className={`nav-item ${!showAdmin && (view === tab || tab === "papers" && ["practice","report"].includes(view) || observer && tab === "home" && view === "report") ? "active" : ""}`} onClick={() => navigate(tab)}><Icon size={21}/><span>{label}</span>{tab === "chat" && <span className="tiny-dot"/>}</button>)}<button className={`nav-item ${view === "access" ? "active" : ""}`} onClick={() => navigate("access")}><UsersRound size={21}/><span>成员与共享</span></button></nav><div className="sidebar-bottom"><div className="sidebar-tip"><Leaf size={22}/><p>{learningRole ? <>不用每次都很快，<br/>想明白，也是进步。</> : <>看见具体的过程，<br/>给出刚刚好的支持。</>}</p></div>{canAdmin && <button className={`nav-item ${showAdmin ? "active" : ""}`} onClick={openAdmin}><GraduationCap size={20}/><span>教研工作台</span><ChevronRight size={15}/></button>}<button className="profile-link" onClick={() => navigate(learningRole ? "profile" : "access")}><span className="avatar">{actorLabel?.slice(0,1) || "同"}</span><span><strong>{actorLabel}</strong><small>{learningRole ? `${state.profile.grade} · 学习档案` : roleNames[role]}</small></span><ChevronRight size={17}/></button></div></aside>
    <div className="workspace"><header className="topbar"><div className="mobile-brand"><Sprout size={25}/><strong>小禾学伴</strong></div><div className="breadcrumb"><span>{learningRole ? "学习空间" : "授权空间"}</span><ChevronRight size={14}/><strong>{currentLabel}</strong></div><div className="topbar-actions">{showAdmin && role === "owner" && <button className="text-link" onClick={()=>navigate("home")}><Home size={15}/>返回学习空间</button>}{contentOnly && view === "access" && <button className="text-link" onClick={openAdmin}><GraduationCap size={15}/>教研工作台</button>}<button className="shell-access-button" onClick={() => navigate("access")} aria-label="成员、授权与切换工作区"><UsersRound size={17}/><span>成员与共享</span></button><span className="course-badge"><span/>{learningRole ? `${state.profile.grade} · 数学` : roleNames[role]}</span><button title={learningRole ? "我的学习档案" : "我的访问权限"} aria-label={learningRole ? "我的学习档案" : "我的访问权限"} className="avatar small" onClick={() => navigate(learningRole ? "profile" : "access")}>{actorLabel?.slice(0,1) || "同"}</button></div></header><main id="main-content" className={`main-content ${showAdmin ? "admin-content-wrap" : ""}`}>{error && <div className="error-banner" role="alert"><CircleHelp size={19}/><span>{error}</span><button aria-label="关闭错误提示" onClick={() => setError("")}><X size={17}/></button></div>}{view === "access" ? <AccessWorkspace/> : showAdmin ? <AdminWorkspace {...props}/> : view === "evidence-review" ? <EvidenceReviewWorkspace {...props} key={`${state.access?.workspaceId ?? state.actor.id}:${state.actor.id}:${selectedId ?? "queue"}`}/> : view === "history-delete" ? <HistoryDeletionWorkspace {...props} key={`${state.access?.workspaceId ?? state.actor.id}:${state.actor.id}`} /> : observer ? <SharedLearning {...props} view={view}/> : learningRole ? view === "home" ? <Today {...props}/> : view === "camera" ? <CameraWorkspace {...props} onUploaded={refresh}/> : view === "chat" ? <CompanionChat {...props}/> : view === "shared-help" ? <div className="card empty-state"><ShieldCheck size={30}/><h2>此页面供已授权教师处理求助</h2><p>你可以在“和小禾聊聊”中查看自己提交的求助。</p><button className="btn btn-secondary" onClick={() => navigate("chat")}>查看我的求助</button></div> : <StudentLearning {...props} view={view}/> : <div className="card empty-state"><ShieldCheck size={33}/><h2>当前身份没有此页面权限</h2><button className="btn btn-secondary" onClick={() => navigate("access")}>查看成员与权限</button></div>}<footer className="workspace-footer"><Sprout size={14}/><span>学习有依据，成长有自己的节奏</span><span>原创演示课程 · 记录会真实保存</span></footer></main></div>
    {!showAdmin && <nav className={`mobile-nav ${!learningRole ? "restricted-mobile-nav" : ""}`} aria-label="手机导航">{mobileItems.map(({view:tab,label,icon:Icon}) => <button className={`${view===tab?"active":""} ${tab==="camera"?"camera-tab":""}`} key={tab} onClick={() => navigate(tab)}><span><Icon size={22}/></span><small>{label}</small></button>)}</nav>}{notice && <div className="toast" role="status"><Check size={18}/>{notice}</div>}
  </div>;
}
function Today({state,navigate,busy,mutate}: AppProps) {
  const completed = state.attempts.filter(a => a.status === "submitted" && !a.id.startsWith("demo-"));
  const independent = completed.reduce((sum,a) => sum + (effectiveAttempt(state, a.id)?.independentCorrect ?? 0),0);
  const pendingPaper = state.papers.find(p => p.status === "in_progress");
  const date = new Intl.DateTimeFormat("zh-CN",{timeZone:"Asia/Shanghai",month:"long",day:"numeric",weekday:"long"}).format(new Date());
  return <div className="today-view"><div className="page-heading"><div><p className="eyebrow"><Sun size={16}/>{date}</p><h1>{state.profile.name}，今天也一起进步吧<span className="heading-dot">。</span></h1><p className="muted">不用一次学很多，把一个小问题想明白就很好。</p></div><span className="badge subtle"><ShieldCheck size={15}/>以作答证据看进步</span></div>
    <section className="hero-card"><div className="hero-copy"><span className="hero-kicker"><Sparkles size={16}/>今天的一小步</span><h2>刚好适合你的，<br/>下一份小练习。</h2><p>根据当前课程和需要巩固的知识点，<br/>挑选一份练习，先看看，再开始。</p><div className="hero-actions"><button className="btn btn-primary" disabled={busy} onClick={async () => {const data=await mutate("paper.create",{count:6,knowledgeIds:["division","geometry"],difficulty:"基础",kind:"practice"});if(data?.result?.id)navigate("papers",data.result.id);}}><Sparkles size={18}/>一键出题<ArrowRight size={18}/></button><button className="hero-secondary" onClick={() => navigate("papers","configure")}>自己选一选<ChevronRight size={16}/></button></div><span className="hero-detail"><span/>数学 · 整数运算与图形 · 约 12 分钟</span></div><img className="hero-art" src="/illustrations/xiaohe-companion.png" alt="小禾与陪伴学习的小芽精灵"/><div className="hero-art-note"><Sprout size={18}/>每一步，都算数</div></section>
    <div className="quick-actions"><button className="quick-card peach" onClick={() => navigate("camera")}><span className="quick-icon"><Camera size={27}/></span><span><strong>拍一道题，找找方法</strong><small>拍照核对后，看看不同的思路</small></span><ArrowRight size={19}/></button><button className="quick-card lavender" onClick={() => navigate("chat")}><span className="quick-icon"><MessageCircle size={27}/></span><span><strong>和小禾聊一聊</strong><small>说说你卡在哪一步，一起想办法</small></span><ArrowRight size={19}/></button></div>
    <div className="dashboard-columns"><div><div className="section-heading"><h2>继续我的学习</h2><button className="text-link" onClick={() => navigate("papers")}>全部练习<ChevronRight size={16}/></button></div><section className="card continue-card"><div className="book-icon"><BookOpen size={29}/></div><div className="continue-copy"><span className="eyebrow">数学 · 当前课程</span><h3>{state.profile.course === "数学" ? "整数运算与图形" : state.profile.course}</h3><p className="muted">{pendingPaper ? "上次的答案和帮助记录已经保存，接着来吧。" : "估一估，再验一验。让每一步都有自己的理由。"}</p><div className="topic-pills"><span>估商与检验</span><span>周长与面积</span></div></div><button className="btn btn-secondary" onClick={() => navigate(pendingPaper?"practice":"knowledge",pendingPaper?.id)}>{pendingPaper?"继续练习":"查看课程"}<ArrowRight size={16}/></button></section><div className="section-heading spaced"><h2>知识点的小进步</h2><button className="text-link" onClick={() => navigate("knowledge")}>查看依据<ChevronRight size={16}/></button></div><section className="card knowledge-preview">{state.knowledge.slice(0,3).map(k => <button key={k.id} className="knowledge-preview-row" onClick={() => navigate("knowledge",k.id)}><span className={`knowledge-dot ${k.id}`}/><span><strong>{k.name}</strong><small>{k.stars===null?"积累一些新题记录，再看看":"每颗星，都有作答记录作为依据"}</small></span><span className="star-group" aria-label={k.stars===null?"待评估":`${k.stars}星`}>{k.stars===null?<span className="badge subtle">{k.learningState==="upcoming"?"未学习":"待评估"}</span>:Array.from({length:5},(_,i)=><Star key={i} size={17} fill={i<k.stars! ? "currentColor":"none"} className={i<k.stars!?"filled":"unfilled"}/>)}</span><ChevronRight size={16}/></button>)}</section></div><aside className="dashboard-right"><section className="growth-card"><span className="eyebrow"><Leaf size={15}/>慢慢积累的成长</span><h2>看见自己的努力</h2><div className="growth-number">{independent}<small>题</small></div><p>当前有效的独立答对</p><div className="growth-divider"/><div className="growth-stats"><div><strong>{completed.length}</strong><span>完成练习</span></div><div><strong>{effectiveMistakes(state).filter(m=>!m.correctedAt&&!m.id.startsWith("demo-")).length}</strong><span>等待巩固</span></div></div><button className="text-link" onClick={() => navigate("profile")}>打开学习档案<ArrowRight size={15}/></button></section><section className="gentle-card"><div className="row"><span className="soft-icon"><Sprout size={22}/></span><strong>小禾想对你说</strong></div><p>“会用提示也很好。下次换一道新题，再试试自己完成。”</p><span className="muted">看懂了，和独立做到了，都值得认真记录。</span></section></aside></div>
  </div>;
}
