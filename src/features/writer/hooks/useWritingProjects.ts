import { useEffect, useMemo, useRef, useState } from "react";
import { api } from "../../../shared/api";
import { useI18n } from "../../../shared/i18n";
import { DEFAULT_CHAPTER_SETTINGS, DEFAULT_PROJECT_NOTES } from "../constants";
import type {
  BookProject,
  Chapter,
  Scene,
  WriterChapterSettings,
  WriterProjectNotes,
  WriterDocxParseMode,
  RagCollection,
  ProviderModel,
  ProviderProfile,
  WriterSummaryLens,
  WriterSummaryLensScope,
  ConsistencyIssue
} from "../../../shared/types/contracts";
import type { BackgroundTask } from "../types";
import { addBackgroundTask, updateBackgroundTask } from "../taskStore";
import { clamp01 } from "../utils";
import { useBackgroundTasks } from "../../../shared/backgroundTasks";
import { triggerBlobDownload } from "../../../shared/download";

export function useWritingProjects() {
  const { t } = useI18n();
  const [projects, setProjects] = useState<BookProject[]>([]);
  const [activeProject, setActiveProject] = useState<BookProject | null>(null);
  const [projectNotes, setProjectNotes] = useState<WriterProjectNotes>({ ...DEFAULT_PROJECT_NOTES });
  const [chapters, setChapters] = useState<Chapter[]>([]);
  const [scenes, setScenes] = useState<Scene[]>([]);
  const [issues, setIssues] = useState<ConsistencyIssue[]>([]);
  const [chapterPrompt, setChapterPrompt] = useState("");
  const [selectedChapterId, setSelectedChapterId] = useState<string | null>(null);
  const [selectedSceneId, setSelectedSceneId] = useState<string | null>(null);
  const [generationLog, setGenerationLog] = useState<string[]>([]);
  const [busy, setBusy] = useState(false);
  const [bookSearchQuery, setBookSearchQuery] = useState("");
  const [renamingProjectId, setRenamingProjectId] = useState<string | null>(null);
  const [renamingProjectTitle, setRenamingProjectTitle] = useState("");
  const [renamingChapterId, setRenamingChapterId] = useState<string | null>(null);
  const [renamingChapterTitle, setRenamingChapterTitle] = useState("");
  const [chapterSettings, setChapterSettings] = useState<WriterChapterSettings>({ ...DEFAULT_CHAPTER_SETTINGS });
  const [providers, setProviders] = useState<ProviderProfile[]>([]);
  const [models, setModels] = useState<ProviderModel[]>([]);
  const [ragCollections, setRagCollections] = useState<RagCollection[]>([]);
  const [writerRagEnabled, setWriterRagEnabled] = useState(false);
  const [writerRagCollectionIds, setWriterRagCollectionIds] = useState<string[]>([]);
  const [writerProviderId, setWriterProviderId] = useState("");
  const [writerModelId, setWriterModelId] = useState("");
  const [activeModelLabel, setActiveModelLabel] = useState("");
  const [loadingModels, setLoadingModels] = useState(false);
  const [summaryLenses, setSummaryLenses] = useState<WriterSummaryLens[]>([]);
  const [lensNameDraft, setLensNameDraft] = useState("");
  const [lensPromptDraft, setLensPromptDraft] = useState("");
  const [lensScopeDraft, setLensScopeDraft] = useState<WriterSummaryLensScope>("project");
  const [lensTargetDraft, setLensTargetDraft] = useState("");
  const [lensBusyId, setLensBusyId] = useState<string | null>(null);
  const [lensOutputExpanded, setLensOutputExpanded] = useState<Record<string, boolean>>({});
  const [docxParseMode, setDocxParseMode] = useState<WriterDocxParseMode>("auto");
  const [docxImportAsBook, setDocxImportAsBook] = useState(false);
  const [docxBookNameDraft, setDocxBookNameDraft] = useState("");

  const bgTasks = useBackgroundTasks();
  const chapterSettingsTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const projectNotesTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const docxImportInputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    api.writerProjectList().then(setProjects);
    api.providerList().then(setProviders).catch(() => {});
    api.ragCollectionList().then(setRagCollections).catch(() => {});
  }, []);

  useEffect(() => {
    if (!writerProviderId) {
      setModels([]);
      setWriterModelId("");
      return;
    }
    setLoadingModels(true);
    api.providerFetchModels(writerProviderId)
      .then((list) => {
        setModels(list);
        setWriterModelId((prev) => {
          if (list.length === 0) return "";
          return list.some((m) => m.id === prev) ? prev : list[0].id;
        });
      })
      .catch(() => {
        setModels([]);
        setWriterModelId("");
      })
      .finally(() => setLoadingModels(false));
  }, [writerProviderId]);

  function log(msg: string) {
    setGenerationLog((prev) => [`[${new Date().toLocaleTimeString()}] ${msg}`, ...prev]);
  }

  function startBgTask(type: BackgroundTask["type"], label: string): string {
    const id = `task-${Date.now()}`;
    const task: BackgroundTask = { id, scope: "writing", type, label, startedAt: Date.now(), status: "running" };
    addBackgroundTask(task);
    return id;
  }

  function finishBgTask(id: string, status: "done" | "error", result?: string) {
    updateBackgroundTask(id, { status, result });
  }

  async function createProject() {
    const defaultName = `${t("writing.defaultBookPrefix")} ${projects.length + 1}`;
    const project = await api.writerProjectCreate(defaultName, t("writing.defaultProjectDescription"), []);
    setProjects((prev) => [project, ...prev]);
    setActiveProject(project);
    setProjectNotes(project.notes || { ...DEFAULT_PROJECT_NOTES });
    setRenamingProjectId(null);
    setRenamingProjectTitle("");
    setRenamingChapterId(null);
    setRenamingChapterTitle("");
    setChapters([]);
    setScenes([]);
    setSelectedChapterId(null);
    setSelectedSceneId(null);
    setChapterSettings({ ...DEFAULT_CHAPTER_SETTINGS });
    setSummaryLenses([]);
    setLensOutputExpanded({});
    setLensNameDraft("");
    setLensPromptDraft("");
    setLensScopeDraft("project");
    setLensTargetDraft("");
    setWriterRagEnabled(false);
    setWriterRagCollectionIds([]);
  }

  function startRenameProject(project: BookProject) {
    setRenamingProjectId(project.id);
    setRenamingProjectTitle(project.name || "");
  }

  function cancelRenameProject() {
    setRenamingProjectId(null);
    setRenamingProjectTitle("");
  }

  async function submitRenameProject(project: BookProject) {
    const nextName = renamingProjectTitle.trim();
    if (!nextName || nextName === project.name) {
      cancelRenameProject();
      return;
    }
    try {
      const updated = await api.writerProjectUpdate(project.id, { name: nextName });
      if (activeProject?.id === updated.id) setActiveProject(updated);
      setProjects((prev) => prev.map((p) => (p.id === updated.id ? updated : p)));
      log(`${t("writing.logBookRenamed")}: ${updated.name}`);
      cancelRenameProject();
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
    }
  }

  async function deleteProject(project: BookProject) {
    if (!confirm(t("writing.confirmDeleteBook"))) return;
    const deletingId = project.id;
    const deletingName = project.name;
    try {
      await api.writerProjectDelete(deletingId);
      const remaining = projects.filter((p) => p.id !== deletingId);
      setProjects(remaining);
      cancelRenameProject();
      if (remaining.length > 0 && activeProject?.id === deletingId) {
        await openProject(remaining[0]);
      } else if (remaining.length === 0) {
        setActiveProject(null);
        setProjectNotes({ ...DEFAULT_PROJECT_NOTES });
        setChapters([]);
        setScenes([]);
        setSelectedChapterId(null);
        setSelectedSceneId(null);
        setChapterSettings({ ...DEFAULT_CHAPTER_SETTINGS });
        setSummaryLenses([]);
        setLensOutputExpanded({});
        setWriterRagEnabled(false);
        setWriterRagCollectionIds([]);
      }
      log(`${t("writing.logBookDeleted")}: ${deletingName}`);
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
    }
  }

  async function openProject(project: BookProject) {
    const [loaded, lenses, ragBinding] = await Promise.all([
      api.writerProjectOpen(project.id),
      api.writerSummaryLensList(project.id).catch(() => []),
      api.writerProjectGetRag(project.id).catch(() => ({ enabled: false, collectionIds: [], updatedAt: null }))
    ]);
    cancelRenameProject();
    setRenamingChapterId(null);
    setRenamingChapterTitle("");
    setActiveProject(loaded.project);
    setProjectNotes(loaded.project.notes || { ...DEFAULT_PROJECT_NOTES });
    setChapters(loaded.chapters);
    setScenes(loaded.scenes);
    setSelectedChapterId(loaded.chapters[0]?.id ?? null);
    setSelectedSceneId(loaded.scenes[0]?.id ?? null);
    setChapterSettings(loaded.chapters[0]?.settings ?? { ...DEFAULT_CHAPTER_SETTINGS });
    setSummaryLenses(lenses);
    setLensOutputExpanded(Object.fromEntries(lenses.map((lens) => [lens.id, false])));
    setWriterRagEnabled(ragBinding.enabled === true);
    setWriterRagCollectionIds(Array.isArray(ragBinding.collectionIds) ? ragBinding.collectionIds : []);
  }

  function startRenameChapter(chapter: Chapter) {
    setRenamingChapterId(chapter.id);
    setRenamingChapterTitle(chapter.title || "");
  }

  function cancelRenameChapter() {
    setRenamingChapterId(null);
    setRenamingChapterTitle("");
  }

  async function submitRenameChapter(chapter: Chapter) {
    const nextTitle = renamingChapterTitle.trim();
    if (!nextTitle || nextTitle === chapter.title) {
      cancelRenameChapter();
      return;
    }
    try {
      const updated = await api.writerChapterUpdate(chapter.id, { title: nextTitle });
      setChapters((prev) => prev.map((item) => (item.id === updated.id ? updated : item)));
      if (selectedChapterId === updated.id) {
        setChapterSettings(updated.settings ?? { ...DEFAULT_CHAPTER_SETTINGS });
      }
      log(`${t("writing.logChapterRenamed")}: ${updated.title}`);
      cancelRenameChapter();
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
    }
  }

  async function deleteChapter(chapter: Chapter) {
    if (!confirm(t("writing.confirmDeleteChapter"))) return;
    const deletingId = chapter.id;
    const deletingTitle = chapter.title;
    const removedSceneIds = new Set(
      scenes.filter((scene) => scene.chapterId === deletingId).map((scene) => scene.id)
    );
    try {
      await api.writerChapterDelete(deletingId);
      const remainingChapters = chapters
        .filter((item) => item.id !== deletingId)
        .map((item, index) => ({ ...item, position: index + 1 }));
      const remainingScenes = scenes.filter((scene) => scene.chapterId !== deletingId);
      const nextSelectedChapterId =
        selectedChapterId === deletingId
          ? (remainingChapters[0]?.id ?? null)
          : selectedChapterId;
      let nextSelectedSceneId = selectedSceneId;
      if (!nextSelectedSceneId || removedSceneIds.has(nextSelectedSceneId)) {
        if (nextSelectedChapterId) {
          nextSelectedSceneId =
            remainingScenes.find((scene) => scene.chapterId === nextSelectedChapterId)?.id ?? null;
        } else {
          nextSelectedSceneId = null;
        }
      }
      setChapters(remainingChapters);
      setScenes(remainingScenes);
      setSelectedChapterId(nextSelectedChapterId);
      setSelectedSceneId(nextSelectedSceneId);
      if (renamingChapterId === deletingId) {
        cancelRenameChapter();
      }
      log(`${t("writing.logChapterDeleted")}: ${deletingTitle}`);
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
    }
  }

  async function createChapter() {
    if (!activeProject) return;
    const chapter = await api.writerChapterCreate(activeProject.id, `${t("writing.defaultChapterPrefix")} ${chapters.length + 1}`);
    setRenamingChapterId(null);
    setRenamingChapterTitle("");
    setChapters((prev) => [...prev, chapter]);
    setSelectedChapterId(chapter.id);
    setChapterSettings(chapter.settings ?? { ...DEFAULT_CHAPTER_SETTINGS });
    log(`${t("writing.logChapterCreated")}: ${chapter.title}`);
  }

  async function generateNextChapter() {
    if (!activeProject || busy) return;
    setBusy(true);
    const taskLabel = chapterPrompt.trim()
      ? `${t("writing.taskGenerateNextChapter")}: "${chapterPrompt.trim().slice(0, 30)}..."`
      : t("writing.taskGenerateNextChapter");
    const taskId = startBgTask("generate", taskLabel);
    log(t("writing.working"));
    try {
      const result = await api.writerGenerateNextChapter(activeProject.id, chapterPrompt.trim() || undefined);
      const { chapter, scene } = result;
      setRenamingChapterId(null);
      setRenamingChapterTitle("");
      setChapters((prev) => [...prev, chapter].sort((a, b) => a.position - b.position));
      setScenes((prev) => [...prev, scene]);
      setSelectedChapterId(chapter.id);
      setSelectedSceneId(scene.id);
      setChapterSettings(chapter.settings ?? { ...DEFAULT_CHAPTER_SETTINGS });
      setChapterPrompt("");
      log(`${t("writing.logNextChapterGenerated")}: ${chapter.title}`);
      finishBgTask(taskId, "done", chapter.title);
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
      finishBgTask(taskId, "error", String(err));
    } finally {
      setBusy(false);
    }
  }

  async function generateDraft() {
    if (!selectedChapterId || busy) return;
    setBusy(true);
    const taskId = startBgTask("generate", `${t("writing.taskGenerate")}: "${chapterPrompt.slice(0, 30)}..."`);
    log(t("writing.working"));
    try {
      const scene = await api.writerGenerateDraft(selectedChapterId, chapterPrompt);
      setScenes((prev) => [...prev, scene]);
      setSelectedSceneId(scene.id);
      log(`${t("writing.logDraftGenerated")}: ${scene.title}`);
      finishBgTask(taskId, "done", scene.title);
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
      finishBgTask(taskId, "error", String(err));
    }
    setBusy(false);
  }

  async function runConsistency() {
    if (!activeProject) return;
    const taskId = startBgTask("consistency", t("writing.taskConsistency"));
    const report = await api.writerConsistencyRun(activeProject.id);
    setIssues(report);
    log(`${t("writing.logConsistencyFound")}: ${report.length}`);
    finishBgTask(taskId, "done", `${report.length} ${t("writing.issuesCount")}`);
  }

  async function expandScene() {
    if (!selectedSceneId || busy) return;
    setBusy(true);
    const taskId = startBgTask("expand", t("writing.taskExpand"));
    log(t("writing.working"));
    try {
      const scene = await api.writerSceneExpand(selectedSceneId);
      setScenes((prev) => prev.map((s) => (s.id === scene.id ? scene : s)));
      log(t("writing.logSceneExpanded"));
      finishBgTask(taskId, "done");
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
      finishBgTask(taskId, "error", String(err));
    }
    setBusy(false);
  }

  async function rewriteScene() {
    if (!selectedSceneId || busy) return;
    setBusy(true);
    const tone = (chapterSettings.tone || DEFAULT_CHAPTER_SETTINGS.tone).trim();
    const taskId = startBgTask("rewrite", `${t("writing.taskRewrite")} (${tone})`);
    log(`${t("writing.rewrite")} (${tone})...`);
    try {
      const scene = await api.writerSceneRewrite(selectedSceneId);
      setScenes((prev) => prev.map((s) => (s.id === scene.id ? scene : s)));
      log(t("writing.logSceneRewritten"));
      finishBgTask(taskId, "done");
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
      finishBgTask(taskId, "error", String(err));
    }
    setBusy(false);
  }

  async function summarizeScene() {
    if (!selectedSceneId || busy) return;
    setBusy(true);
    const taskId = startBgTask("summarize", t("writing.taskSummarize"));
    log(t("writing.working"));
    try {
      const summary = await api.writerSceneSummarize(selectedSceneId);
      log(`${t("writing.logSummary")}: ${summary}`);
      finishBgTask(taskId, "done", String(summary).slice(0, 100));
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
      finishBgTask(taskId, "error", String(err));
    }
    setBusy(false);
  }

  async function deleteScene(scene: Scene) {
    if (!confirm(t("writing.confirmDeleteScene"))) return;
    const deletingId = scene.id;
    const deletingTitle = scene.title;
    try {
      await api.writerSceneDelete(deletingId);
      const remainingScenes = scenes.filter((item) => item.id !== deletingId);
      const nextSelectedSceneId =
        selectedSceneId === deletingId
          ? (remainingScenes.find((item) => item.chapterId === scene.chapterId)?.id
            ?? remainingScenes[0]?.id
            ?? null)
          : selectedSceneId;
      setScenes(remainingScenes);
      setSelectedSceneId(nextSelectedSceneId);
      log(`${t("writing.logSceneDeleted")}: ${deletingTitle}`);
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
    }
  }

  async function applyWriterModel() {
    if (!writerProviderId || !writerModelId) return;
    try {
      const result = await api.providerActivateModel(writerProviderId, writerModelId);
      const updated = result.settings;
      if (updated.activeProviderId) setWriterProviderId(updated.activeProviderId);
      if (result.actualModelId) {
        setWriterModelId(result.actualModelId);
        setActiveModelLabel(result.activeModelLabel || result.actualModelId);
      } else {
        setActiveModelLabel(writerModelId);
      }
      log(`${t("writing.modelSet")}: ${result.activeModelLabel || updated.activeModel || writerModelId}`);
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
    }
  }

  async function updateWriterRag(nextEnabled: boolean, nextCollectionIds: string[]) {
    if (!activeProject) return;
    const normalizedIds = Array.from(new Set(nextCollectionIds.filter(Boolean)));
    setWriterRagEnabled(nextEnabled);
    setWriterRagCollectionIds(normalizedIds);
    try {
      const binding = await api.writerProjectSaveRag(activeProject.id, nextEnabled, normalizedIds);
      setWriterRagEnabled(binding.enabled === true);
      setWriterRagCollectionIds(Array.isArray(binding.collectionIds) ? binding.collectionIds : []);
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
    }
  }

  function updateProjectNotes(patch: Partial<WriterProjectNotes>) {
    if (!activeProject) return;
    const next: WriterProjectNotes = { ...projectNotes, ...patch };
    setProjectNotes(next);
    setActiveProject((prev) => (prev ? { ...prev, notes: next } : prev));
    setProjects((prev) => prev.map((project) => (
      project.id === activeProject.id ? { ...project, notes: next } : project
    )));
    if (projectNotesTimerRef.current) clearTimeout(projectNotesTimerRef.current);
    projectNotesTimerRef.current = setTimeout(() => {
      void api.writerProjectUpdateNotes(activeProject.id, next).catch((err) => {
        log(`${t("writing.logError")}: ${String(err)}`);
      });
    }, 350);
  }

  async function summarizeBook(force = false) {
    if (!activeProject || busy) return;
    setBusy(true);
    const taskId = startBgTask("summarize", t("writing.summarizeBook"));
    try {
      const result = await api.writerProjectSummarize(activeProject.id, force);
      updateProjectNotes({ summary: result.summary });
      log(`${t("writing.logSummary")}: ${result.cached ? t("writing.summaryCached") : t("writing.summaryRefreshed")}`);
      finishBgTask(taskId, "done", `${result.chapterCount} ${t("writing.chShort")}`);
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
      finishBgTask(taskId, "error", String(err));
    } finally {
      setBusy(false);
    }
  }

  async function exportMarkdown() {
    if (!activeProject) return;
    try {
      const blob = await api.writerExportMarkdownDownload(activeProject.id);
      const filename = `${(activeProject.name || "book").replace(/[<>:"/\\|?*\u0000-\u001F]/g, " ").trim() || "book"}.md`;
      await triggerBlobDownload(blob, filename);
      log(`${t("writing.logMarkdownExported")}: ${filename}`);
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
    }
  }

  async function exportDocx() {
    if (!activeProject) return;
    try {
      const blob = await api.writerExportDocxDownload(activeProject.id);
      const filename = `${(activeProject.name || "book").replace(/[<>:"/\\|?*\u0000-\u001F]/g, " ").trim() || "book"}.docx`;
      await triggerBlobDownload(blob, filename);
      log(`${t("writing.logDocxExported")}: ${filename}`);
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
    }
  }

  async function handleDocxImport(file: File | null) {
    if (!file) return;
    if (!docxImportAsBook && !activeProject) return;
    if (busy) return;
    setBusy(true);
    const taskLabel = docxImportAsBook ? t("writing.importDocxAsBook") : t("writing.importDocx");
    const taskId = startBgTask("generate", `${taskLabel}: ${file.name}`);
    try {
      const base64Data = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onerror = () => reject(new Error("Failed to read DOCX file"));
        reader.onload = () => resolve(String(reader.result || ""));
        reader.readAsDataURL(file);
      });
      if (docxImportAsBook) {
        const payload = await api.writerImportDocxAsBook(
          base64Data,
          file.name,
          docxParseMode,
          docxBookNameDraft.trim() || undefined
        );
        setProjects((prev) => [payload.project, ...prev.filter((p) => p.id !== payload.project.id)]);
        await openProject(payload.project);
        setDocxBookNameDraft("");
        log(`${t("writing.logBookImported")}: ${payload.project.name} (${payload.chaptersCreated} ${t("writing.chShort")})`);
        finishBgTask(taskId, "done", `${payload.chaptersCreated}/${payload.scenesCreated}`);
      } else if (activeProject) {
        const payload = await api.writerProjectImportDocx(activeProject.id, base64Data, file.name, docxParseMode);
        await openProject(activeProject);
        log(`${t("writing.logDocxImported")}: ${payload.chaptersCreated} ${t("writing.chShort")}, ${payload.scenesCreated} ${t("writing.scenesShort")}`);
        finishBgTask(taskId, "done", `${payload.chaptersCreated}/${payload.scenesCreated}`);
      }
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
      finishBgTask(taskId, "error", String(err));
    } finally {
      setBusy(false);
      if (docxImportInputRef.current) {
        docxImportInputRef.current.value = "";
      }
    }
  }

  async function createSummaryLens() {
    if (!activeProject) return;
    const prompt = lensPromptDraft.trim();
    if (!prompt) {
      log(`${t("writing.logError")}: ${t("writing.lensPromptRequired")}`);
      return;
    }
    const targetId = lensScopeDraft === "project"
      ? null
      : (lensTargetDraft || (lensScopeDraft === "chapter" ? selectedChapterId : selectedSceneId) || null);
    try {
      const created = await api.writerSummaryLensCreate(activeProject.id, {
        name: lensNameDraft.trim() || t("writing.lensDefaultName"),
        prompt,
        scope: lensScopeDraft,
        targetId
      });
      setSummaryLenses((prev) => [created, ...prev]);
      setLensOutputExpanded((prev) => ({ ...prev, [created.id]: false }));
      setLensNameDraft("");
      setLensPromptDraft("");
      log(`${t("writing.logLensCreated")}: ${created.name}`);
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
    }
  }

  async function runSummaryLens(lensId: string, force = false) {
    if (!activeProject || lensBusyId) return;
    const lens = summaryLenses.find((item) => item.id === lensId);
    if (!lens) return;
    const taskId = startBgTask("summarize", `${t("writing.taskLensRun")}: ${lens.name}`);
    setLensBusyId(lensId);
    try {
      const result = await api.writerSummaryLensRun(activeProject.id, lensId, force);
      setSummaryLenses((prev) => prev.map((item) => (item.id === lensId ? result.lens : item)));
      setLensOutputExpanded((prev) => ({ ...prev, [lensId]: true }));
      log(`${t("writing.logLensReady")}: ${result.lens.name}${result.cached ? ` (${t("writing.summaryCached")})` : ""}`);
      finishBgTask(taskId, "done", `${Math.round(result.sourceChars / 1000)}k`);
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
      finishBgTask(taskId, "error", String(err));
    } finally {
      setLensBusyId(null);
    }
  }

  async function removeSummaryLens(lensId: string) {
    if (!activeProject) return;
    if (!confirm(t("writing.confirmDeleteLens"))) return;
    try {
      await api.writerSummaryLensDelete(activeProject.id, lensId);
      setSummaryLenses((prev) => prev.filter((item) => item.id !== lensId));
      setLensOutputExpanded((prev) => {
        const next = { ...prev };
        delete next[lensId];
        return next;
      });
      log(t("writing.logLensDeleted"));
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
    }
  }

  function loadLensToDraft(lens: WriterSummaryLens) {
    setLensNameDraft(lens.name);
    setLensPromptDraft(lens.prompt);
    setLensScopeDraft(lens.scope);
    setLensTargetDraft(lens.targetId || "");
  }

  async function toggleProjectCharacter(characterId: string) {
    if (!activeProject) return;
    const currentIds = activeProject.characterIds || [];
    const nextIds = currentIds.includes(characterId)
      ? currentIds.filter((id) => id !== characterId)
      : [...currentIds, characterId];
    try {
      const updated = await api.writerProjectSetCharacters(activeProject.id, nextIds);
      setActiveProject(updated);
      setProjects((prev) => prev.map((project) => (project.id === updated.id ? updated : project)));
      log(`${t("writing.logCastUpdated")} (${updated.characterIds.length} ${t("writing.charactersCountSuffix")})`);
    } catch (err) {
      log(`${t("writing.logError")}: ${String(err)}`);
    }
  }

  function updateSelectedChapterSettings(patch: Partial<WriterChapterSettings>) {
    if (!selectedChapterId) return;
    const merged: WriterChapterSettings = {
      ...chapterSettings,
      ...patch,
      creativity: clamp01(Number((patch.creativity ?? chapterSettings.creativity))),
      tension: clamp01(Number((patch.tension ?? chapterSettings.tension))),
      detail: clamp01(Number((patch.detail ?? chapterSettings.detail))),
      dialogue: clamp01(Number((patch.dialogue ?? chapterSettings.dialogue)))
    };
    setChapterSettings(merged);
    setChapters((prev) => prev.map((chapter) => (
      chapter.id === selectedChapterId ? { ...chapter, settings: merged } : chapter
    )));
    if (chapterSettingsTimerRef.current) clearTimeout(chapterSettingsTimerRef.current);
    chapterSettingsTimerRef.current = setTimeout(() => {
      api.writerChapterUpdateSettings(selectedChapterId, merged)
        .then((updatedChapter) => {
          setChapters((prev) => prev.map((chapter) => (
            chapter.id === updatedChapter.id ? updatedChapter : chapter
          )));
        })
        .catch((err) => log(`Error: ${String(err)}`));
    }, 250);
  }

  const selectedChapter = useMemo(
    () => chapters.find((chapter) => chapter.id === selectedChapterId) ?? null,
    [chapters, selectedChapterId]
  );

  const selectedScene = useMemo(() => scenes.find((s) => s.id === selectedSceneId) ?? null, [scenes, selectedSceneId]);

  const writerRagCollectionsAvailable = useMemo(
    () => ragCollections.filter((collection) => collection.scope === "global" || collection.scope === "writer"),
    [ragCollections]
  );

  const filteredProjects = useMemo(() => {
    const q = bookSearchQuery.trim().toLowerCase();
    if (!q) return projects;
    return projects.filter((project) => {
      const haystack = `${project.name || ""} ${project.description || ""}`.toLowerCase();
      return haystack.includes(q);
    });
  }, [projects, bookSearchQuery]);

  const selectedChapterScenes = useMemo(() => {
    if (!selectedChapterId) return [];
    return scenes.filter((scene) => scene.chapterId === selectedChapterId);
  }, [scenes, selectedChapterId]);

  const lensSceneOptions = useMemo(() => {
    if (lensScopeDraft === "chapter" && lensTargetDraft) {
      return scenes.filter((scene) => scene.chapterId === lensTargetDraft);
    }
    if (selectedChapterId) {
      return scenes.filter((scene) => scene.chapterId === selectedChapterId);
    }
    return scenes;
  }, [scenes, lensScopeDraft, lensTargetDraft, selectedChapterId]);

  const runningTasks = bgTasks.filter((task) => task.status === "running");

  // Sync effects
  useEffect(() => {
    setChapterSettings(selectedChapter?.settings ?? { ...DEFAULT_CHAPTER_SETTINGS });
  }, [selectedChapterId, selectedChapter]);

  useEffect(() => {
    setProjectNotes(activeProject?.notes || { ...DEFAULT_PROJECT_NOTES });
  }, [activeProject?.id]);

  useEffect(() => {
    if (lensScopeDraft === "project") {
      if (lensTargetDraft) setLensTargetDraft("");
      return;
    }
    if (lensScopeDraft === "chapter" && !lensTargetDraft && selectedChapterId) {
      setLensTargetDraft(selectedChapterId);
      return;
    }
    if (lensScopeDraft === "scene" && !lensTargetDraft && selectedSceneId) {
      setLensTargetDraft(selectedSceneId);
    }
  }, [lensScopeDraft, lensTargetDraft, selectedChapterId, selectedSceneId]);

  useEffect(() => {
    return () => {
      if (chapterSettingsTimerRef.current) clearTimeout(chapterSettingsTimerRef.current);
      if (projectNotesTimerRef.current) clearTimeout(projectNotesTimerRef.current);
    };
  }, []);

  return {
    // Data
    projects, activeProject, projectNotes, chapters, scenes, issues,
    chapterPrompt, selectedChapterId, selectedSceneId, generationLog, busy,
    bookSearchQuery, renamingProjectId, renamingProjectTitle,
    renamingChapterId, renamingChapterTitle, chapterSettings,
    providers, models, ragCollections, writerRagEnabled, writerRagCollectionIds,
    writerProviderId, writerModelId, activeModelLabel, loadingModels,
    summaryLenses, lensNameDraft, lensPromptDraft, lensScopeDraft, lensTargetDraft,
    lensBusyId, lensOutputExpanded, docxParseMode, docxImportAsBook, docxBookNameDraft,
    bgTasks, docxImportInputRef, lensSceneOptions,
    // Derived
    selectedChapter, selectedScene, writerRagCollectionsAvailable,
    filteredProjects, selectedChapterScenes, runningTasks,
    // Setters
    setChapterPrompt, setSelectedChapterId, setSelectedSceneId,
    setBookSearchQuery, setRenamingProjectTitle, setRenamingChapterTitle,
    setWriterProviderId, setWriterModelId, setWriterRagEnabled, setWriterRagCollectionIds,
    setLensNameDraft, setLensPromptDraft, setLensScopeDraft, setLensTargetDraft,
    setLensOutputExpanded, setDocxParseMode, setDocxImportAsBook, setDocxBookNameDraft,
    setScenes,
    // Actions
    log, createProject, startRenameProject, cancelRenameProject, submitRenameProject,
    deleteProject, openProject, startRenameChapter, cancelRenameChapter,
    submitRenameChapter, deleteChapter, createChapter, generateNextChapter,
    generateDraft, runConsistency, expandScene, rewriteScene, summarizeScene,
    deleteScene, applyWriterModel, updateWriterRag, updateProjectNotes,
    summarizeBook, exportMarkdown, exportDocx, handleDocxImport,
    createSummaryLens, runSummaryLens, removeSummaryLens, loadLensToDraft,
    toggleProjectCharacter, updateSelectedChapterSettings
  };
}
