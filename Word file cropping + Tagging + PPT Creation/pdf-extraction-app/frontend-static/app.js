const state = {
  jobId: null,
  projectName: "Generated Paper",
  allQuestions: [],
  selectedQuestions: [],
  replacingId: null,
  sourceSummary: null,
  sourceDocuments: [],
  analysisStatus: null,
  statusTimer: null,
  currentRoute: "home",
  builderMode: "pdf",
};

const routeButtons = Array.from(document.querySelectorAll("[data-route]"));
const screens = Array.from(document.querySelectorAll("[data-screen]"));
const builderModeButtons = Array.from(document.querySelectorAll("[data-builder-mode]"));
const builderPanels = Array.from(document.querySelectorAll("[data-builder-panel]"));
const uploadForm = document.getElementById("upload-form");
const builderForm = document.getElementById("builder-form");
const finalizeBtn = document.getElementById("finalize-btn");
const openLibraryBtn = document.getElementById("open-library-btn");
const closeLibraryBtn = document.getElementById("close-library-btn");
const libraryModal = document.getElementById("library-modal");
const libraryTitle = document.getElementById("library-title");
const librarySearch = document.getElementById("library-search");
const libraryDifficulty = document.getElementById("library-difficulty");
const libraryType = document.getElementById("library-type");
const summaryEl = document.getElementById("summary");
const selectedListEl = document.getElementById("selected-list");
const bankListEl = document.getElementById("bank-list");
const downloadsEl = document.getElementById("downloads");
const sourceFrame = document.getElementById("source-frame");
const sourceStatus = document.getElementById("source-status");
const launchSourceBtn = document.getElementById("launch-source-btn");
const sourceExternalLink = document.getElementById("source-external-link");
const sourceScreen = document.querySelector('[data-screen="paper-builder"]');
const finalFrame = document.getElementById("final-frame");
const finalStatus = document.getElementById("final-status");
const launchFinalBtn = document.getElementById("launch-final-btn");
const finalExternalLink = document.getElementById("final-external-link");
const SOURCE_WORKSPACE_URL = "https://script.google.com/a/macros/pw.live/s/AKfycbxXytlS5_cqTSVnoono-LLjEnJa1m4Xm0iuroV47yEJetyRa6hJjyAUFf5uybZ6HkjM/exec";
const FINAL_PACKAGE_URL = "https://paper2-assets.vercel.app/";

function byId(id) {
  return document.getElementById(id);
}

function setBuilderMode(mode) {
  state.builderMode = mode;
  document.body.dataset.builderMode = mode;
  builderPanels.forEach((panel) => {
    panel.classList.toggle("is-active", panel.dataset.builderPanel === mode);
  });
  builderModeButtons.forEach((button) => {
    if (button.classList.contains("builder-tab")) {
      button.classList.toggle("is-active", button.dataset.builderMode === mode);
    }
  });
  if (state.currentRoute === "paper-builder") {
    if (mode === "word") startSourceWorkspace();
  }
}

function setRoute(route) {
  state.currentRoute = route;
  document.body.dataset.route = route;
  screens.forEach((screen) => {
    const active = screen.dataset.screen === route;
    screen.classList.toggle("is-active", active);
    screen.hidden = !active;
  });
  routeButtons.forEach((button) => {
    const active = button.dataset.route === route;
    if (button.classList.contains("nav-btn")) {
      button.classList.toggle("is-active", active);
    }
  });
  if (libraryModal && !libraryModal.classList.contains("hidden")) {
    closeLibrary();
  }
  window.scrollTo({ top: 0, behavior: "smooth" });
  if (route === "final-package") {
    startFinalPackageWorkspace();
  }
  if (route === "paper-builder" && state.builderMode === "word") {
    startSourceWorkspace();
  }
}

function startSourceWorkspace() {
  if (!sourceFrame || !sourceStatus) return;
  if (sourceScreen) sourceScreen.classList.remove("workspace-loaded");
  sourceStatus.textContent = "Opening workspace...";
  sourceFrame.src = `${SOURCE_WORKSPACE_URL}${SOURCE_WORKSPACE_URL.includes("?") ? "&" : "?"}embed=${Date.now()}`;
  window.setTimeout(() => {
    if (state.currentRoute === "paper-builder" && state.builderMode === "word") {
      sourceStatus.textContent = "If the workspace does not render below, use Open Full View.";
    }
  }, 4000);
}

function startFinalPackageWorkspace() {
  if (!finalFrame || !finalStatus) return;
  finalStatus.textContent = "Opening workspace...";
  finalFrame.src = `${FINAL_PACKAGE_URL}${FINAL_PACKAGE_URL.includes("?") ? "&" : "?"}embed=${Date.now()}`;
  window.setTimeout(() => {
    if (state.currentRoute === "final-package") {
      finalStatus.textContent = "If the workspace does not render below, use Open Full View.";
    }
  }, 4000);
}

function setBusy(button, busy, label) {
  if (!button) return;
  button.disabled = busy;
  if (busy) {
    button.dataset.originalText = button.textContent;
    button.textContent = label;
  } else if (button.dataset.originalText) {
    button.textContent = button.dataset.originalText;
  }
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function renderSummary(summary, documents = []) {
  if (!summary || !summary.total_questions) {
    summaryEl.textContent = "No data yet.";
    return;
  }

  const docPills = documents.map((doc) => `<span class="pill">${escapeHtml(doc.original_filename)}: ${doc.question_count}</span>`).join("");
  const diffPills = Object.entries(summary.difficulty_counts || {}).map(([key, value]) => `<span class="pill">${escapeHtml(key)}: ${value}</span>`).join("");
  const typePills = Object.entries(summary.question_type_counts || {})
    .filter(([key, value]) => key !== "Unknown" && value > 0)
    .map(([key, value]) => `<span class="pill">${escapeHtml(key)}: ${value}</span>`)
    .join("");

  summaryEl.innerHTML = `
    <div class="summary-grid">
      <div>
        <strong>Total Questions</strong>
        <div>${summary.total_questions}</div>
      </div>
      <div>
        <strong>Uploaded PDFs</strong>
        <div>${documents.length}</div>
      </div>
    </div>
    <h3>Documents</h3>
    <div class="pill-box">${docPills || "<span class='pill'>None</span>"}</div>
    <h3>Difficulty Split</h3>
    <div class="pill-box">${diffPills}</div>
    <h3>Question Types</h3>
    <div class="pill-box">${typePills || "<span class='pill'>Pending</span>"}</div>
  `;
}

function setBuilderEnabled(enabled) {
  const submitBtn = builderForm.querySelector("button");
  submitBtn.disabled = !enabled;
}

function updateStatusText() {
  if (!state.analysisStatus) {
    downloadsEl.textContent = "Exports will appear here.";
    return;
  }

  const status = state.analysisStatus.status;
  const done = state.analysisStatus.processed_questions || 0;
  const total = state.analysisStatus.total_questions || 0;

  if (status === "queued") {
    downloadsEl.textContent = `Extraction ready. Gemini tagging queued for ${total} questions.`;
    return;
  }
  if (status === "running") {
    downloadsEl.textContent = `Tagging questions ${done}/${total}...`;
    return;
  }
  if (status === "completed") {
    downloadsEl.textContent = "Analysis ready.";
    return;
  }
  if (status === "failed") {
    downloadsEl.textContent = `Analysis failed${state.analysisStatus.error ? `: ${state.analysisStatus.error}` : "."}`;
    return;
  }
  downloadsEl.textContent = "Exports will appear here.";
}

async function pollJobStatus() {
  if (!state.jobId) return;
  try {
    const response = await fetch(`/api/jobs/${state.jobId}/status`);
    if (!response.ok) {
      throw new Error("Status fetch failed.");
    }
    const data = await response.json();
    state.analysisStatus = data.analysis_status || null;
    state.sourceSummary = data.summary || state.sourceSummary;
    state.sourceDocuments = data.documents || state.sourceDocuments;
    renderSummary(state.sourceSummary, state.sourceDocuments);
    updateStatusText();
    setBuilderEnabled(state.analysisStatus?.status === "completed");

    if (state.analysisStatus?.status === "completed") {
      clearInterval(state.statusTimer);
      state.statusTimer = null;
      const fullJob = await fetch(`/api/jobs/${state.jobId}`).then((res) => res.json());
      state.allQuestions = fullJob.questions || state.allQuestions;
      syncQuestionCollections();
    }
    if (state.analysisStatus?.status === "failed" && state.statusTimer) {
      clearInterval(state.statusTimer);
      state.statusTimer = null;
    }
  } catch (error) {
    console.error(error);
  }
}

function questionCard(question, mode) {
  const actions = [];
  if (mode === "selected") {
    actions.push(`<button class="ghost" data-action="move-up" data-id="${question.id}">Move Up</button>`);
    actions.push(`<button class="ghost" data-action="move-down" data-id="${question.id}">Move Down</button>`);
    actions.push(`<button class="ghost" data-action="replace" data-id="${question.id}">Replace</button>`);
    actions.push(`<button class="warn" data-action="remove" data-id="${question.id}">Remove</button>`);
  } else {
    const disabled = state.selectedQuestions.some((item) => item.id === question.id) ? "disabled" : "";
    const label = state.replacingId ? "Use as Replacement" : "Add to Paper";
    actions.push(`<button class="ghost" data-action="add" data-id="${question.id}" ${disabled}>${label}</button>`);
  }

  return `
    <article class="question-card">
      <img src="${question.image_url}" alt="${escapeHtml(question.display_label)}">
      <div class="body">
        <strong>${escapeHtml(question.display_label)}</strong>
        <div class="meta">
          <span>${escapeHtml(question.question_type || "Unknown")}</span>
          <span>${escapeHtml(question.difficulty || "Medium")}</span>
          <span>${escapeHtml(question.subject || "No subject")}</span>
        </div>
        <div>${escapeHtml(question.chapter || question.topic || question.notes || "")}</div>
        <div class="actions">${actions.join("")}</div>
      </div>
    </article>
  `;
}

function renderSelectedQuestions() {
  if (!state.selectedQuestions.length) {
    selectedListEl.className = "card-grid empty";
    selectedListEl.textContent = "No questions selected yet.";
  } else {
    selectedListEl.className = "card-grid";
    selectedListEl.innerHTML = state.selectedQuestions.map((question) => questionCard(question, "selected")).join("");
  }

  const enabled = Boolean(state.jobId && state.allQuestions.length);
  openLibraryBtn.disabled = !enabled;
  finalizeBtn.disabled = !(state.jobId && state.selectedQuestions.length);
}

function filteredBankQuestions() {
  const search = librarySearch.value.trim().toLowerCase();
  const difficulty = libraryDifficulty.value;
  const questionType = libraryType.value;

  return state.allQuestions.filter((question) => {
    if (difficulty && question.difficulty !== difficulty) return false;
    if (questionType && question.question_type !== questionType) return false;
    if (!search) return true;

    const haystack = [
      question.display_label,
      question.subject,
      question.chapter,
      question.topic,
      question.subtopic,
      question.question_type,
      question.difficulty,
      question.source_pdf,
    ].join(" ").toLowerCase();

    return haystack.includes(search);
  });
}

function renderBankQuestions() {
  if (!state.allQuestions.length) {
    bankListEl.className = "card-grid empty";
    bankListEl.textContent = "Upload PDFs to populate the question bank.";
    return;
  }

  const questions = filteredBankQuestions();
  if (!questions.length) {
    bankListEl.className = "card-grid empty";
    bankListEl.textContent = "No questions matched the current filters.";
    return;
  }

  bankListEl.className = "card-grid";
  bankListEl.innerHTML = questions.map((question) => questionCard(question, "bank")).join("");
}

function syncQuestionCollections() {
  renderSelectedQuestions();
  renderBankQuestions();
}

function openLibrary() {
  libraryTitle.textContent = state.replacingId ? "Replace Question" : "Question Library";
  libraryModal.classList.remove("hidden");
  renderBankQuestions();
}

function closeLibrary() {
  libraryModal.classList.add("hidden");
  if (!state.replacingId) {
    librarySearch.value = "";
    libraryDifficulty.value = "";
    libraryType.value = "";
  }
}

async function postForm(url, formData) {
  const response = await fetch(url, { method: "POST", body: formData });
  if (!response.ok) {
    const error = await response.json().catch(() => ({ detail: "Request failed." }));
    throw new Error(error.detail || "Request failed.");
  }
  return response.json();
}

async function postJson(url, payload) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    const error = await response.json().catch(() => ({ detail: "Request failed." }));
    throw new Error(error.detail || "Request failed.");
  }
  return response.json();
}

uploadForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const submitBtn = uploadForm.querySelector("button");
  setBusy(submitBtn, true, "Analyzing...");

  try {
    const files = byId("pdf-files").files;
    if (!files.length) {
      throw new Error("Please choose one or more PDFs.");
    }

    const formData = new FormData();
    Array.from(files).forEach((file) => formData.append("files", file));
    formData.append("start_page", byId("start-page").value || "1");

    const data = await postForm("/api/analyze-pdfs", formData);
    state.jobId = data.job_id;
    state.allQuestions = data.questions || [];
    state.selectedQuestions = [];
    state.replacingId = null;
    state.sourceSummary = data.summary || null;
    state.sourceDocuments = data.documents || [];
    state.analysisStatus = data.analysis_status || null;
    librarySearch.value = "";
    libraryDifficulty.value = "";
    libraryType.value = "";
    renderSummary(state.sourceSummary, state.sourceDocuments);
    syncQuestionCollections();
    updateStatusText();
    setBuilderEnabled(state.analysisStatus?.status === "completed");
    if (state.statusTimer) clearInterval(state.statusTimer);
    if (state.analysisStatus?.status !== "completed") {
      state.statusTimer = setInterval(pollJobStatus, 3000);
    }
  } catch (error) {
    alert(error.message);
  } finally {
    setBusy(submitBtn, false);
  }
});

builderForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!state.jobId) {
    alert("Analyze PDFs first.");
    return;
  }
  if (state.analysisStatus?.status !== "completed") {
    alert("Analysis is still running. Please wait for tagging to finish.");
    return;
  }

  const submitBtn = builderForm.querySelector("button");
  setBusy(submitBtn, true, "Selecting...");

  try {
    state.projectName = byId("project-name").value.trim() || "Generated Paper";
    const payload = {
      project_name: state.projectName,
      total_questions: Number(byId("total-questions").value || 0),
      difficulty_targets: {
        Easy: Number(byId("diff-easy").value || 0),
        Medium: Number(byId("diff-medium").value || 0),
        Hard: Number(byId("diff-hard").value || 0),
      },
      type_targets: {
        "Single Choice": Number(byId("type-single").value || 0),
        "Multiple Choice": Number(byId("type-multiple").value || 0),
        "Assertion & Reason": Number(byId("type-ar").value || 0),
        "Diagram Based": Number(byId("type-diagram").value || 0),
        Matching: Number(byId("type-matching").value || 0),
        Integer: Number(byId("type-integer").value || 0),
        Subjective: Number(byId("type-subjective").value || 0),
      },
    };

    const data = await postJson(`/api/jobs/${state.jobId}/build-paper`, payload);
    state.selectedQuestions = data.selected_questions || [];
    state.replacingId = null;
    renderSummary(state.sourceSummary, state.sourceDocuments);
    syncQuestionCollections();

    const shortageText = [];
    if (data.shortages?.total_shortfall) shortageText.push(`total shortfall ${data.shortages.total_shortfall}`);
    for (const [key, value] of Object.entries(data.shortages?.difficulty_shortfall || {})) shortageText.push(`${key} short by ${value}`);
    for (const [key, value] of Object.entries(data.shortages?.type_shortfall || {})) shortageText.push(`${key} short by ${value}`);
    downloadsEl.textContent = shortageText.length
      ? `Draft ready. ${shortageText.join(", ")}.`
      : "Draft ready.";
  } catch (error) {
    alert(error.message);
  } finally {
    setBusy(submitBtn, false);
  }
});

selectedListEl.addEventListener("click", (event) => {
  const button = event.target.closest("button");
  if (!button) return;

  const id = button.dataset.id;
  const action = button.dataset.action;
  const index = state.selectedQuestions.findIndex((question) => question.id === id);
  if (index === -1) return;

  if (action === "move-up" && index > 0) {
    [state.selectedQuestions[index - 1], state.selectedQuestions[index]] = [state.selectedQuestions[index], state.selectedQuestions[index - 1]];
  }
  if (action === "move-down" && index < state.selectedQuestions.length - 1) {
    [state.selectedQuestions[index + 1], state.selectedQuestions[index]] = [state.selectedQuestions[index], state.selectedQuestions[index + 1]];
  }
  if (action === "remove") {
    state.selectedQuestions.splice(index, 1);
  }
  if (action === "replace") {
    state.replacingId = state.selectedQuestions[index].id;
    downloadsEl.textContent = "Choose a replacement.";
    openLibrary();
  }

  syncQuestionCollections();
});

bankListEl.addEventListener("click", (event) => {
  const button = event.target.closest("button");
  if (!button || button.dataset.action !== "add") return;

  const question = state.allQuestions.find((item) => item.id === button.dataset.id);
  if (!question) return;
  if (state.selectedQuestions.some((item) => item.id === question.id)) return;

  if (state.replacingId) {
    const selectedIndex = state.selectedQuestions.findIndex((item) => item.id === state.replacingId);
    if (selectedIndex !== -1) {
      state.selectedQuestions[selectedIndex] = question;
    }
    state.replacingId = null;
    closeLibrary();
    downloadsEl.textContent = "Replacement applied.";
  } else {
    state.selectedQuestions.push(question);
    downloadsEl.textContent = "Question added.";
  }

  syncQuestionCollections();
});

finalizeBtn.addEventListener("click", async () => {
  if (!state.jobId || !state.selectedQuestions.length) {
    alert("Select at least one question first.");
    return;
  }

  setBusy(finalizeBtn, true, "Finalizing...");
  try {
    const data = await postJson(`/api/jobs/${state.jobId}/finalize`, {
      project_name: state.projectName,
      selected_ids: state.selectedQuestions.map((question) => question.id),
    });

    const links = [];
    links.push(`<a href="${data.downloads.zip}" target="_blank">Download ZIP Bundle</a>`);
    links.push(`<a href="${data.downloads.xlsx}" target="_blank">Download Tagging XLSX</a>`);
    links.push(`<a href="${data.downloads.word}" target="_blank">Download Word</a>`);
    links.push(`<a href="${data.downloads.pdf}" target="_blank">Download PDF</a>`);
    (data.downloads.ppt || []).forEach((url, index) => {
      links.push(`<a href="${url}" target="_blank">Download PPT ${index + 1}</a>`);
    });
    downloadsEl.innerHTML = links.join("");
  } catch (error) {
    alert(error.message);
  } finally {
    setBusy(finalizeBtn, false);
  }
});

openLibraryBtn.addEventListener("click", () => {
  state.replacingId = null;
  openLibrary();
});

closeLibraryBtn.addEventListener("click", closeLibrary);
libraryModal.addEventListener("click", (event) => {
  const closer = event.target.closest("[data-close-modal='true']");
  if (closer) closeLibrary();
});
librarySearch.addEventListener("input", renderBankQuestions);
libraryDifficulty.addEventListener("change", renderBankQuestions);
libraryType.addEventListener("change", renderBankQuestions);

setBuilderEnabled(false);
routeButtons.forEach((button) => {
  button.addEventListener("click", () => {
    if (button.dataset.builderMode) {
      setBuilderMode(button.dataset.builderMode);
    }
    setRoute(button.dataset.route);
  });
});
builderModeButtons.forEach((button) => {
  if (!button.classList.contains("builder-tab")) return;
  button.addEventListener("click", () => setBuilderMode(button.dataset.builderMode));
});
if (launchSourceBtn) {
  launchSourceBtn.addEventListener("click", () => {
    setBuilderMode("word");
    setRoute("paper-builder");
  });
}
if (sourceFrame) {
  sourceFrame.addEventListener("load", () => {
    if (sourceScreen && state.currentRoute === "paper-builder" && state.builderMode === "word") {
      sourceScreen.classList.add("workspace-loaded");
    }
    if (sourceStatus) sourceStatus.textContent = "Workspace loaded.";
  });
}
if (sourceExternalLink) {
  sourceExternalLink.href = SOURCE_WORKSPACE_URL;
}
if (launchFinalBtn) {
  launchFinalBtn.addEventListener("click", () => {
    setRoute("final-package");
  });
}
if (finalFrame) {
  finalFrame.addEventListener("load", () => {
    if (finalStatus) finalStatus.textContent = "Workspace loaded.";
  });
}
if (finalExternalLink) {
  finalExternalLink.href = FINAL_PACKAGE_URL;
}
setBuilderMode("pdf");
setRoute("home");
