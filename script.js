const state = {
  user: null,
  isTeacher: false,
  profileSlug: "",
  classroomId: "",
  classroomName: "",
  classrooms: [],
  classroomStudents: {},
  teacherResultsClassroomFilter: "",
  teacherResultsStudentSearch: "",
  customQuizzes: [],
  quizPosts: [],
  quizConfigs: {},
  nonogramTemplatesByQuiz: {},
  remoteAttemptsByQuiz: {},
  remoteAttemptCountsByQuiz: {},
  remoteAttemptsLoaded: false,
  retakeReleases: {},
  selectedQuizId: null,
  activeQuestions: [],
  studentName: "",
  currentQuestionIndex: 0,
  answers: {},
  isActive: false,
  isCancelled: false,
  timedOutQuestionId: null,
  editingQuizId: null,
  teacherAccessSource: "",
  isRenamingProfile: false,
  reviewBackScreen: "result",
  violationCount: 0,
  isViolationGraceActive: false,
  violationTimeoutId: null,
  violationIntervalId: null,
  violationDeadline: 0,
  attemptStartedAtMs: 0,
  attemptStartedAtLabel: "",
  baselineDPR: 1,
  skipQuestionUsesRemaining: 3,
  builderDraftGeneratedWithAi: false,
  builderDraftAiPrompt: ""
};

let pendingModalAction = null;
let lastHomeRenderSignature = "";
let lastContextMenuBlockedAt = 0;
let focusWaitPenaltyInFlight = false;
let lastFocusWaitPenaltyAt = 0;
let focusWaitGraceTimeoutId = null;
let focusWaitGraceIntervalId = null;
let focusWaitGraceDeadlineMs = 0;
let focusWaitGraceReason = "";
let userProfileListenerUnsubscribe = null;
let questionImageZoomState = {
  src: "",
  scale: 1,
  minScale: 1,
  maxScale: 4,
  translateX: 0,
  translateY: 0,
  dragging: false,
  dragStartX: 0,
  dragStartY: 0,
  startTranslateX: 0,
  startTranslateY: 0
};

const FOCUS_WAIT_PROCESS_NAME = "Espera em Foco da Turma";
const FOCUS_WAIT_RELEASE_BUTTON_LABEL = "Liberar modo de espera";
const WAIT_NONOGRAM_LOBBY_PREFIX = "labdg-nonogram";
const WAIT_NONOGRAM_GRID_SIZE = 8;
const NONOGRAM_TEMPLATE_GLOBAL_KEY = "global";
const NONOGRAM_DRAWING_SUGGESTIONS = ["Foguete", "Arvore", "Carro", "Gato", "Peixe", "Coroa", "Flor", "Lapis"];

const waitNonogramState = {
  active: false,
  quizId: "",
  lobbyKey: "",
  dateKey: "",
  puzzleIndex: 0,
  size: WAIT_NONOGRAM_GRID_SIZE,
  solution: [],
  marks: [],
  rowClues: [],
  colClues: [],
  errors: 0,
  roundScore: 0,
  totalScore: 0,
  solvedCount: 0,
  roundCounted: false,
  score: 0,
  progress: 0,
  teacherRevealMode: false,
  level: 1,
  currentPuzzleName: "",
  creatorName: "",
  creatorMatrix: [],
  completed: false,
  submitted: false,
  autoAdvanceTimeoutId: null
};

const authScreen = document.getElementById("auth-screen");
const bootScreen = document.getElementById("boot-screen");
const appHeader = document.querySelector(".app-header");
const homeScreen = document.getElementById("home-screen");
const teacherScreen = document.getElementById("teacher-screen");
const teacherClassroomsScreen = document.getElementById("teacher-classrooms-screen");
const builderScreen = document.getElementById("builder-screen");
const userChip = document.getElementById("user-chip");
const userChipPopup = document.getElementById("user-chip-popup");
const userChipPopupEmail = document.getElementById("user-chip-popup-email");
const userChipPopupRename = document.getElementById("user-chip-popup-rename");
const homeHeaderButton = document.getElementById("home-header-button");
const teacherPanelButton = document.getElementById("teacher-panel-button");
const teacherClassroomsButton = document.getElementById("teacher-classrooms-button");
const teacherBuilderButton = document.getElementById("teacher-builder-button");
const teacherNonogramEditorButton = document.getElementById("teacher-nonogram-editor-button");
const teacherRefreshButton = document.getElementById("teacher-refresh-button");
const teacherUsersList = document.getElementById("teacher-users-list");
const teacherMessage = document.getElementById("teacher-message");
const builderSubmitButton = document.getElementById("builder-submit-button");
const builderForm = document.getElementById("builder-form");
function getPendingUnloadKey() {
  return `quiz_pending_unload_cancellation_${state.user?.uid || "anon"}`;
}

function getPendingFocusWaitPenaltyKey() {
  return `quiz_pending_focus_wait_penalty_${state.user?.uid || "anon"}`;
}
const builderTitleInput = document.getElementById("builder-title");
const builderDescriptionInput = document.getElementById("builder-description");
const builderAiPromptInput = document.getElementById("builder-ai-prompt");
const builderAiApiKeyInput = document.getElementById("builder-ai-api-key");
const builderAiQuestionCountInput = document.getElementById("builder-ai-question-count");
const builderAiDifficultyInput = document.getElementById("builder-ai-difficulty");
const builderAiSchoolLevelInput = document.getElementById("builder-ai-school-level");
const builderAiGenerateButton = document.getElementById("builder-ai-generate-button");
const builderAiMessage = document.getElementById("builder-ai-message");
const builderDurationInput = document.getElementById("builder-duration");
const builderAlternativesVisibilityInput = document.getElementById("builder-alternatives-visibility");
const builderMaxAttemptsInput = document.getElementById("builder-max-attempts");
const builderQuestions = document.getElementById("builder-questions");
const builderAddQuestionButton = document.getElementById("builder-add-question-button");
const builderMessage = document.getElementById("builder-message");
const GEMINI_API_KEY_STORAGE_KEY = "labdg_gemini_api_key";
const logoutButton = document.getElementById("logout-button");
const authGoogleBtn = document.getElementById("auth-google");
const authMessage = document.getElementById("auth-message");
const profileScreen = document.getElementById("profile-screen");
const profileScreenTitle = document.getElementById("profile-screen-title");
const profileScreenSubtitle = document.getElementById("profile-screen-subtitle");
const profileSubmitButton = document.getElementById("profile-submit-button");
const profileCancelButton = document.getElementById("profile-cancel-button");
const profileForm = document.getElementById("profile-form");
const profileNameInput = document.getElementById("profile-name");
const profileClassroomSelect = document.getElementById("profile-classroom");
const profileMessage = document.getElementById("profile-message");
const teacherClassroomForm = document.getElementById("teacher-classroom-form");
const teacherClassroomNameInput = document.getElementById("teacher-classroom-name");
const teacherClassroomsList = document.getElementById("teacher-classrooms-list");
const teacherClassroomsMessage = document.getElementById("teacher-classrooms-message");

const screens = {
  start: document.getElementById("start-screen"),
  quiz: document.getElementById("quiz-screen"),
  cancel: document.getElementById("cancel-screen"),
  result: document.getElementById("result-screen"),
  review: document.getElementById("review-screen"),
  classrooms: document.getElementById("teacher-classrooms-screen"),
  builder: document.getElementById("builder-screen")
};

const quizGrid = document.getElementById("quiz-grid");
const quizSearch = document.getElementById("quiz-search");
const selectedQuizTitle = document.getElementById("selected-quiz-title");
const selectedQuizDescription = document.getElementById("selected-quiz-description");
const selectedQuizMeta = document.getElementById("selected-quiz-meta");
const startRulesList = document.getElementById("start-rules-list");
const timedWarningDisplay = document.getElementById("timed-warning-display");
const startForm = document.getElementById("start-form");
const startQuizButton = document.getElementById("start-quiz-button");
const startReloadButton = document.getElementById("start-reload-button");
const studentNameDisplay = document.getElementById("student-name-display");

const questionTitle = document.getElementById("question-title");
const questionDescription = document.getElementById("question-description");
const questionInfoArea = document.getElementById("question-info-area");
const questionForm = document.getElementById("question-form");
const progress = document.getElementById("progress");
const quizScreen = document.getElementById("quiz-screen");
const policyOverlay = document.getElementById("policy-overlay");
const policyCountdown = document.getElementById("policy-countdown");
const policyChancesText = document.getElementById("policy-chances-text");
const policyReturnButton = document.getElementById("policy-return-button");
const nextButton = document.getElementById("next-button");
const clearAnswerButton = document.getElementById("clear-answer-button");
const skipQuestionButton = document.getElementById("skip-question-button");
const cancelButton = document.getElementById("cancel-button");
const cancelModal = document.getElementById("cancel-modal");
const cancelModalTitle = document.getElementById("cancel-modal-title");
const cancelModalText = document.getElementById("cancel-modal-text");
const cancelModalConfirm = document.getElementById("cancel-modal-confirm");
const cancelModalCancel = document.getElementById("cancel-modal-cancel");
const questionImageViewer = document.getElementById("question-image-viewer");
const questionImageViewerTitle = document.getElementById("question-image-viewer-title");
const questionImageViewerCaption = document.getElementById("question-image-viewer-caption");
const questionImageViewerClose = document.getElementById("question-image-viewer-close");
const questionImageViewerStage = document.getElementById("question-image-viewer-stage");
const questionImageViewerImage = document.getElementById("question-image-viewer-image");
const questionImageZoomOut = document.getElementById("question-image-zoom-out");
const questionImageZoomReset = document.getElementById("question-image-zoom-reset");
const questionImageZoomIn = document.getElementById("question-image-zoom-in");
const cancelMessage = document.getElementById("cancel-message");
const cancelBackHomeButton = document.getElementById("cancel-back-home-button");
const cancelReviewButton = document.getElementById("cancel-review-button");

const resultAlreadyCompleted = document.getElementById("result-already-completed");
const resultQuizTitle = document.getElementById("result-quiz-title");
const resultStudentName = document.getElementById("result-student-name");
const resultScore = document.getElementById("result-score");
const resultPercent = document.getElementById("result-percent");
const resultDetails = document.getElementById("result-details");
const resultReviewButton = document.getElementById("result-review-button");
const resultBackHomeButton = document.getElementById("result-back-home-button");
const reviewSummaryMeta = document.getElementById("review-summary-meta");
const reviewList = document.getElementById("review-list");
const reviewBackButton = document.getElementById("review-back-button");

startForm.addEventListener("submit", handleStartQuiz);
nextButton.addEventListener("click", handleNextQuestion);
if (startReloadButton) {
  startReloadButton.addEventListener("click", () => { window.location.reload(); });
}
if (clearAnswerButton) {
  clearAnswerButton.addEventListener("click", () => {
    if (state.timedOutQuestionId) {
      return;
    }

    const quiz = getSelectedQuiz();
    if (!quiz) {
      return;
    }

    const activeQuestions = state.activeQuestions.length ? state.activeQuestions : quiz.questions;
    const question = activeQuestions[state.currentQuestionIndex];
    if (!question) {
      return;
    }

    const selected = questionForm?.querySelector(`input[name="${question.id}"]:checked`);
    if (selected) {
      selected.checked = false;
    }

    delete state.answers[question.id];
    const alternativesVisibilityMode = quiz.alternativesVisibilityMode === "hidden" ? "hidden" : "revealed";
    if (alternativesVisibilityMode === "revealed" && questionForm) {
      questionForm.classList.remove("masked");
    }
    clearValidationMessage();
    updateNextButtonVisibility();
  });
}
if (skipQuestionButton) {
  skipQuestionButton.addEventListener("click", handleSkipQuestion);
}
cancelButton.addEventListener("click", () => {
  openActionModal({
    title: "Bloquear avaliação?",
    text: "Se você bloquear agora, não terá mais acesso a este quiz.",
    confirmLabel: "Sim, bloquear",
    onConfirm: async () => {
      await cancelQuiz("Questionário bloqueado ao bloquear a avaliação.", { blockedByViolation: true });
    }
  });
});

if (cancelModalConfirm) {
  cancelModalConfirm.addEventListener("click", async () => {
    const action = pendingModalAction;
    hideCancelModal();
    if (typeof action === "function") {
      await action();
    }
  });
}
if (cancelModalCancel) {
  cancelModalCancel.addEventListener("click", () => {
    if (cancelModal) cancelModal.classList.add("hidden");
  });
}
policyReturnButton.addEventListener("click", handleReturnToQuiz);
resultBackHomeButton.addEventListener("click", async () => {
  await showHome({ refresh: true });
});
cancelBackHomeButton.addEventListener("click", async () => {
  await showHome({ refresh: true });
});
if (resultReviewButton) {
  resultReviewButton.addEventListener("click", async () => {
    if (state.selectedQuizId) {
      await syncLatestAttemptForQuiz(state.selectedQuizId);
    }
    openReviewScreen("result");
  });
}
if (cancelReviewButton) {
  cancelReviewButton.addEventListener("click", async () => {
    if (state.selectedQuizId) {
      await syncLatestAttemptForQuiz(state.selectedQuizId);
    }
    openReviewScreen("cancel");
  });
}
if (reviewBackButton) {
  reviewBackButton.addEventListener("click", async () => {
    if (state.reviewBackScreen === "teacher") {
      await openTeacherPanel();
      return;
    }

    if (state.reviewBackScreen === "home") {
      await showHome({ refresh: true });
      return;
    }

    if (state.selectedQuizId) {
      await syncLatestAttemptForQuiz(state.selectedQuizId);
    }

    const quiz = getSelectedQuiz();
    const attempt = getAttemptForSelectedQuiz();

    if (state.reviewBackScreen === "cancel") {
      const base = attempt?.cancelReason || "Avaliação bloqueada por violação de regras.";
      showResult(attempt?.result || {}, "", { isBlocked: true, reason: base });
      if (quiz && attempt) {
        updateReviewButtons(quiz, attempt);
      }
      return;
    }

    if (attempt?.result) {
      if (attempt.status === "cancelled") {
        const reason = attempt.cancelReason || "Avaliação bloqueada por violação de regras.";
        showResult(attempt.result, "", { isBlocked: true, reason });
      } else {
        showResult(attempt.result, "");
      }
      if (quiz) {
        updateReviewButtons(quiz, attempt);
      }
      return;
    }

    showOnlyScreen("result");
  });
}
teacherRefreshButton.addEventListener("click", openTeacherPanel);
if (homeHeaderButton) {
  homeHeaderButton.addEventListener("click", async () => {
    if (state.isActive && getSelectedQuiz()) {
      if (state.isTeacher) {
        await showHome({ refresh: true });
        return;
      }
      openActionModal({
        title: "Ir para a home?",
        text: "Ao voltar para a home, este questionário será bloqueado e você perderá o acesso a ele.",
        confirmLabel: "Sim, ir para a home",
        onConfirm: async () => {
          await cancelQuiz("Questionário bloqueado ao voltar para a home durante a avaliação.", { blockedByViolation: true, skipScreen: true });
          await showHome({ refresh: true });
        }
      });
      return;
    }
    await showHome({ refresh: true });
  });
}
if (builderAddQuestionButton) {
  builderAddQuestionButton.addEventListener("click", () => addBuilderQuestionCard());
}
if (builderAiGenerateButton) {
  builderAiGenerateButton.addEventListener("click", handleBuilderGenerateWithAi);
}
if (builderAiApiKeyInput) {
  const savedApiKey = getStoredGeminiApiKey();
  if (savedApiKey) {
    builderAiApiKeyInput.value = savedApiKey;
  }
  builderAiApiKeyInput.addEventListener("change", () => {
    setStoredGeminiApiKey(builderAiApiKeyInput.value || "");
  });
}

profileForm.addEventListener("submit", handleProfileSubmit);

if (profileCancelButton) {
  profileCancelButton.addEventListener("click", () => {
    state.isRenamingProfile = false;
    profileMessage.textContent = "";
    showAuthenticatedUI();
  });
}

if (userChip) {
  userChip.addEventListener("click", (e) => {
    e.stopPropagation();
    if (!userChipPopup) return;
    const email = state.user?.email || "";
    if (userChipPopupEmail) userChipPopupEmail.textContent = email || "E-mail não disponível";
    userChipPopup.classList.toggle("hidden");
  });
}

if (userChipPopupRename) {
  userChipPopupRename.addEventListener("click", () => {
    if (userChipPopup) userChipPopup.classList.add("hidden");

    if (state.isActive && getSelectedQuiz() && !state.isTeacher) {
      openActionModal({
        title: "Alterar nome durante a avaliação?",
        text: "Se você continuar, o questionário atual será bloqueado e você perderá o acesso a ele.",
        confirmLabel: "Sim, alterar nome",
        onConfirm: async () => {
          await cancelQuiz("Questionário bloqueado ao alterar o nome durante a avaliação.", { blockedByViolation: true, skipScreen: true });
          openRenameProfileScreen();
        }
      });
      return;
    }

    openRenameProfileScreen();
  });
}

document.addEventListener("click", () => {
  if (userChipPopup) userChipPopup.classList.add("hidden");
});

if (builderForm) {
  builderForm.addEventListener("submit", handleBuilderCreateQuiz);
}
if (teacherClassroomForm) {
  teacherClassroomForm.addEventListener("submit", handleCreateClassroom);
}

document.querySelectorAll("form").forEach((form) => {
  form.setAttribute("autocomplete", "off");
});
document.querySelectorAll("input, textarea, select").forEach((field) => {
  field.setAttribute("autocomplete", "off");
});

teacherPanelButton.addEventListener("click", async () => {
  if (!state.isTeacher) {
    alert("Acesso negado ao painel do professor.");
    return;
  }

  await openTeacherPanel();
});

if (teacherClassroomsButton) {
  teacherClassroomsButton.addEventListener("click", async () => {
    if (!state.isTeacher) {
      alert("Acesso negado às turmas.");
      return;
    }

    await openTeacherClassroomsScreen();
  });
}

if (teacherBuilderButton) {
  teacherBuilderButton.addEventListener("click", () => {
    if (!state.isTeacher) {
      alert("Acesso negado ao criador de quiz.");
      return;
    }
    openBuilderScreen();
  });
}

if (teacherNonogramEditorButton) {
  teacherNonogramEditorButton.addEventListener("click", () => {
    if (!state.isTeacher) {
      alert("Acesso negado ao editor de nonograma.");
      return;
    }
    openHomeNonogramEditorModal();
  });
}

logoutButton.addEventListener("click", async () => {
  if (state.isActive && getSelectedQuiz()) {
    if (state.isTeacher) {
      if (typeof window.firebaseSignOut === "function") await window.firebaseSignOut();
      return;
    }
    openActionModal({
      title: "Sair da plataforma?",
      text: "Ao sair agora, este questionário será bloqueado e você perderá o acesso a ele.",
      confirmLabel: "Sim, sair",
      onConfirm: async () => {
        await cancelQuiz("Questionário bloqueado ao sair da plataforma durante a avaliação.", { blockedByViolation: true, skipScreen: true });
        if (typeof window.firebaseSignOut === "function") {
          await window.firebaseSignOut();
        }
      }
    });
    return;
  }
  if (typeof window.firebaseSignOut === "function") {
    await window.firebaseSignOut();
  }
});

if (authGoogleBtn) {
  authGoogleBtn.addEventListener("click", async () => {
    authMessage.textContent = "";
    try {
      await window.firebaseSignInWithGoogle();
    } catch (error) {
      // Friendly, actionable messages for common Firebase auth errors
      console.error(error);
      if (error && (error.code === "auth/operation-not-allowed" || /operation-not-allowed/.test(error.message))) {
        authMessage.textContent = "Login com Google desativado: habilite o provedor Google em Firebase Console → Authentication → Sign-in method.";
      } else if (error && (error.code === "auth/popup-blocked" || error.code === "auth/popup-closed-by-user")) {
        authMessage.textContent = "Popup bloqueado ou fechado. Verifique bloqueadores de popup e tente novamente.";
      } else {
        authMessage.textContent = "Não foi possível entrar com Google. Verifique o console para mais detalhes.";
      }
      // Helpful suggestion when running from file:// (popups/auth won't work)
      if (location.protocol === "file:") {
        authMessage.textContent += " Execute um servidor local (ex.: `python -m http.server 8000`) e acesse via http://localhost:8000.";
      }
    }
  });
}

quizSearch.addEventListener("input", renderHome);

document.addEventListener("visibilitychange", () => {
  if (state.isActive && !state.isTeacher && document.visibilityState !== "visible") {
    handlePolicyViolation();
    return;
  }

  if (!state.isTeacher && document.visibilityState !== "visible") {
    triggerFocusWaitGrace("Troca de aba ou minimização da página antes da liberação do professor.");
  }

  if (!state.isTeacher && document.visibilityState === "visible" && isFocusWaitModeActive() && !isFullscreenActive()) {
    triggerFocusWaitGrace("O modo foco exige tela cheia até o professor liberar o modo de espera.");
  }
});

window.addEventListener("blur", () => {
  setTimeout(() => {
    if (state.isActive && !state.isTeacher && !document.hasFocus()) {
      handlePolicyViolation();
      return;
    }

    if (!state.isTeacher && !document.hasFocus()) {
      triggerFocusWaitGrace("Perda de foco da janela antes da liberação do professor.");
    }
  }, 0);
});

document.addEventListener("fullscreenchange", () => {
  if (state.isActive && !state.isTeacher && !isFullscreenActive()) {
    handlePolicyViolation();
    return;
  }

  if (!state.isTeacher && isFocusWaitModeActive() && !isFullscreenActive()) {
    triggerFocusWaitGrace("Você saiu da tela cheia durante o modo foco.");
    return;
  }

  if (isFocusWaitGraceActive() && isFocusWaitRecoverySatisfied()) {
    resolveFocusWaitGraceWithoutPenalty();
  }
});

document.addEventListener("wheel", (e) => {
  if (state.isActive && !state.isTeacher && e.ctrlKey) {
    e.preventDefault();
  }
}, { passive: false });

document.addEventListener("keydown", (e) => {
  if (tryHandleReloadShortcut(e)) {
    return;
  }

  if (state.isActive && !state.isTeacher && e.ctrlKey) {
    if (["=", "+", "-", "_", "0"].includes(e.key)) {
      e.preventDefault();
    }
  }
});

document.addEventListener("contextmenu", (event) => {
  if (state.isTeacher) {
    return;
  }

  event.preventDefault();
  const now = Date.now();
  if (now - lastContextMenuBlockedAt > 1800) {
    showBottomNotice("Clique direito desativado para alunos.", "info", 1800);
    lastContextMenuBlockedAt = now;
  }
});

window.addEventListener("beforeunload", (event) => {
  if (state.isTeacher) {
    return;
  }

  if (state.isActive && getSelectedQuiz()) {
    // Browser-native beforeunload dialogs pause JS execution, freezing timers.
    // Persist cancellation and allow unload without opening the native prompt.
    persistUnloadCancellation();
    return;
  }

  persistFocusWaitPenaltyOnUnload();
});

document.addEventListener("DOMContentLoaded", initializeAuth);

function endBootPhase() {
  document.body.classList.remove("app-booting");
  if (bootScreen) {
    bootScreen.classList.add("hidden");
  }
}

function initializeAuth() {
  if (typeof window.firebaseOnAuthStateChanged !== "function") {
    authMessage.textContent = "Erro ao carregar autenticação do Firebase.";
    return;
  }

  window.firebaseOnAuthStateChanged(async (user) => {
    state.user = user;
    if (user) {
      await ensureUserProfile();
      return;
    }

    showUnauthenticatedUI();
  });
}

function normalizeNameSlug(name) {
  return name
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9 ]/g, "")
    .trim()
    .replace(/\s+/g, "-") || "usuario";
}

function escapeHtml(str) {
  if (str == null) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getBlockedReviewNoteHtml(baseMessage = "") {
  const prefix = baseMessage ? `${escapeHtml(baseMessage)} ` : "";
  return `${prefix}<span class="quiz-review-note">Resumo bloqueado no momento.</span>`;
}

let bottomNoticeTimeoutId = null;
function showBottomNotice(message, type = "info", durationMs = 2800) {
  if (!message) {
    return;
  }

  let notice = document.getElementById("app-bottom-notice");
  if (!notice) {
    notice = document.createElement("div");
    notice.id = "app-bottom-notice";
    notice.className = "app-bottom-notice hidden";
    notice.setAttribute("role", "status");
    notice.setAttribute("aria-live", "polite");
    document.body.appendChild(notice);
  }

  notice.textContent = String(message);
  notice.classList.remove("hidden", "is-info", "is-success", "is-error");
  notice.classList.add(type === "success" ? "is-success" : type === "error" ? "is-error" : "is-info");

  if (bottomNoticeTimeoutId) {
    clearTimeout(bottomNoticeTimeoutId);
  }

  bottomNoticeTimeoutId = setTimeout(() => {
    notice.classList.add("hidden");
  }, Math.max(1200, Number(durationMs) || 2800));
}

function normalizeClassroomId(name) {
  return String(name || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9 ]/g, "")
    .trim()
    .replace(/\s+/g, "-");
}

function renderProfileClassroomOptions(selectedId = "") {
  if (!profileClassroomSelect) return;
  const noClassMsg = document.getElementById("profile-no-classroom-message");
  const saveBtn = document.getElementById("profile-submit-button");
  const classGroup = document.getElementById("profile-classroom-group");

  // Se for professor, esconde grupo de turma e remove required
  if (state.isTeacher) {
    if (classGroup) classGroup.style.display = "none";
    if (profileClassroomSelect) profileClassroomSelect.required = false;
    if (saveBtn) saveBtn.disabled = false;
    return;
  } else {
    if (classGroup) classGroup.style.display = "";
    if (profileClassroomSelect) profileClassroomSelect.required = true;
  }

  const options = ['<option value="">Selecione uma turma</option>'];
  state.classrooms.forEach((classroom) => {
    options.push(`<option value="${classroom.id}">${classroom.nome}</option>`);
  });
  profileClassroomSelect.innerHTML = options.join("");
  profileClassroomSelect.value = selectedId || "";

  if (state.classrooms.length === 0) {
    profileClassroomSelect.disabled = true;
    if (saveBtn) saveBtn.disabled = true;
    if (noClassMsg) {
      noClassMsg.style.display = 'block';
      noClassMsg.textContent = 'Nenhuma turma disponível. Aguarde o professor criar uma turma.';
    }
  } else {
    profileClassroomSelect.disabled = false;
    if (saveBtn) saveBtn.disabled = false;
    if (noClassMsg) noClassMsg.style.display = 'none';
  }
}

function renderTeacherClassrooms() {
  if (!teacherClassroomsList) {
    return;
  }

  if (!state.classrooms.length) {
    teacherClassroomsList.innerHTML = "<p class=\"teacher-empty-state\">Nenhuma turma criada ainda.</p>";
    return;
  }

  teacherClassroomsList.innerHTML = state.classrooms
    .map((classroom) => `
      <article class="teacher-classroom-manage-card">
        <div class="teacher-classroom-manage-main">
          <h3>${classroom.nome}</h3>
          <p>ID: ${classroom.id}</p>
          <div class="teacher-classroom-students teacher-classroom-dropzone" data-classroom-id="${classroom.id}">
            <strong>Alunos</strong>
            ${Array.isArray(state.classroomStudents[classroom.id]) && state.classroomStudents[classroom.id].length
        ? `<div class="teacher-classroom-student-list">${state.classroomStudents[classroom.id]
          .map((student) => `<span class="teacher-classroom-student-chip teacher-student-draggable" draggable="true" data-student-slug="${escapeHtml(student.slug || "")}" data-source-classroom-id="${classroom.id}">${escapeHtml(student.nome || student.email || student.slug || "Aluno sem nome")}</span>`)
          .join("")}</div>`
        : `<p class="teacher-classroom-student-empty">Nenhum aluno vinculado.</p>`}
          </div>
        </div>
        <div class="teacher-classroom-manage-actions">
          <button type="button" class="btn btn-danger-soft teacher-delete-classroom-btn" data-classroom-id="${classroom.id}">Excluir</button>
        </div>
      </article>
    `)
    .join("");

  teacherClassroomsList.querySelectorAll(".teacher-delete-classroom-btn").forEach((button) => {
    button.addEventListener("click", async () => {
      const classroomId = button.dataset.classroomId;
      if (!classroomId) {
        return;
      }

      await deleteClassroom(classroomId);
    });
  });

  bindTeacherClassroomDnD();
}

function bindTeacherClassroomDnD() {
  if (!teacherClassroomsList || !state.isTeacher) {
    return;
  }

  teacherClassroomsList.querySelectorAll(".teacher-student-draggable").forEach((chip) => {
    chip.addEventListener("dragstart", (event) => {
      const studentSlug = String(chip.dataset.studentSlug || "");
      const sourceClassroomId = String(chip.dataset.sourceClassroomId || "");
      if (!studentSlug || !sourceClassroomId || !event.dataTransfer) {
        event.preventDefault();
        return;
      }

      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", JSON.stringify({ studentSlug, sourceClassroomId }));
      chip.classList.add("is-dragging");
    });

    chip.addEventListener("dragend", () => {
      chip.classList.remove("is-dragging");
      teacherClassroomsList.querySelectorAll(".teacher-classroom-dropzone").forEach((zone) => {
        zone.classList.remove("is-drag-over");
      });
    });
  });

  teacherClassroomsList.querySelectorAll(".teacher-classroom-dropzone").forEach((zone) => {
    zone.addEventListener("dragover", (event) => {
      event.preventDefault();
      if (event.dataTransfer) {
        event.dataTransfer.dropEffect = "move";
      }
      zone.classList.add("is-drag-over");
    });

    zone.addEventListener("dragleave", () => {
      zone.classList.remove("is-drag-over");
    });

    zone.addEventListener("drop", async (event) => {
      event.preventDefault();
      zone.classList.remove("is-drag-over");

      const targetClassroomId = String(zone.dataset.classroomId || "");
      const payload = event.dataTransfer?.getData("text/plain") || "";
      if (!payload || !targetClassroomId) {
        return;
      }

      try {
        const parsed = JSON.parse(payload);
        await moveStudentToClassroom(String(parsed.studentSlug || ""), String(parsed.sourceClassroomId || ""), targetClassroomId);
      } catch (error) {
        console.error("Erro ao processar arrastar/soltar de aluno:", error);
      }
    });
  });
}

async function moveStudentToClassroom(studentSlug, sourceClassroomId, targetClassroomId) {
  if (!studentSlug || !sourceClassroomId || !targetClassroomId || sourceClassroomId === targetClassroomId) {
    return;
  }

  const sourceStudents = Array.isArray(state.classroomStudents[sourceClassroomId])
    ? [...state.classroomStudents[sourceClassroomId]]
    : [];
  const targetStudents = Array.isArray(state.classroomStudents[targetClassroomId])
    ? [...state.classroomStudents[targetClassroomId]]
    : [];

  const studentIndex = sourceStudents.findIndex((student) => student.slug === studentSlug);
  if (studentIndex < 0) {
    return;
  }

  const student = sourceStudents[studentIndex];
  const targetClassroom = state.classrooms.find((item) => item.id === targetClassroomId);
  if (!targetClassroom) {
    return;
  }

  const previousState = JSON.parse(JSON.stringify(state.classroomStudents || {}));
  sourceStudents.splice(studentIndex, 1);
  targetStudents.push(student);
  targetStudents.sort((a, b) => {
    const aName = a.nome || a.email || a.slug;
    const bName = b.nome || b.email || b.slug;
    return aName.localeCompare(bName, "pt-BR");
  });

  state.classroomStudents[sourceClassroomId] = sourceStudents;
  state.classroomStudents[targetClassroomId] = targetStudents;
  renderTeacherClassrooms();

  if (!window.firebaseDB || !window.firebaseDoc || !window.firebaseSetDoc) {
    return;
  }

  try {
    const userRef = window.firebaseDoc(window.firebaseDB, "usuarios", student.slug);
    await window.firebaseSetDoc(userRef, {
      turmaId: targetClassroomId,
      turmaNome: targetClassroom.nome,
      atualizadoEmIso: new Date().toISOString(),
      atualizadoPorUid: state.user?.uid || "",
      atualizadoPorEmail: state.user?.email || ""
    }, { merge: true });

    if (student.uid) {
      const userIndexRef = window.firebaseDoc(window.firebaseDB, "usuarios_index", student.uid);
      await window.firebaseSetDoc(userIndexRef, {
        turmaId: targetClassroomId,
        turmaNome: targetClassroom.nome,
        atualizadoEmIso: new Date().toISOString()
      }, { merge: true });
    }

    if (teacherClassroomsMessage) {
      const displayName = student.nome || student.email || student.slug;
      teacherClassroomsMessage.textContent = `${displayName} movido(a) para ${targetClassroom.nome}.`;
    }

    // Se o painel de Resultados estiver aberto, recarregar seus dados
    if (!teacherScreen.classList.contains("hidden")) {
      await loadTeacherPanelData();
    }
  } catch (error) {
    state.classroomStudents = previousState;
    renderTeacherClassrooms();
    if (teacherClassroomsMessage) {
      teacherClassroomsMessage.textContent = "Não foi possível mover o aluno para outra turma.";
    }
    console.error("Erro ao mover aluno de turma:", error);
  }
}

async function refreshClassrooms() {
  if (!window.firebaseDB || !window.firebaseCollection || !window.firebaseGetDocs) {
    state.classrooms = [];
    renderProfileClassroomOptions();
    renderTeacherClassrooms();
    return;
  }

  try {
    const classesRef = window.firebaseCollection(window.firebaseDB, "turmas");
    const snap = await window.firebaseGetDocs(classesRef);
    const next = [];

    snap.forEach((docSnap) => {
      const data = docSnap.data() || {};
      if (data.ativo === false) {
        return;
      }
      const nome = String(data.nome || "").trim();
      if (!nome) {
        return;
      }
      next.push({ id: docSnap.id, nome });
    });

    next.sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR"));
    state.classrooms = next;
    renderProfileClassroomOptions(state.classroomId);
    renderTeacherClassrooms();
  } catch (error) {
    console.error("Erro ao carregar turmas:", error);
  }
}

async function handleCreateClassroom(event) {
  event.preventDefault();
  if (!state.isTeacher) {
    alert("Apenas professores podem criar turmas.");
    return;
  }

  const classroomName = String(teacherClassroomNameInput?.value || "").trim();
  if (!classroomName) {
    return;
  }

  const classroomId = normalizeClassroomId(classroomName);
  if (!classroomId) {
    if (teacherClassroomsMessage) teacherClassroomsMessage.textContent = "Nome de turma inválido.";
    return;
  }

  if (!window.firebaseDB || !window.firebaseDoc || !window.firebaseSetDoc) {
    if (teacherClassroomsMessage) teacherClassroomsMessage.textContent = "Firebase indisponível para criar turma.";
    return;
  }

  try {
    const classRef = window.firebaseDoc(window.firebaseDB, "turmas", classroomId);
    await window.firebaseSetDoc(classRef, {
      nome: classroomName,
      ativo: true,
      criadoPorUid: state.user?.uid || "",
      criadoPorEmail: state.user?.email || "",
      atualizadoEmIso: new Date().toISOString()
    }, { merge: true });

    if (teacherClassroomForm) {
      teacherClassroomForm.reset();
    }
    if (teacherClassroomsMessage) teacherClassroomsMessage.textContent = "Turma criada com sucesso.";
    await refreshClassrooms();
    await loadTeacherClassroomData();
  } catch (error) {
    if (teacherClassroomsMessage) teacherClassroomsMessage.textContent = "Erro ao criar turma.";
    console.error("Erro ao criar turma:", error);
  }
}

async function deleteClassroom(classroomId) {
  if (!state.isTeacher || !window.firebaseDB || !window.firebaseDoc || !window.firebaseSetDoc) {
    return;
  }

  const classroom = state.classrooms.find((item) => item.id === classroomId);
  if (!classroom) {
    return;
  }

  const ok = confirm(`Excluir a turma '${classroom.nome}'? Ela deixará de aparecer para novos alunos.`);
  if (!ok) {
    return;
  }

  try {
    const classRef = window.firebaseDoc(window.firebaseDB, "turmas", classroomId);
    await window.firebaseSetDoc(classRef, {
      ativo: false,
      atualizadoEmIso: new Date().toISOString(),
      atualizadoPorUid: state.user?.uid || "",
      atualizadoPorEmail: state.user?.email || ""
    }, { merge: true });

    if (teacherClassroomsMessage) {
      teacherClassroomsMessage.textContent = "Turma excluída com sucesso.";
    }
    await refreshClassrooms();
    await loadTeacherClassroomData();
  } catch (error) {
    if (teacherClassroomsMessage) {
      teacherClassroomsMessage.textContent = "Erro ao excluir turma.";
    }
    console.error("Erro ao excluir turma:", error);
  }
}

async function loadTeacherClassroomData() {
  if (!teacherClassroomsList) {
    return;
  }

  if (teacherClassroomsMessage && !teacherClassroomsMessage.textContent) {
    teacherClassroomsMessage.textContent = "";
  }

  await refreshClassrooms();

  state.classroomStudents = {};
  if (!window.firebaseDB || !window.firebaseCollection || !window.firebaseGetDocs) {
    renderTeacherClassrooms();
    return;
  }

  try {
    const usersRef = window.firebaseCollection(window.firebaseDB, "usuarios");
    const usersSnap = await window.firebaseGetDocs(usersRef);
    const nextStudents = {};

    usersSnap.forEach((docSnap) => {
      const data = docSnap.data() || {};
      const classroomId = String(data.turmaId || "").trim();
      if (!classroomId) {
        return;
      }

      if (!nextStudents[classroomId]) {
        nextStudents[classroomId] = [];
      }

      nextStudents[classroomId].push({
        slug: docSnap.id,
        uid: String(data.ultimoUid || "").trim(),
        nome: String(data.nome || "").trim(),
        email: String(data.email || "").trim()
      });
    });

    Object.keys(nextStudents).forEach((classroomId) => {
      nextStudents[classroomId].sort((a, b) => {
        const aName = a.nome || a.email || a.slug;
        const bName = b.nome || b.email || b.slug;
        return aName.localeCompare(bName, "pt-BR");
      });
    });

    state.classroomStudents = nextStudents;
  } catch (error) {
    console.error("Erro ao carregar alunos por turma:", error);
    if (teacherClassroomsMessage && !teacherClassroomsMessage.textContent) {
      teacherClassroomsMessage.textContent = "Não foi possível carregar a lista de alunos das turmas.";
    }
  }

  renderTeacherClassrooms();
}

function savePendingUnloadCancellation(payload) {
  localStorage.setItem(getPendingUnloadKey(), JSON.stringify(payload));
}

function savePendingFocusWaitPenalty(payload) {
  localStorage.setItem(getPendingFocusWaitPenaltyKey(), JSON.stringify(payload));
}

function readPendingUnloadCancellation() {
  const key = getPendingUnloadKey();
  const raw = localStorage.getItem(key);
  if (!raw) {
    return null;
  }

  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function readPendingFocusWaitPenalty() {
  const key = getPendingFocusWaitPenaltyKey();
  const raw = localStorage.getItem(key);
  if (!raw) {
    return null;
  }

  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function clearPendingUnloadCancellation() {
  localStorage.removeItem(getPendingUnloadKey());
}

function clearPendingFocusWaitPenalty() {
  localStorage.removeItem(getPendingFocusWaitPenaltyKey());
}

function consumePendingUnloadCancellation() {
  const pending = readPendingUnloadCancellation();
  if (pending) {
    clearPendingUnloadCancellation();
  }
  return pending;
}

function persistUnloadCancellation() {
  const quiz = getSelectedQuiz();
  if (!state.isActive || state.isTeacher || !quiz) {
    return;
  }

  const cancelAt = new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" });
  const reason = "Questionário bloqueado ao recarregar ou fechar a página.";
  const answersSnapshot = cloneAnswersSnapshot();
  const cancelled = buildCancelledAttemptResult(quiz, answersSnapshot, cancelAt);
  const result = cancelled.result;
  const questionResults = cancelled.questionResults;

  // Salva localmente
  saveStoredAttempt(quiz.id, {
    status: "cancelled",
    cancelReason: reason,
    blockedByViolation: true,
    at: cancelAt,
    result,
    answers: answersSnapshot,
    questionResults
  });
  lockQuizAccess(quiz.id);
  // Salva no Firestore para o painel do professor
  saveCancelledAttemptToFirestore(quiz, {
    reason,
    blockedByViolation: true,
    at: cancelAt,
    result,
    answers: answersSnapshot,
    questionResults
  }).catch((error) => {
    console.error("Erro ao salvar cancelamento no descarregamento da página:", error);
  });

  savePendingUnloadCancellation({
    quizId: quiz.id,
    quizTitle: quiz.title,
    questionsLength: quiz.questions.length,
    reason,
    blockedByViolation: true,
    at: cancelAt,
    result,
    answers: answersSnapshot,
    questionResults
  });
}

async function flushPendingUnloadCancellation() {
  const pending = readPendingUnloadCancellation();
  if (!pending) {
    return;
  }

  const quiz = getAllQuizzes().find((item) => item.id === pending.quizId);
  const quizLike = quiz || {
    id: pending.quizId,
    title: pending.quizTitle || pending.quizId,
    questions: Array.from({ length: pending.questionsLength || 0 }, (_, index) => ({ id: `q${index + 1}` }))
  };
  const rebuilt = buildCancelledAttemptResult(
    quizLike,
    pending.answers || {},
    pending?.result?.date || pending.at || new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" })
  );
  const normalizedResult = quiz ? rebuilt.result : (pending.result || rebuilt.result);
  const normalizedQuestionResults = Array.isArray(pending.questionResults) && pending.questionResults.length
    ? pending.questionResults
    : rebuilt.questionResults;

  try {
    await saveCancelledAttemptToFirestore(quizLike, {
      reason: pending.reason,
      blockedByViolation: pending.blockedByViolation,
      at: pending.at,
      result: normalizedResult,
      answers: pending.answers || {},
      questionResults: normalizedQuestionResults
    });
    clearPendingUnloadCancellation();
  } catch (error) {
    console.error("Erro ao reenviar cancelamento pendente:", error);
  }
}

async function refreshQuizConfigs(options = {}) {
  if (!window.firebaseDB || !window.firebaseCollection || !window.firebaseGetDocs) {
    state.quizConfigs = {};
    return;
  }

  try {
    const configRef = window.firebaseCollection(window.firebaseDB, "quizzes_config");
    const snap = await window.firebaseGetDocs(configRef);
    const nextMap = {};

    snap.forEach((docSnap) => {
      const data = docSnap.data() || {};
      nextMap[docSnap.id] = {
        allowStudentReview: Boolean(data.allowStudentReview),
        allowStudentStart: Boolean(data.allowStudentStart),
        allowFocusExitAfterFinish: Boolean(data.allowFocusExitAfterFinish),
        hidden: Boolean(data.hidden)
      };
    });

    state.quizConfigs = nextMap;
    if (options.render !== false) {
      renderHome();
    }
  } catch (error) {
    console.error("Erro ao carregar configuração de quizzes:", error);
  }
}

function isRetakeReleased(quizId) {
  return Boolean(state.retakeReleases?.[quizId]);
}

async function refreshRetakeReleases(options = {}) {
  if (!window.firebaseDB || !window.firebaseDoc || !window.firebaseGetDoc || !state.profileSlug) {
    state.retakeReleases = {};
    return;
  }

  try {
    const userRef = window.firebaseDoc(window.firebaseDB, "usuarios", state.profileSlug);
    const userSnap = await window.firebaseGetDoc(userRef);
    if (!userSnap.exists()) {
      state.retakeReleases = {};
      return;
    }

    const data = userSnap.data() || {};
    state.retakeReleases = data.liberacoes || {};
    // Limpa tentativas locais se houver liberação para refazer
    Object.keys(state.retakeReleases).forEach((quizId) => {
      if (state.retakeReleases[quizId]) {
        localStorage.removeItem(getStorageKey(quizId));
        if (state.remoteAttemptsByQuiz) delete state.remoteAttemptsByQuiz[quizId];
      }
    });
    if (options.render !== false) {
      renderHome();
    }
  } catch (error) {
    console.error("Erro ao carregar liberações de refazer:", error);
    state.retakeReleases = {};
  }
}

async function setRetakeReleaseForStudent(slug, quizId, liberar) {
  if (!window.firebaseDB || !window.firebaseDoc || !window.firebaseSetDoc) {
    throw new Error("Firebase indisponível para liberar nova tentativa.");
  }

  const userRef = window.firebaseDoc(window.firebaseDB, "usuarios", slug);
  await window.firebaseSetDoc(userRef, {
    liberacoes: { [quizId]: Boolean(liberar) },
    atualizadoEmIso: new Date().toISOString()
  }, { merge: true });
}

async function consumeRetakeRelease(quizId) {
  if (!isRetakeReleased(quizId) || !state.profileSlug) {
    return;
  }

  state.retakeReleases[quizId] = false;
  try {
    await setRetakeReleaseForStudent(state.profileSlug, quizId, false);
  } catch (error) {
    console.error("Erro ao consumir liberação de refazer:", error);
  }
}

async function gradeQuizAttemptSecure(quiz, answersSnapshot, at) {
  if (!window.firebaseHttpsCallable || !window.firebaseFunctions) {
    const error = new Error("Firebase Functions não está disponível.");
    error.code = "functions/unavailable";
    throw error;
  }

  const callable = window.firebaseHttpsCallable(window.firebaseFunctions, "gradeQuizAttempt");
  let response;
  try {
    response = await callable({
      quizId: quiz.id,
      studentName: state.studentName,
      profileSlug: state.profileSlug,
      date: at,
      answers: answersSnapshot || {}
    });
  } catch (rawError) {
    const error = rawError || new Error("Falha ao chamar função de correção.");
    error.code = rawError?.code || rawError?.details?.code || "functions/unknown";
    throw error;
  }

  const data = response?.data || {};
  if (!data || !data.result) {
    throw new Error("Resposta inválida da correção segura.");
  }

  return {
    result: data.result,
    questionResults: Array.isArray(data.questionResults) ? data.questionResults : []
  };
}

function getSecureGradingErrorMessage(error) {
  const code = String(error?.code || "").toLowerCase();
  if (code.includes("unauthenticated")) {
    return "Sua sessão expirou. Entre novamente com Google e tente de novo.";
  }
  if (code.includes("not-found")) {
    return "Este quiz não foi encontrado no servidor. Atualize a página e tente novamente.";
  }
  if (code.includes("failed-precondition")) {
    return "Este quiz está sem gabarito no servidor. Peça ao professor para revisar o cadastro.";
  }
  if (code.includes("permission-denied")) {
    return "Sem permissão para concluir esta operação no momento.";
  }
  if (code.includes("unavailable") || code.includes("deadline-exceeded") || code.includes("internal")) {
    return "Serviço de correção indisponível agora. Se persistir, verifique se Cloud Functions está ativa no projeto Firebase.";
  }
  return "Não foi possível corrigir e salvar seu resultado agora. Verifique sua conexão e tente novamente em instantes.";
}

function getLocalGradingErrorMessage(error) {
  const code = String(error?.code || "").toLowerCase();
  if (code.includes("local-grading/missing-key")) {
    return "Este quiz ainda não foi atualizado para correção local no plano Spark. Peça ao professor para abrir e salvar o quiz novamente no Criador de quiz.";
  }

  return "Não foi possível corrigir localmente este quiz no momento.";
}

async function getQuizAnswerKeySecure(quizId) {
  if (!window.firebaseHttpsCallable || !window.firebaseFunctions) {
    return {};
  }

  try {
    const callable = window.firebaseHttpsCallable(window.firebaseFunctions, "getQuizAnswerKey");
    const response = await callable({ quizId });
    const data = response?.data || {};
    return data.answerKey && typeof data.answerKey === "object" ? data.answerKey : {};
  } catch (error) {
    console.error("Não foi possível carregar o gabarito privado para edição:", error);
    return {};
  }
}

function getAllQuizzes() {
  const all = [...state.customQuizzes];

  return all
    .map((quiz) => {
      const cfg = state.quizConfigs[quiz.id] || {};
      return {
        ...quiz,
        allowStudentReview: cfg.allowStudentReview ?? Boolean(quiz.allowStudentReview),
        allowStudentStart: cfg.allowStudentStart ?? Boolean(quiz.allowStudentStart),
        allowFocusExitAfterFinish: cfg.allowFocusExitAfterFinish ?? Boolean(quiz.allowFocusExitAfterFinish),
        hidden: cfg.hidden ?? false
      };
    })
    .filter((quiz) => !quiz.hidden);
}

function shuffleArray(list) {
  const copy = [...list];
  for (let i = copy.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

function buildShuffledQuestionSet(quiz) {
  const shuffledQuestions = quiz.questions.map((question, index) => ({
    ...question,
    id: question.id || `q${index + 1}`,
    options: shuffleArray((question.options || []).map((option) => ({ ...option })))
  }));

  return shuffleArray(shuffledQuestions);
}

function buildAnswerToken(quizId, questionId, answerValue) {
  const cleanQuizId = String(quizId || "quiz");
  const cleanQuestionId = String(questionId || "q1");
  const cleanAnswerValue = String(answerValue || "");
  if (!cleanAnswerValue) {
    return "";
  }

  const plain = `${cleanAnswerValue}::${cleanQuestionId}::${cleanQuizId}`;
  const key = `${cleanQuizId}|${cleanQuestionId}|obf_v1`;
  const bytes = [];
  for (let i = 0; i < plain.length; i += 1) {
    bytes.push(plain.charCodeAt(i) ^ key.charCodeAt(i % key.length));
  }
  return btoa(String.fromCharCode(...bytes));
}

function decodeAnswerToken(quizId, questionId, token) {
  const cleanToken = String(token || "");
  if (!cleanToken) {
    return "";
  }

  try {
    const cleanQuizId = String(quizId || "quiz");
    const cleanQuestionId = String(questionId || "q1");
    const key = `${cleanQuizId}|${cleanQuestionId}|obf_v1`;
    const encoded = atob(cleanToken);
    const chars = [];
    for (let i = 0; i < encoded.length; i += 1) {
      chars.push(String.fromCharCode(encoded.charCodeAt(i) ^ key.charCodeAt(i % key.length)));
    }

    const [answerValue, decodedQuestionId, decodedQuizId] = chars.join("").split("::");
    if (decodedQuestionId !== cleanQuestionId || decodedQuizId !== cleanQuizId) {
      return "";
    }

    return String(answerValue || "");
  } catch (error) {
    console.error("Falha ao decodificar token de resposta:", error);
    return "";
  }
}

function resolveQuestionCorrectValue(quizId, question, index) {
  const questionId = String(question?.id || `q${index + 1}`);
  const byToken = decodeAnswerToken(quizId, questionId, question?.answerToken);
  if (byToken) {
    return byToken;
  }

  return String(question?.correctAnswer || "");
}

function gradeQuizAttemptObfuscated(quiz, answersSnapshot, at) {
  const activeQuestions = state.activeQuestions.length ? state.activeQuestions : quiz.questions;
  let earnedPoints = 0;
  let maxPoints = 0;
  let foundAnyKey = false;

  const questionResults = activeQuestions.map((question, index) => {
    const questionId = String(question?.id || `q${index + 1}`);
    const selectedValue = String(answersSnapshot?.[questionId] || "");
    const correctValue = resolveQuestionCorrectValue(quiz.id, question, index);
    if (correctValue) {
      foundAnyKey = true;
    }

    const points = Number(question?.points || 1);
    const normalizedPoints = Number.isFinite(points) && points > 0 ? points : 1;
    maxPoints += normalizedPoints;

    const options = Array.isArray(question?.options) ? question.options : [];
    const selectedOption = options.find((option) => String(option.value) === selectedValue);
    const correctOption = options.find((option) => String(option.value) === correctValue);

    const isCorrect = Boolean(selectedValue) && Boolean(correctValue) && selectedValue === correctValue;
    if (isCorrect) {
      earnedPoints += normalizedPoints;
    }

    return {
      questionId,
      questionTitle: question?.title || `Questão ${index + 1}`,
      selectedValue,
      selectedLabel: selectedOption ? String(selectedOption.label || "") : "",
      correctValue,
      correctLabel: correctOption ? String(correctOption.label || "") : "",
      isCorrect,
      // include options for review rendering
      questionExplanation: String(question?.explanation || ""),
      options: options.map((opt) => ({
        value: String(opt.value),
        label: String(opt.label || "")
      }))
    };
  });

  if (!foundAnyKey) {
    const err = new Error("Quiz sem token de resposta para correção local.");
    err.code = "local-grading/missing-key";
    throw err;
  }

  const percent = maxPoints > 0 ? Math.round((earnedPoints / maxPoints) * 100) : 0;
  return {
    result: {
      quizId: quiz.id,
      quizTitle: quiz.title,
      studentName: state.studentName,
      earnedPoints,
      maxPoints,
      percent,
      date: at,
      questionResults
    },
    questionResults
  };
}

function normalizeCustomQuiz(rawQuiz, docId) {
  if (!rawQuiz || !Array.isArray(rawQuiz.questions)) {
    return null;
  }

  const normalizedQuestions = rawQuiz.questions
    .map((question, index) => {
      const options = Array.isArray(question.options)
        ? question.options
          .map((option, optionIndex) => {
            if (typeof option === "string") {
              return { value: String.fromCharCode(97 + optionIndex), label: option };
            }
            if (!option || typeof option.label !== "string") {
              return null;
            }
            return {
              value: String(option.value || String.fromCharCode(97 + optionIndex)),
              label: option.label
            };
          })
          .filter(Boolean)
        : [];

      if (!question?.title || options.length < 2) {
        return null;
      }

      return {
        id: question.id || `q${index + 1}`,
        type: "single-choice",
        title: question.title,
        description: question.description || "Selecione apenas uma alternativa.",
        explanation: String(question.explanation || question.questionExplanation || question.justification || question.justificativa || ""),
        points: Number(question.points || 1),
        timer: Number(question.timer || 0),
        options,
        image: question.image || null,
        answerToken: String(question.answerToken || "")
      };
    })
    .filter(Boolean);

  if (!rawQuiz.title || normalizedQuestions.length === 0) {
    return null;
  }

  return {
    id: docId,
    title: rawQuiz.title,
    description: rawQuiz.description || "",
    duration: rawQuiz.duration || "10 min",
    maxAttempts: Math.max(1, Number.parseInt(rawQuiz.maxAttempts, 10) || 1),
    alternativesVisibilityMode: rawQuiz.alternativesVisibilityMode === "hidden" ? "hidden" : "revealed",
    allowStudentReview: Boolean(rawQuiz.allowStudentReview),
    allowStudentStart: rawQuiz.allowStudentStart !== false,
    allowFocusExitAfterFinish: Boolean(rawQuiz.allowFocusExitAfterFinish),
    aiGenerationPrompt: String(rawQuiz.aiGenerationPrompt || rawQuiz.aiPrompt || rawQuiz.generationPrompt || "").trim(),
    aiGenerationProvider: String(rawQuiz.aiGenerationProvider || "").trim(),
    questions: normalizedQuestions,
    isCustom: true
  };
}

function normalizeStoredQuizQuestion(question = {}, index = 0) {
  const options = Array.isArray(question.options)
    ? question.options
      .map((option, optionIndex) => {
        if (typeof option === "string") {
          return { value: String.fromCharCode(97 + optionIndex), label: option };
        }
        if (!option || typeof option.label !== "string") {
          return null;
        }
        return {
          value: String(option.value || String.fromCharCode(97 + optionIndex)),
          label: String(option.label || "").trim()
        };
      })
      .filter((option) => option && option.label)
    : [];

  const correctAnswer = String(question.correctAnswer || question.correctOptionValue || question.answerValue || "").trim();
  const correctAnswerIndex = Number.parseInt(question.correctAnswerIndex, 10);
  let resolvedCorrectAnswer = correctAnswer;

  if (!resolvedCorrectAnswer && Number.isFinite(correctAnswerIndex) && options[correctAnswerIndex]) {
    resolvedCorrectAnswer = options[correctAnswerIndex].value;
  }

  if (!resolvedCorrectAnswer && options[0]) {
    resolvedCorrectAnswer = options[0].value;
  }

  return {
    id: String(question.id || `q${index + 1}`),
    type: String(question.type || "single-choice"),
    title: String(question.title || `Questão ${index + 1}`),
    description: String(question.description || "Selecione apenas uma alternativa."),
    explanation: String(question.explanation || question.questionExplanation || question.justification || question.justificativa || ""),
    points: Math.max(1, Number.parseInt(question.points, 10) || 1),
    timer: Math.max(0, Number.parseInt(question.timer, 10) || 0),
    options,
    image: typeof question.image === "string" ? question.image : null,
    answerToken: String(question.answerToken || ""),
    correctAnswer: resolvedCorrectAnswer
  };
}

async function loadQuizQuestionsFromFirestore(quizId, fallbackQuestions = [], maxCount = null, revision = "") {
  if (!quizId || !window.firebaseDB || !window.firebaseCollection || !window.firebaseGetDocs) {
    return Array.isArray(fallbackQuestions) ? [...fallbackQuestions] : [];
  }

  try {
    const questionsRef = window.firebaseCollection(window.firebaseDB, "quizzes_custom", quizId, "questions");
    const snap = await window.firebaseGetDocs(questionsRef);
    const loadedQuestions = [];

    snap.forEach((docSnap) => {
      const data = docSnap.data() || {};
      if (revision && String(data.revision || "") !== String(revision)) {
        return;
      }
      loadedQuestions.push({
        id: docSnap.id,
        ...data,
        order: Number.isFinite(Number(data.order)) ? Number(data.order) : Number.parseInt(String(docSnap.id || "").replace(/\D+/g, ""), 10) || loadedQuestions.length + 1
      });
    });

    loadedQuestions.sort((a, b) => (Number(a.order || 0) - Number(b.order || 0)) || String(a.id || "").localeCompare(String(b.id || ""), "pt-BR"));
    const limitedQuestions = Number.isFinite(Number(maxCount)) && Number(maxCount) > 0
      ? loadedQuestions.slice(0, Number(maxCount))
      : loadedQuestions;

    const normalized = limitedQuestions
      .map((question, index) => normalizeStoredQuizQuestion(question, index))
      .filter((question) => question.title && Array.isArray(question.options) && question.options.length >= 2);

    return normalized.length ? normalized : (Array.isArray(fallbackQuestions) ? [...fallbackQuestions] : []);
  } catch (error) {
    console.error(`Erro ao carregar questões do quiz ${quizId}:`, error);
    return Array.isArray(fallbackQuestions) ? [...fallbackQuestions] : [];
  }
}

async function loadCustomQuizFromFirestore(quizId, fallbackData = null) {
  if (!quizId || !window.firebaseDB || !window.firebaseDoc || !window.firebaseGetDoc) {
    return fallbackData ? normalizeCustomQuiz(fallbackData, quizId) : null;
  }

  try {
    const quizRef = window.firebaseDoc(window.firebaseDB, "quizzes_custom", quizId);
    const snap = await window.firebaseGetDoc(quizRef);
    if (!snap.exists()) {
      return fallbackData ? normalizeCustomQuiz(fallbackData, quizId) : null;
    }

    const data = snap.data() || {};
    const expectedCount = Number.parseInt(String(data.questionCount || fallbackData?.questionCount || 0), 10) || 0;
    const revision = String(data.questionsRevision || fallbackData?.questionsRevision || "").trim();
    const questions = await loadQuizQuestionsFromFirestore(quizId, Array.isArray(data.questions) ? data.questions : (Array.isArray(fallbackData?.questions) ? fallbackData.questions : []), expectedCount, revision);
    const merged = {
      ...fallbackData,
      ...data,
      questionsRevision: revision,
      questions
    };

    return normalizeCustomQuiz(merged, quizId);
  } catch (error) {
    console.error(`Erro ao carregar quiz customizado ${quizId}:`, error);
    return fallbackData ? normalizeCustomQuiz(fallbackData, quizId) : null;
  }
}

async function replaceQuizQuestionsInFirestore(quizId, questions, revision) {
  if (!quizId || !window.firebaseDB || !window.firebaseCollection || !window.firebaseSetDoc || !window.firebaseDoc) {
    return;
  }

  const questionsRef = window.firebaseCollection(window.firebaseDB, "quizzes_custom", quizId, "questions");
  const writes = (Array.isArray(questions) ? questions : []).map((question, index) => {
    const questionId = String(question.id || `q${index + 1}`);
    const questionRef = window.firebaseDoc(window.firebaseDB, "quizzes_custom", quizId, "questions", questionId);
    const payload = {
      ...question,
      id: questionId,
      order: index + 1,
      revision: String(revision || ""),
      atualizadoEmIso: new Date().toISOString()
    };
    return window.firebaseSetDoc(questionRef, payload);
  });

  await Promise.all(writes);
}

function buildPostedQuizFromPost(rawPost, docId, sourceQuiz = null) {
  if (!rawPost || rawPost.ativo === false) {
    return null;
  }

  const mergedQuestions = Array.isArray(sourceQuiz?.questions) && sourceQuiz.questions.length
    ? sourceQuiz.questions
    : (Array.isArray(rawPost.questions) ? rawPost.questions : []);

  const normalized = normalizeCustomQuiz({
    title: String(sourceQuiz?.title || rawPost.title || "").trim(),
    description: String(sourceQuiz?.description || rawPost.description || "").trim(),
    duration: sourceQuiz?.duration || rawPost.duration,
    maxAttempts: sourceQuiz?.maxAttempts || rawPost.maxAttempts,
    alternativesVisibilityMode: sourceQuiz?.alternativesVisibilityMode || rawPost.alternativesVisibilityMode,
    questions: mergedQuestions,
    allowStudentReview: sourceQuiz?.allowStudentReview ?? rawPost.allowStudentReview,
    allowStudentStart: sourceQuiz?.allowStudentStart ?? rawPost.allowStudentStart,
    allowFocusExitAfterFinish: sourceQuiz?.allowFocusExitAfterFinish ?? rawPost.allowFocusExitAfterFinish,
    aiGenerationPrompt: sourceQuiz?.aiGenerationPrompt || rawPost.aiGenerationPrompt || "",
    aiGenerationProvider: sourceQuiz?.aiGenerationProvider || rawPost.aiGenerationProvider || ""
  }, String(rawPost.quizId || sourceQuiz?.id || docId));

  if (!normalized) {
    return null;
  }

  return {
    ...normalized,
    postId: docId,
    sourceQuizId: String(rawPost.quizId || normalized.id),
    postedToClassroomId: String(rawPost.turmaId || ""),
    postedToClassroomName: String(rawPost.turmaNome || ""),
    atualizadoEmIso: String(rawPost.atualizadoEmIso || rawPost.publicadoEmIso || "")
  };
}

function getPostedClassroomsByQuizId() {
  const map = {};
  (state.quizPosts || []).forEach((post) => {
    const postId = String(post.id || "").trim();
    const quizId = String(post.quizId || "").trim();
    const classroomId = String(post.turmaId || "").trim();
    if (!quizId || !classroomId || !postId) {
      return;
    }

    if (!map[quizId]) {
      map[quizId] = [];
    }

    if (!map[quizId].some((item) => item.id === classroomId)) {
      map[quizId].push({
        postId,
        id: classroomId,
        nome: String(post.turmaNome || classroomId)
      });
    }
  });

  Object.keys(map).forEach((quizId) => {
    map[quizId].sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR"));
  });

  return map;
}

async function refreshQuizPosts(options = {}) {
  if (!window.firebaseDB || !window.firebaseCollection || !window.firebaseGetDocs) {
    state.quizPosts = [];
    return;
  }

  try {
    const postsRef = window.firebaseCollection(window.firebaseDB, "quizzes_posts");
    let snap;

    if (state.isTeacher) {
      snap = await window.firebaseGetDocs(postsRef);
    } else if (window.firebaseQuery && window.firebaseWhere && state.classroomId) {
      const postsQuery = window.firebaseQuery(
        postsRef,
        window.firebaseWhere("turmaId", "==", state.classroomId),
        window.firebaseWhere("ativo", "==", true)
      );
      snap = await window.firebaseGetDocs(postsQuery);
    } else {
      state.quizPosts = [];
      if (options.render !== false) {
        renderHome();
      }
      return;
    }

    const nextPosts = [];
    snap.forEach((docSnap) => {
      const data = docSnap.data() || {};
      if (data.ativo === false) {
        return;
      }

      nextPosts.push({
        id: docSnap.id,
        ...data,
        quizId: String(data.quizId || "").trim(),
        turmaId: String(data.turmaId || "").trim(),
        turmaNome: String(data.turmaNome || "").trim()
      });
    });

    state.quizPosts = nextPosts;
    if (options.render !== false) {
      renderHome();
    }
  } catch (error) {
    console.error("Erro ao carregar postagens de quizzes:", error);
    const code = String(error?.code || "").toLowerCase();
    if (code.includes("permission-denied") && state.isTeacher) {
      alert("Sem permissão para ler quizzes_posts. Isso normalmente significa que as regras novas ainda não foram publicadas no Firebase.");
    }
  }
}

async function refreshCustomQuizzes(options = {}) {
  if (!window.firebaseDB || !window.firebaseCollection || !window.firebaseGetDocs) {
    state.customQuizzes = [];
    state.quizPosts = [];
    return;
  }

  try {
    if (state.isTeacher) {
      const customRef = window.firebaseCollection(window.firebaseDB, "quizzes_custom");
      const snap = await window.firebaseGetDocs(customRef);
      const nextCustom = [];

      const loaded = await Promise.all(snap.docs.map(async (docSnap) => {
        const data = docSnap.data() || {};
        if (data?.ativo === false) {
          return null;
        }

        const quizData = await loadCustomQuizFromFirestore(docSnap.id, data);
        return quizData;
      }));

      loaded.forEach((quiz) => {
        if (quiz) {
          nextCustom.push(quiz);
        }
      });

      state.customQuizzes = nextCustom;
      await refreshQuizPosts({ render: false });
    } else {
      await refreshQuizPosts({ render: false });
      const uniqueQuizIds = Array.from(new Set((state.quizPosts || []).map((post) => String(post.quizId || "").trim()).filter(Boolean)));
      const sourceQuizMap = {};

      await Promise.all(uniqueQuizIds.map(async (quizId) => {
        const sourceQuiz = await loadCustomQuizFromFirestore(quizId, null);
        if (sourceQuiz) {
          sourceQuizMap[quizId] = sourceQuiz;
        }
      }));

      const postedVisible = (state.quizPosts || [])
        .map((post) => buildPostedQuizFromPost(post, post.id, sourceQuizMap[String(post.quizId || "").trim()] || null))
        .filter(Boolean);

      // Se houver mais de uma postagem do mesmo quiz para a turma, mantém a mais recente por atualizadoEmIso.
      const byQuizId = {};
      postedVisible.forEach((quiz) => {
        const existing = byQuizId[quiz.id];
        if (!existing) {
          byQuizId[quiz.id] = quiz;
          return;
        }
        const currentTs = Date.parse(existing.atualizadoEmIso || 0) || 0;
        const nextTs = Date.parse(quiz.atualizadoEmIso || 0) || 0;
        if (nextTs >= currentTs) {
          byQuizId[quiz.id] = quiz;
        }
      });

      state.customQuizzes = Object.values(byQuizId);
    }

    if (options.render !== false) {
      renderHome();
    }
  } catch (error) {
    console.error("Erro ao carregar quizzes customizados:", error);
  }
}

async function setQuizConfig(quizId, patch) {
  if (!window.firebaseDB || !window.firebaseDoc || !window.firebaseSetDoc) {
    throw new Error("Firebase indisponível para atualizar configuração do quiz.");
  }

  const cfgRef = window.firebaseDoc(window.firebaseDB, "quizzes_config", quizId);
  await window.firebaseSetDoc(cfgRef, {
    ...patch,
    atualizadoEmIso: new Date().toISOString(),
    atualizadoPorUid: state.user?.uid || "",
    atualizadoPorEmail: state.user?.email || ""
  }, { merge: true });
}

async function postQuizToClassroom(quizId, classroomId) {
  if (!state.isTeacher) {
    return;
  }

  const quiz = state.customQuizzes.find((item) => item.id === quizId);
  const classroom = state.classrooms.find((item) => item.id === classroomId);
  if (!quiz || !classroom) {
    showBottomNotice("Quiz ou turma não encontrados para postagem.", "error");
    return;
  }

  if (!window.firebaseDB || !window.firebaseDoc || !window.firebaseSetDoc) {
    showBottomNotice("Firebase indisponível para postar quiz para turma.", "error");
    return;
  }

  try {
    const postId = `${quiz.id}__${classroom.id}`;
    const postRef = window.firebaseDoc(window.firebaseDB, "quizzes_posts", postId);
    await window.firebaseSetDoc(postRef, {
      quizId: quiz.id,
      turmaId: classroom.id,
      turmaNome: classroom.nome,
      title: quiz.title,
      description: quiz.description,
      duration: quiz.duration,
      maxAttempts: Math.max(1, Number.parseInt(quiz.maxAttempts, 10) || 1),
      alternativesVisibilityMode: quiz.alternativesVisibilityMode === "hidden" ? "hidden" : "revealed",
      questionCount: Array.isArray(quiz.questions) ? quiz.questions.length : 0,
      allowStudentReview: Boolean(quiz.allowStudentReview),
      allowStudentStart: true,
      allowFocusExitAfterFinish: Boolean(quiz.allowFocusExitAfterFinish),
      ativo: true,
      publicadoEmIso: new Date().toISOString(),
      publicadoPorUid: state.user?.uid || "",
      publicadoPorEmail: state.user?.email || "",
      atualizadoEmIso: new Date().toISOString()
    }, { merge: true });

    await refreshQuizPosts({ render: false });
    renderHome();
    showBottomNotice(`Quiz '${quiz.title}' visível para '${classroom.nome}'.`, "success");
  } catch (error) {
    console.error("Erro ao postar quiz para turma:", error);
    const code = String(error?.code || "").toLowerCase();
    if (code.includes("permission-denied")) {
      const diagnosis = await diagnoseTeacherAccess();
      const base = "Sem permissão para postar quiz em quizzes_posts.";
      if (diagnosis?.anyActiveMatch) {
        showBottomNotice(`${base} Regras do Firestore parecem desatualizadas no projeto ativo.`, "error", 4300);
      } else if (diagnosis) {
        showBottomNotice(`${base} Esta conta não foi encontrada como professor permitido.`, "error", 4300);
      } else {
        showBottomNotice(`${base} Não foi possível diagnosticar automaticamente.`, "error", 4300);
      }
      return;
    }
    if (code.includes("unavailable")) {
      showBottomNotice("Firestore indisponível no momento. Tente novamente em instantes.", "error", 3600);
      return;
    }
    showBottomNotice(`Não foi possível postar o quiz para a turma selecionada.`, "error", 3600);
  }
}

async function removeQuizPostFromClassroom(postId, quizTitle, classroomName) {
  if (!state.isTeacher) {
    return;
  }

  if (!postId || !window.firebaseDB || !window.firebaseDoc || !window.firebaseSetDoc) {
    showBottomNotice("Firebase indisponível para remover acesso da turma.", "error");
    return;
  }

  try {
    const postRef = window.firebaseDoc(window.firebaseDB, "quizzes_posts", postId);
    await window.firebaseSetDoc(postRef, {
      ativo: false,
      removidoEmIso: new Date().toISOString(),
      removidoPorUid: state.user?.uid || "",
      removidoPorEmail: state.user?.email || "",
      atualizadoEmIso: new Date().toISOString()
    }, { merge: true });

    await refreshQuizPosts({ render: false });
    renderHome();
    showBottomNotice(`A turma '${classroomName}' não verá mais '${quizTitle}'. Dados dos alunos preservados.`, "success", 3400);
  } catch (error) {
    console.error("Erro ao remover acesso do quiz para turma:", error);
    showBottomNotice("Não foi possível remover o acesso desta turma agora.", "error", 3400);
  }
}

async function toggleQuizReview(quizId) {
  if (!state.isTeacher) {
    return;
  }

  const quiz = getAllQuizzes().find((item) => item.id === quizId);
  if (!quiz) {
    return;
  }

  try {
    await setQuizConfig(quizId, { allowStudentReview: !Boolean(quiz.allowStudentReview), hidden: false });
    await refreshQuizConfigs({ render: false });
    await refreshCustomQuizzes();
  } catch (error) {
    console.error("Erro ao atualizar liberação do resumo:", error);
    alert("Não foi possível atualizar a liberação de resumo.");
  }
}

async function toggleQuizStart(quizId) {
  if (!state.isTeacher) {
    return;
  }

  const quiz = getAllQuizzes().find((item) => item.id === quizId);
  if (!quiz) {
    return;
  }

  try {
    await setQuizConfig(quizId, { allowStudentStart: !Boolean(quiz.allowStudentStart), hidden: false });
    await refreshQuizConfigs({ render: false });
    await refreshCustomQuizzes();
  } catch (error) {
    console.error("Erro ao atualizar liberação de início do quiz:", error);
    alert("Não foi possível atualizar a liberação de início do quiz.");
  }
}

async function toggleQuizFocusExitRelease(quizId) {
  if (!state.isTeacher) {
    return;
  }

  const quiz = getAllQuizzes().find((item) => item.id === quizId);
  if (!quiz) {
    return;
  }

  const nextValue = !Boolean(quiz.allowFocusExitAfterFinish);

  try {
    await setQuizConfig(quizId, { allowFocusExitAfterFinish: nextValue, hidden: false });
    await refreshQuizConfigs({ render: false });
    await refreshCustomQuizzes();
    if (nextValue) {
      showBottomNotice(`${FOCUS_WAIT_PROCESS_NAME}: alunos podem sair do foco sem penalidade neste quiz.`, "success", 3600);
    } else {
      showBottomNotice(`${FOCUS_WAIT_PROCESS_NAME}: modo de espera reativado e penalidade voltou a valer até nova liberação.`, "info", 3800);
    }
  } catch (error) {
    console.error("Erro ao atualizar liberação do modo de espera:", error);
    alert("Não foi possível atualizar a liberação do modo de espera.");
  }
}

async function removeQuiz(quizId) {
  if (!state.isTeacher) {
    alert("Apenas professores podem remover quizzes.");
    return;
  }

  const quiz = getAllQuizzes().find((item) => item.id === quizId);
  if (!quiz) {
    alert("Quiz não encontrado para remoção.");
    return;
  }

  const ok = confirm(`Deseja remover o quiz '${quiz.title}'?`);
  if (!ok) {
    return;
  }

  const username = (state.studentName || state.user?.displayName || state.user?.email || "").trim();
  if (!username) {
    alert("Não foi possível validar o usuário atual para confirmar a exclusão.");
    return;
  }

  const typedUsername = prompt(`Para confirmar a exclusão, digite seu nome de usuário exatamente como está: ${username}`);
  if (typedUsername === null) {
    return;
  }

  if (typedUsername.trim() !== username) {
    alert("Nome de usuário incorreto. Exclusão cancelada.");
    return;
  }

  if (!window.firebaseDB || !window.firebaseDoc || !window.firebaseSetDoc) {
    alert("Firebase indisponível para remover quiz.");
    return;
  }

  try {
    if (quiz.isCustom) {
      const quizRef = window.firebaseDoc(window.firebaseDB, "quizzes_custom", quizId);
      const answersRef = window.firebaseDoc(window.firebaseDB, "quizzes_answer_keys", quizId);
      const questionsRef = window.firebaseCollection ? window.firebaseCollection(window.firebaseDB, "quizzes_custom", quizId, "questions") : null;
      if (window.firebaseDeleteDoc) {
        let questionDeletes = [];
        if (questionsRef && window.firebaseGetDocs) {
          try {
            const questionsSnap = await window.firebaseGetDocs(questionsRef);
            questionDeletes = questionsSnap.docs.map((docSnap) => window.firebaseDeleteDoc(docSnap.ref));
          } catch (error) {
            console.error("Erro ao listar questões para exclusão do quiz:", error);
          }
        }
        await Promise.all([
          window.firebaseDeleteDoc(quizRef),
          window.firebaseDeleteDoc(answersRef),
          ...questionDeletes
        ]);
      } else {
        await Promise.all([
          window.firebaseSetDoc(quizRef, {
            ativo: false,
            questionCount: 0,
            atualizadoEmIso: new Date().toISOString()
          }, { merge: true }),
          window.firebaseSetDoc(answersRef, {
            answers: {},
            atualizadoEmIso: new Date().toISOString()
          }, { merge: true })
        ]);
      }
    }

    await setQuizConfig(quizId, { hidden: true });
    await refreshQuizConfigs();
    await refreshCustomQuizzes();
  } catch (error) {
    console.error("Erro ao remover quiz:", error);
    alert("Não foi possível remover o quiz.");
  }
}

function getProfileStorageKey() {
  return `quiz_profile_${state.user?.uid || "anon"}`;
}

function getRenameStorageKey() {
  return `quiz_last_rename_${state.user?.uid || "anon"}`;
}

function getLastRenameDate() {
  return localStorage.getItem(getRenameStorageKey()) || null;
}

function saveLastRenameDate() {
  const today = new Date().toISOString().slice(0, 10); // YYYY-MM-DD
  localStorage.setItem(getRenameStorageKey(), today);
}

function canRenameToday() {
  const last = getLastRenameDate();
  if (!last) return true;
  const today = new Date().toISOString().slice(0, 10);
  return last !== today;
}

function openRenameProfileScreen() {
  if (!canRenameToday()) {
    const last = getLastRenameDate();
    const next = new Date(last + "T00:00:00");
    next.setDate(next.getDate() + 1);
    const nextStr = next.toLocaleDateString("pt-BR");
    alert(`Você já alterou seu nome hoje. Tente novamente a partir de ${nextStr}.`);
    return;
  }

  state.isRenamingProfile = true;
  if (profileScreenTitle) profileScreenTitle.textContent = "Alterar nome";
  if (profileScreenSubtitle) profileScreenSubtitle.textContent = "Você pode alterar seu nome uma vez por dia.";
  if (profileSubmitButton) profileSubmitButton.textContent = "Salvar novo nome";
  if (profileCancelButton) profileCancelButton.classList.remove("hidden");
  profileMessage.textContent = "";
  profileNameInput.value = state.studentName || "";
  if (profileClassroomSelect) {
    renderProfileClassroomOptions(state.classroomId);
    profileClassroomSelect.disabled = true;
  }

  authScreen.classList.add("hidden");
  homeScreen.classList.add("hidden");
  profileScreen.classList.remove("hidden");
  showOnlyScreen(null);
}

async function showProfileScreen() {
  endBootPhase();
  state.isRenamingProfile = false;
  if (profileScreenTitle) profileScreenTitle.textContent = "Primeiro acesso";
  if (profileScreenSubtitle) profileScreenSubtitle.textContent = "Antes de continuar, informe o nome que será usado em todos os quizzes.";
  if (profileSubmitButton) profileSubmitButton.textContent = "Salvar e continuar";
  if (profileCancelButton) profileCancelButton.classList.add("hidden");
  profileMessage.textContent = "";
  profileNameInput.value = state.studentName || "";
  await refreshClassrooms();
  if (profileClassroomSelect) {
    profileClassroomSelect.disabled = false;
    renderProfileClassroomOptions(state.classroomId);
  }
  // Mensagem específica é exibida no formulário quando não há turmas.
  authScreen.classList.add("hidden");
  homeScreen.classList.add("hidden");
  profileScreen.classList.remove("hidden");
  userChip.classList.add("hidden");
  if (homeHeaderButton) {
    homeHeaderButton.classList.add("hidden");
  }
  logoutButton.classList.add("hidden");
  showOnlyScreen(null);
}

function setupUserProfileListener() {
  // Desinscrever de qualquer listener anterior
  if (userProfileListenerUnsubscribe) {
    userProfileListenerUnsubscribe();
    userProfileListenerUnsubscribe = null;
  }

  // Configurar listener para mudanças no perfil do usuário
  if (!window.firebaseDoc || !window.firebaseDB || !window.firebaseOnSnapshot || !state.user?.uid) {
    return;
  }

  try {
    const indexRef = window.firebaseDoc(window.firebaseDB, "usuarios_index", state.user.uid);
    userProfileListenerUnsubscribe = window.firebaseOnSnapshot(indexRef, (doc) => {
      if (!doc.exists()) return;

      const data = doc.data();
      const newClassroomId = data.turmaId || "";
      const newClassroomName = data.turmaNome || "";

      // Verificar se a turma mudou
      if (newClassroomId && newClassroomId !== state.classroomId) {
        console.log(`Turma do estudante mudou de "${state.classroomId}" para "${newClassroomId}"`);

        // Atualizar estado
        state.classroomId = newClassroomId;
        state.classroomName = newClassroomName;

        // Atualizar cache local
        localStorage.setItem(getProfileStorageKey(), JSON.stringify({
          name: state.studentName,
          slug: state.profileSlug,
          classroomId: state.classroomId,
          classroomName: state.classroomName
        }));

        // Recarregar dados de quizzes da turma
        refreshQuizPosts({ render: true });

        // Notificar usuário
        showBottomNotice("Você foi movido para uma nova turma. Recarregando avaliações...", "info", 3000);
      }
    });
  } catch (error) {
    console.error("Erro ao configurar listener de perfil:", error);
  }
}

async function ensureUserProfile() {
  const cachedRaw = localStorage.getItem(getProfileStorageKey());
  if (cachedRaw) {
    const cached = JSON.parse(cachedRaw);
    state.studentName = cached.name;
    state.profileSlug = cached.slug;
    state.classroomId = cached.classroomId || "";
    state.classroomName = cached.classroomName || "";
    await refreshTeacherAccess();
    if (!state.classroomId && !state.isTeacher) {
      await showProfileScreen();
      return;
    }
    showAuthenticatedUI();
    if (!state.isTeacher) {
      setupUserProfileListener();
    }
    return;
  }

  if (window.firebaseDoc && window.firebaseGetDoc && window.firebaseDB) {
    try {
      const indexRef = window.firebaseDoc(window.firebaseDB, "usuarios_index", state.user.uid);
      const indexSnap = await window.firebaseGetDoc(indexRef);
      if (indexSnap.exists()) {
        const data = indexSnap.data();
        state.studentName = data.nome || "";
        state.profileSlug = data.slug || normalizeNameSlug(data.nome || state.user.displayName || state.user.email || "usuario");
        state.classroomId = data.turmaId || "";
        state.classroomName = data.turmaNome || "";
        localStorage.setItem(getProfileStorageKey(), JSON.stringify({
          name: state.studentName,
          slug: state.profileSlug,
          classroomId: state.classroomId,
          classroomName: state.classroomName
        }));
        await refreshTeacherAccess();
        if (!state.classroomId && !state.isTeacher) {
          await showProfileScreen();
          return;
        }
        showAuthenticatedUI();
        if (!state.isTeacher) {
          setupUserProfileListener();
        }
        return;
      }
    } catch (error) {
      console.error("Erro ao carregar perfil:", error);
    }
  }

  await refreshTeacherAccess();
  await showProfileScreen();
}

async function handleProfileSubmit(event) {
  event.preventDefault();
  profileMessage.textContent = "";

  const typedName = profileNameInput.value.trim();
  if (!typedName) {
    profileNameInput.focus();
    return;
  }

  const selectedClassroomId = profileClassroomSelect ? String(profileClassroomSelect.value || "") : "";
  // Professores não precisam selecionar turma
  if (!state.isRenamingProfile && !selectedClassroomId && !state.isTeacher) {
    profileMessage.textContent = "Selecione sua turma para continuar.";
    if (profileClassroomSelect) profileClassroomSelect.focus();
    return;
  }

  // Enforce daily rename limit when editing (not first access)
  if (state.isRenamingProfile) {
    if (!canRenameToday()) {
      const last = getLastRenameDate();
      const next = new Date(last + "T00:00:00");
      next.setDate(next.getDate() + 1);
      profileMessage.textContent = `Você já alterou seu nome hoje. Tente novamente a partir de ${next.toLocaleDateString("pt-BR")}.`;
      return;
    }
  }

  const slug = normalizeNameSlug(typedName);

  // Check if name is already taken by another user
  if (window.firebaseDoc && window.firebaseGetDoc && window.firebaseDB) {
    try {
      profileMessage.textContent = "Verificando disponibilidade do nome...";
      const userFolderRef = window.firebaseDoc(window.firebaseDB, "usuarios", slug);
      const existingSnap = await window.firebaseGetDoc(userFolderRef);
      if (existingSnap.exists()) {
        const existingUid = existingSnap.data()?.ultimoUid;
        if (existingUid && existingUid !== state.user?.uid) {
          profileMessage.textContent = "Este nome já está em uso por outro usuário. Escolha um nome diferente.";
          return;
        }
      }
      profileMessage.textContent = "";
    } catch (error) {
      profileMessage.textContent = "Não foi possível verificar o nome agora.";
      console.error("Erro ao verificar nome:", error);
      return;
    }
  }

  state.studentName = typedName;
  state.profileSlug = slug;
  if (!state.isRenamingProfile) {
    state.classroomId = selectedClassroomId;
    const selectedClassroom = state.classrooms.find((item) => item.id === selectedClassroomId);
    state.classroomName = selectedClassroom?.nome || "";
  }

  localStorage.setItem(getProfileStorageKey(), JSON.stringify({
    name: typedName,
    slug,
    classroomId: state.classroomId,
    classroomName: state.classroomName
  }));

  if (state.isRenamingProfile) {
    saveLastRenameDate();
  }

  if (window.firebaseDoc && window.firebaseSetDoc && window.firebaseDB && state.user) {
    try {
      const userFolderRef = window.firebaseDoc(window.firebaseDB, "usuarios", slug);
      const identityRef = window.firebaseDoc(window.firebaseDB, "usuarios", slug, "identidades", state.user.uid);
      const indexRef = window.firebaseDoc(window.firebaseDB, "usuarios_index", state.user.uid);

      await Promise.all([
        window.firebaseSetDoc(userFolderRef, {
          nome: typedName,
          slug,
          turmaId: state.classroomId || "",
          turmaNome: state.classroomName || "",
          atualizadoEmIso: new Date().toISOString(),
          ultimoUid: state.user.uid
        }, { merge: true }),
        window.firebaseSetDoc(identityRef, {
          uid: state.user.uid,
          email: state.user.email || "",
          nome: typedName,
          turmaId: state.classroomId || "",
          turmaNome: state.classroomName || "",
          criadoEmIso: new Date().toISOString()
        }, { merge: true }),
        window.firebaseSetDoc(indexRef, {
          uid: state.user.uid,
          nome: typedName,
          slug,
          turmaId: state.classroomId || "",
          turmaNome: state.classroomName || "",
          email: state.user.email || "",
          atualizadoEmIso: new Date().toISOString()
        }, { merge: true })
      ]);
    } catch (error) {
      profileMessage.textContent = "Não foi possível salvar seu perfil agora.";
      console.error("Erro ao salvar perfil:", error);
      return;
    }
  }

  state.isRenamingProfile = false;
  profileScreen.classList.add("hidden");
  showAuthenticatedUI();
  if (!state.isTeacher) {
    setupUserProfileListener();
  }
}

function showAuthenticatedUI() {
  endBootPhase();
  authScreen.classList.add("hidden");
  profileScreen.classList.add("hidden");
  teacherScreen.classList.add("hidden");
  if (teacherClassroomsScreen) {
    teacherClassroomsScreen.classList.add("hidden");
  }
  if (builderScreen) {
    builderScreen.classList.add("hidden");
  }
  homeScreen.classList.remove("hidden");
  userChip.classList.remove("hidden");
  if (homeHeaderButton) {
    homeHeaderButton.classList.remove("hidden");
  }
  teacherPanelButton.classList.add("hidden");
  if (teacherClassroomsButton) {
    teacherClassroomsButton.classList.add("hidden");
  }
  if (teacherBuilderButton) {
    teacherBuilderButton.classList.add("hidden");
  }
  if (teacherNonogramEditorButton) {
    teacherNonogramEditorButton.classList.add("hidden");
  }
  logoutButton.classList.remove("hidden");

  const display = state.studentName || state.user.displayName || state.user.email || "Usuário";
  userChip.textContent = state.classroomName ? `${display} - ${state.classroomName}` : display;
  if (studentNameDisplay) {
    studentNameDisplay.textContent = state.studentName || "-";
  }

  showOnlyScreen(null);
  renderHome();
  flushPendingUnloadCancellation().catch((error) => {
    console.error("Erro ao processar cancelamento pendente:", error);
  });
  flushPendingFocusWaitPenalty().catch((error) => {
    console.error("Erro ao processar penalidade pendente de espera em foco:", error);
  });
  refreshTeacherAccess({ render: false })
    .then(() => Promise.allSettled([
      refreshRemoteAttempts({ render: false }),
      refreshQuizConfigs({ render: false }),
      refreshRetakeReleases({ render: false }),
      refreshCustomQuizzes({ render: false }),
      refreshClassrooms()
    ]))
    .then(() => {
      renderHome();
    })
    .catch((error) => {
      console.error("Erro ao atualizar dados iniciais da home:", error);
    });
}

function showUnauthenticatedUI() {
  endBootPhase();
  state.isTeacher = false;
  state.studentName = "";
  state.profileSlug = "";
  state.classroomId = "";
  state.classroomName = "";
  state.classrooms = [];
  state.customQuizzes = [];
  state.quizPosts = [];
  state.quizConfigs = {};
  state.remoteAttemptsByQuiz = {};
  state.remoteAttemptCountsByQuiz = {};
  state.remoteAttemptsLoaded = false;
  state.retakeReleases = {};
  authScreen.classList.remove("hidden");
  profileScreen.classList.add("hidden");
  teacherScreen.classList.add("hidden");
  if (teacherClassroomsScreen) {
    teacherClassroomsScreen.classList.add("hidden");
  }
  if (builderScreen) {
    builderScreen.classList.add("hidden");
  }
  homeScreen.classList.add("hidden");
  userChip.classList.add("hidden");
  if (homeHeaderButton) {
    homeHeaderButton.classList.add("hidden");
  }
  teacherPanelButton.classList.add("hidden");
  if (teacherClassroomsButton) {
    teacherClassroomsButton.classList.add("hidden");
  }
  if (teacherBuilderButton) {
    teacherBuilderButton.classList.add("hidden");
  }
  if (teacherNonogramEditorButton) {
    teacherNonogramEditorButton.classList.add("hidden");
  }
  logoutButton.classList.add("hidden");
  showOnlyScreen(null);
  renderFocusWaitPersistentNotice();
}

async function refreshTeacherAccess(options = {}) {
  state.isTeacher = false;
  state.teacherAccessSource = "";
  teacherPanelButton.classList.add("hidden");
  if (teacherClassroomsButton) {
    teacherClassroomsButton.classList.add("hidden");
  }
  if (teacherBuilderButton) {
    teacherBuilderButton.classList.add("hidden");
  }
  if (teacherNonogramEditorButton) {
    teacherNonogramEditorButton.classList.add("hidden");
  }

  if (!state.user?.email || !window.firebaseDoc || !window.firebaseGetDoc || !window.firebaseDB) {
    return;
  }

  const emailKey = state.user.email.trim().toLowerCase();
  const uidKey = String(state.user.uid || "").trim();
  const emailRawKey = state.user.email.trim();

  try {
    const teacherIds = [uidKey, emailRawKey, emailKey].filter(Boolean);
    let allowed = false;

    for (const teacherId of teacherIds) {
      const teacherRef = window.firebaseDoc(window.firebaseDB, "professores_permitidos", teacherId);
      const teacherSnap = await window.firebaseGetDoc(teacherRef);
      if (teacherSnap.exists() && teacherSnap.data()?.ativo !== false) {
        allowed = true;
        state.teacherAccessSource = teacherId;
        break;
      }
    }

    state.isTeacher = Boolean(allowed);

    if (state.isTeacher) {
      teacherPanelButton.classList.remove("hidden");
      if (teacherClassroomsButton) {
        teacherClassroomsButton.classList.remove("hidden");
      }
      if (teacherBuilderButton) {
        teacherBuilderButton.classList.remove("hidden");
      }
      if (teacherNonogramEditorButton) {
        teacherNonogramEditorButton.classList.remove("hidden");
      }
    }

    // Garante que os cards reflitam imediatamente permissões de professor
    if (options.render !== false) {
      renderHome();
    }
  } catch (error) {
    console.error("Erro ao validar permissão de professor:", error);
  }
}

async function diagnoseTeacherAccess() {
  if (!state.user || !window.firebaseDB || !window.firebaseDoc || !window.firebaseGetDoc) {
    return null;
  }

  const emailRaw = String(state.user.email || "").trim();
  const emailLower = emailRaw.toLowerCase();
  const uid = String(state.user.uid || "").trim();
  const checkedIds = [uid, emailRaw, emailLower].filter(Boolean);

  let anyActiveMatch = false;
  let activeSource = "";

  for (const id of checkedIds) {
    try {
      const ref = window.firebaseDoc(window.firebaseDB, "professores_permitidos", id);
      const snap = await window.firebaseGetDoc(ref);
      if (snap.exists() && snap.data()?.ativo !== false) {
        anyActiveMatch = true;
        activeSource = id;
        break;
      }
    } catch (error) {
      console.error("Erro ao diagnosticar permissão de professor:", error);
    }
  }

  return {
    email: emailRaw,
    uid,
    checkedIds,
    anyActiveMatch,
    activeSource
  };
}

function parsePtBrDate(dateString) {
  if (!dateString || typeof dateString !== "string") {
    return 0;
  }
  const parts = dateString.split(" ");
  if (parts.length < 2) {
    return 0;
  }
  const dmy = parts[0].split("/");
  const hms = parts[1].split(":");
  if (dmy.length !== 3 || hms.length < 2) {
    return 0;
  }
  const day = Number(dmy[0]);
  const month = Number(dmy[1]) - 1;
  const year = Number(dmy[2]);
  const hour = Number(hms[0]);
  const minute = Number(hms[1]);
  const second = Number(hms[2] || 0);
  return new Date(year, month, day, hour, minute, second).getTime();
}

function computeAttemptDurationSeconds(startMs, endMs = Date.now()) {
  const start = Number(startMs || 0);
  const end = Number(endMs || 0);
  if (!Number.isFinite(start) || !Number.isFinite(end) || start <= 0 || end <= start) {
    return 0;
  }
  return Math.max(0, Math.round((end - start) / 1000));
}

function formatDurationFromSeconds(totalSeconds) {
  const value = Number(totalSeconds || 0);
  if (!Number.isFinite(value) || value <= 0) {
    return "-";
  }

  const hours = Math.floor(value / 3600);
  const minutes = Math.floor((value % 3600) / 60);
  const seconds = value % 60;

  if (hours > 0) {
    return `${hours}h ${String(minutes).padStart(2, "0")}min ${String(seconds).padStart(2, "0")}s`;
  }
  if (minutes > 0) {
    return `${minutes}min ${String(seconds).padStart(2, "0")}s`;
  }
  return `${seconds}s`;
}

function attachAttemptTimingMetadata(result, endLabel, endMs = Date.now()) {
  if (!result || typeof result !== "object") {
    return;
  }

  const durationSeconds = computeAttemptDurationSeconds(state.attemptStartedAtMs, endMs);
  result.startedAt = result.startedAt || state.attemptStartedAtLabel || "";
  result.endedAt = result.endedAt || endLabel || result.date || "";
  result.durationSeconds = Number.isFinite(Number(result.durationSeconds)) && Number(result.durationSeconds) > 0
    ? Number(result.durationSeconds)
    : durationSeconds;
  result.durationLabel = result.durationLabel || formatDurationFromSeconds(result.durationSeconds);
  result.date = result.date || result.endedAt;
}

function getLatestAttemptsByQuiz(attempts) {
  const latestMap = {};
  attempts.forEach((attempt) => {
    const current = latestMap[attempt.quizId];
    if (!current) {
      latestMap[attempt.quizId] = attempt;
      return;
    }
    const currentTs = parsePtBrDate(current.data);
    const nextTs = parsePtBrDate(attempt.data);
    if (nextTs >= currentTs) {
      latestMap[attempt.quizId] = attempt;
    }
  });
  return Object.values(latestMap);
}

async function releaseQuizForStudent(slug, quizId) {
  try {
    // Remove todas as tentativas do quiz do aluno no Firestore
    if (window.firebaseDB && window.firebaseCollection && window.firebaseQuery && window.firebaseWhere && window.firebaseGetDocs && window.firebaseDeleteDoc) {
      const attemptsRef = window.firebaseCollection(window.firebaseDB, "usuarios", slug, "tentativas");
      const q = window.firebaseQuery(
        attemptsRef,
        window.firebaseWhere("quizId", "==", quizId)
      );
      const snap = await window.firebaseGetDocs(q);
      const deletePromises = [];
      snap.forEach((docSnap) => {
        deletePromises.push(window.firebaseDeleteDoc(docSnap.ref));
      });
      await Promise.all(deletePromises);
    }

    // Remove local attempt para este quiz/aluno
    if (state.user && state.profileSlug === slug) {
      const storageKey = getStorageKey(quizId);
      localStorage.removeItem(storageKey);
      setStoredAttemptCount(quizId, 0);
      clearQuizAccessLock(quizId);
      if (state.remoteAttemptsByQuiz) {
        delete state.remoteAttemptsByQuiz[quizId];
      }
      if (state.remoteAttemptCountsByQuiz) {
        state.remoteAttemptCountsByQuiz[quizId] = 0;
      }
    }

    // Remove liberação de retake
    await setRetakeReleaseForStudent(slug, quizId, false);

    return true;
  } catch (error) {
    console.error("Erro ao limpar tentativas do quiz:", error);
    return false;
  }
}

async function loadLatestAttemptBySlugQuiz(slug, quizId) {
  if (!window.firebaseDB || !window.firebaseCollection || !window.firebaseQuery || !window.firebaseWhere || !window.firebaseGetDocs) {
    return null;
  }

  try {
    const attemptsRef = window.firebaseCollection(window.firebaseDB, "usuarios", slug, "tentativas");
    const q = window.firebaseQuery(
      attemptsRef,
      window.firebaseWhere("quizId", "==", quizId)
    );
    const snap = await window.firebaseGetDocs(q);

    if (snap.empty) {
      return null;
    }

    let latest = null;
    snap.forEach((docSnap) => {
      const data = docSnap.data();
      if (!latest || parsePtBrDate(data.data || "") >= parsePtBrDate(latest.data || "")) {
        latest = data;
      }
    });

    if (!latest) {
      return null;
    }

    return {
      status: latest.status || "completed",
      cancelReason: latest.motivoCancelamento || "",
      blockedByViolation: Boolean(latest.bloqueadoPorViolacao),
      focusExitPenaltyApplied: Boolean(latest.penalidadeSaidaFoco),
      focusExitPenaltyReason: String(latest.penalidadeSaidaFocoMotivo || ""),
      focusExitPenaltyAt: String(latest.penalidadeSaidaFocoEm || ""),
      result: {
        quizId: latest.quizId,
        quizTitle: latest.quizTitulo,
        studentName: latest.nome,
        earnedPoints: latest.pontos ?? 0,
        maxPoints: latest.total ?? 0,
        percent: latest.percentual ?? 0,
        date: latest.data,
        startedAt: latest.inicioAvaliacao || "",
        endedAt: latest.fimAvaliacao || latest.data,
        durationSeconds: Number(latest.duracaoSegundos || 0),
        durationLabel: latest.duracaoTexto || "",
        questionResults: latest.correcaoQuestoes || [],
        focusExitPenaltyApplied: Boolean(latest.penalidadeSaidaFoco),
        focusExitPenaltyReason: String(latest.penalidadeSaidaFocoMotivo || ""),
        focusExitPenaltyAt: String(latest.penalidadeSaidaFocoEm || "")
      },
      answers: latest.respostas || {},
      questionResults: latest.correcaoQuestoes || []
    };
  } catch (error) {
    console.error("Erro ao carregar tentativa do aluno:", error);
    return null;
  }
}

function openTeacherAttemptReview(attempt) {
  if (!attempt) {
    return;
  }

  const quizId = String(attempt.quizId || attempt?.result?.quizId || "");
  const quizTitle = String(attempt.quizTitulo || attempt?.result?.quizTitle || quizId || "Quiz removido");
  const points = Number(attempt.pontos ?? attempt?.result?.earnedPoints ?? 0);
  const total = Number(attempt.total ?? attempt?.result?.maxPoints ?? 0);
  const percent = Number(attempt.percentual ?? attempt?.result?.percent ?? 0);
  const storedQuestionResults = Array.isArray(attempt.correcaoQuestoes)
    ? attempt.correcaoQuestoes
    : (Array.isArray(attempt.questionResults)
      ? attempt.questionResults
      : (Array.isArray(attempt?.result?.questionResults) ? attempt.result.questionResults : []));

  const quiz = getAllQuizzes().find((item) => item.id === quizId) || {
    id: quizId || "quiz-removido",
    title: quizTitle,
    questions: storedQuestionResults.map((item, index) => ({
      id: String(item?.questionId || `q${index + 1}`),
      title: String(item?.questionTitle || `Questão ${index + 1}`),
      explanation: String(item?.questionExplanation || ""),
      options: Array.isArray(item?.options) ? item.options : []
    }))
  };

  if (!reviewList || !reviewSummaryMeta) {
    alert("Não foi possível abrir o resumo deste quiz.");
    return;
  }

  state.reviewBackScreen = "teacher";
  reviewSummaryMeta.textContent = `${points} / ${total || quiz.questions.length} pontos • ${percent}% de acertos`;

  const questionResults = storedQuestionResults;
  if (questionResults.length > 0) {
    reviewList.innerHTML = questionResults.map((item, index) => {
      const statusClass = item.isCorrect ? "review-card-correct" : (item.selectedValue ? "review-card-wrong" : "review-card-blank");
      const statusText = item.selectedValue ? (item.isCorrect ? "Acertou" : "Errou") : "Em branco";
      const allOptionsHtml = renderReviewOptionsHtml(quiz, item, index);
      const questionExplanation = getReviewExplanationForQuestion(quiz, item, index);

      return `
        <article class="builder-question-card review-card ${statusClass}">
          <div class="builder-question-head">
            <h4>Questão ${index + 1}</h4>
            <span class="quiz-status ${item.isCorrect ? "available" : "cancelled"}">${statusText}</span>
          </div>
          <p class="review-question-title">${item.questionTitle || `Questão ${index + 1}`}</p>
          <div class="review-options-block">
            <p class="review-options-title">Alternativas</p>
            ${allOptionsHtml}
          </div>
          ${questionExplanation ? `<div class="review-explanation-block"><p class="review-options-title">Explicação</p><p class="review-question-explanation">${escapeHtml(questionExplanation)}</p></div>` : ""}
        </article>
      `;
    }).join("");
  } else {
    reviewList.innerHTML = `<article class="builder-question-card review-card review-card-blank"><p>Este registro é antigo e não possui detalhamento por questão. Novas tentativas mostrarão o resumo completo.</p></article>`;
  }

  showOnlyScreen("review");
}

async function loadTeacherPanelData() {
  if (!window.firebaseDB || !window.firebaseCollection || !window.firebaseGetDocs) {
    teacherMessage.textContent = "Firestore indisponível no momento.";
    return;
  }

  teacherMessage.textContent = "Carregando dados...";
  teacherUsersList.innerHTML = "";

  try {
    refreshClassrooms();
    const usersRef = window.firebaseCollection(window.firebaseDB, "usuarios");
    const usersSnap = await window.firebaseGetDocs(usersRef);

    if (usersSnap.empty) {
      teacherMessage.textContent = "Nenhum aluno encontrado.";
      return;
    }

    const groups = {};
    const userCards = await Promise.all(usersSnap.docs.map(async (userDoc) => {
      const userData = userDoc.data();
      const slug = userDoc.id;
      const turmaNome = String(userData.turmaNome || "").trim() || "Sem turma";
      const attemptsRef = window.firebaseCollection(window.firebaseDB, "usuarios", slug, "tentativas");
      const attemptsSnap = await window.firebaseGetDocs(attemptsRef);

      const attempts = [];
      attemptsSnap.forEach((docSnap) => {
        attempts.push({ id: docSnap.id, ...docSnap.data() });
      });

      const latestAttempts = getLatestAttemptsByQuiz(attempts);

      const rowsHtml = latestAttempts.length
        ? latestAttempts.map((attempt) => `
            <tr>
              <td class="teacher-cell-quiz">${attempt.quizTitulo || attempt.quizId}</td>
              <td class="teacher-cell-score">${attempt.status === "cancelled" ? "-" : `${attempt.pontos ?? 0}/${attempt.total ?? 0}`}</td>
              <td class="teacher-cell-percent">${attempt.status === "cancelled" ? "-" : `${attempt.percentual ?? 0}%`}</td>
              <td class="teacher-status-cell">
                <div class="teacher-status-stack">
                  <span class="quiz-status ${attempt.status === "cancelled" ? "cancelled" : "completed"}">
                    ${attempt.status === "cancelled"
            ? ((attempt.blockedByViolation || attempt.bloqueadoPorViolacao) ? "Bloqueado por violação" : "Bloqueado")
            : "Concluído"}
                  </span>
                  ${attempt.status === "cancelled" && (attempt.motivoCancelamento || attempt.cancelReason)
            ? `<small class="teacher-status-note ${(attempt.blockedByViolation || attempt.bloqueadoPorViolacao) ? "teacher-status-note-danger" : "teacher-status-note-muted"}">${attempt.motivoCancelamento || attempt.cancelReason}</small>`
            : ""}
                  ${(attempt.penalidadeSaidaFoco || attempt.focusExitPenaltyApplied)
            ? `<small class="teacher-status-note teacher-status-note-danger">Penalidade: saiu do foco antes da liberação do professor.</small>`
            : ""}
                </div>
              </td>
              <td class="teacher-cell-date">${attempt.data || "-"}</td>
              <td class="teacher-cell-actions">
                <div class="teacher-row-actions">
                  <button class="btn btn-ghost teacher-release-btn" data-slug="${slug}" data-quizid="${attempt.quizId}">
                    Limpar
                  </button>
                  <button class="btn btn-ghost teacher-summary-btn" data-slug="${slug}" data-quizid="${attempt.quizId}">
                    Resumo
                  </button>
                </div>
              </td>
            </tr>
          `).join("")
        : `<tr><td colspan="6">Sem tentativas registradas.</td></tr>`;

      const userCard = `
        <article class="teacher-user-card">
          <div class="teacher-student-head">
            <h3>${userData.nome || slug}</h3>
            <p class="teacher-student-meta">${userData.email || "Sem e-mail"} • Turma: ${turmaNome}</p>
          </div>
          <div class="teacher-table-wrap">
            <table class="teacher-table">
              <thead>
                <tr>
                  <th class="teacher-col-quiz">Quiz</th>
                  <th class="teacher-col-score">Nota</th>
                  <th class="teacher-col-percent">%</th>
                  <th class="teacher-col-status">Status</th>
                  <th class="teacher-col-date">Data</th>
                  <th class="teacher-col-actions">Ação</th>
                </tr>
              </thead>
              <tbody>${rowsHtml}</tbody>
            </table>
          </div>
        </article>
      `;

      const searchText = `${String(userData.nome || "")} ${String(userData.email || "")} ${String(slug || "")}`
        .toLowerCase()
        .trim();

      return { turmaNome, userCard, slug, userData, searchText };
    }));

    userCards.forEach(({ turmaNome, userCard, searchText }) => {
      if (!groups[turmaNome]) {
        groups[turmaNome] = [];
      }
      groups[turmaNome].push({ html: userCard, searchText });
    });

    const orderedTurmas = Object.keys(groups).sort((a, b) => {
      if (a === "Sem turma") return 1;
      if (b === "Sem turma") return -1;
      return a.localeCompare(b, "pt-BR");
    });

    if (!orderedTurmas.length) {
      teacherUsersList.innerHTML = `<p class="teacher-empty-state">Nenhuma turma encontrada.</p>`;
      teacherMessage.textContent = "";
      return;
    }

    let selectedTurma = orderedTurmas.includes(state.teacherResultsClassroomFilter)
      ? state.teacherResultsClassroomFilter
      : orderedTurmas[0];
    let studentSearch = String(state.teacherResultsStudentSearch || "");
    let optionsOpen = false;

    const getMembrosByTurma = (turmaNome) => {
      const membros = [];
      usersSnap.docs.forEach((userDoc) => {
        const data = userDoc.data() || {};
        const turma = String(data.turmaNome || "").trim() || "Sem turma";
        if (turma === turmaNome) {
          membros.push({ slug: userDoc.id, nome: data.nome || userDoc.id });
        }
      });
      return membros;
    };

    const bindTeacherResultCardActions = () => {
      teacherUsersList.querySelectorAll(".teacher-release-btn").forEach((button) => {
        button.textContent = "Limpar";
        button.title = "Excluir todas as tentativas deste quiz para este aluno";
        button.addEventListener("click", async () => {
          const slug = button.dataset.slug;
          const quizId = button.dataset.quizid;
          if (!slug || !quizId) {
            return;
          }

          const ok = confirm(`Excluir todas as tentativas do quiz '${quizId}' para este aluno? Isso permitirá que ele comece do zero.`);
          if (!ok) {
            return;
          }

          button.disabled = true;
          button.textContent = "Limpando...";

          const cleaned = await releaseQuizForStudent(slug, quizId);
          if (cleaned) {
            button.textContent = "Limpo";
            teacherMessage.textContent = "Tentativas removidas com sucesso.";
            await loadTeacherPanelData();
          } else {
            button.disabled = false;
            button.textContent = "Limpar";
            teacherMessage.textContent = "Erro ao limpar tentativas.";
          }
        });
      });

      teacherUsersList.querySelectorAll(".teacher-summary-btn").forEach((button) => {
        button.addEventListener("click", async () => {
          const slug = button.dataset.slug;
          const quizId = button.dataset.quizid;
          if (!slug || !quizId) {
            return;
          }

          const attempt = await loadLatestAttemptBySlugQuiz(slug, quizId);
          if (!attempt) {
            alert("Este aluno ainda não possui tentativa para este quiz.");
            return;
          }

          openTeacherAttemptReview(attempt);
        });
      });

      const reportButton = teacherUsersList.querySelector(".teacher-classroom-report-btn");
      if (reportButton) {
        reportButton.addEventListener("click", () => {
          const turmaNome = String(reportButton.dataset.turma || "").trim();
          if (!turmaNome) {
            return;
          }
          openClassroomReportModal(turmaNome, getMembrosByTurma(turmaNome));
        });
      }
    };

    const renderSelectedTurmaResults = (options = {}) => {
      state.teacherResultsClassroomFilter = selectedTurma;
      state.teacherResultsStudentSearch = studentSearch;
      const optionsHtml = orderedTurmas
        .map((turmaNome) => `
          <button type="button" class="btn btn-ghost btn-sm teacher-classroom-option-btn ${turmaNome === selectedTurma ? "is-selected" : ""}" data-select-turma="${escapeHtml(turmaNome)}">${escapeHtml(turmaNome)}</button>
        `)
        .join("");

      const normalizedSearch = studentSearch.trim().toLowerCase();
      const turmaEntries = groups[selectedTurma] || [];
      const filteredEntries = normalizedSearch
        ? turmaEntries.filter((entry) => String(entry.searchText || "").includes(normalizedSearch))
        : turmaEntries;
      const usersHtml = filteredEntries.length
        ? filteredEntries.map((entry) => entry.html).join("")
        : `<p class="teacher-empty-state">Nenhum aluno encontrado para o filtro informado.</p>`;

      teacherUsersList.innerHTML = `
        <section class="teacher-classroom-section">
          <div class="teacher-classroom-filter-bar">
            <button type="button" class="teacher-classroom-filter-chip teacher-classroom-filter-chip-btn" data-toggle-turma-options aria-expanded="${optionsOpen ? "true" : "false"}" title="Trocar turma">
              <span class="teacher-classroom-chip-name">${escapeHtml(selectedTurma)}</span>
              <span class="teacher-classroom-chip-caret">${optionsOpen ? "▴" : "▾"}</span>
            </button>
            <input type="text" class="teacher-student-search-input" data-student-search placeholder="Buscar aluno" value="${escapeHtml(studentSearch)}" autocomplete="off" />
            <button type="button" class="btn btn-ghost btn-sm teacher-classroom-report-btn" data-turma="${escapeHtml(selectedTurma)}">Relatório da turma</button>
          </div>
          <div class="teacher-classroom-options ${optionsOpen ? "" : "hidden"}" data-turma-options>${optionsHtml}</div>
          <div class="teacher-users-list">
            ${usersHtml}
          </div>
        </section>
      `;

      const toggleButton = teacherUsersList.querySelector("[data-toggle-turma-options]");
      const optionsPanel = teacherUsersList.querySelector("[data-turma-options]");
      if (toggleButton && optionsPanel) {
        toggleButton.addEventListener("click", () => {
          optionsOpen = !optionsOpen;
          optionsPanel.classList.toggle("hidden", !optionsOpen);
          toggleButton.setAttribute("aria-expanded", optionsOpen ? "true" : "false");
          const caret = toggleButton.querySelector(".teacher-classroom-chip-caret");
          if (caret) {
            caret.textContent = optionsOpen ? "▴" : "▾";
          }
        });
      }

      const studentSearchInput = teacherUsersList.querySelector("[data-student-search]");
      if (studentSearchInput) {
        studentSearchInput.addEventListener("input", () => {
          const selectionStart = studentSearchInput.selectionStart;
          const selectionEnd = studentSearchInput.selectionEnd;
          studentSearch = String(studentSearchInput.value || "");
          renderSelectedTurmaResults({
            preserveSearchFocus: true,
            searchSelectionStart: selectionStart,
            searchSelectionEnd: selectionEnd
          });
        });
      }

      teacherUsersList.querySelectorAll("[data-select-turma]").forEach((button) => {
        button.addEventListener("click", () => {
          const turma = String(button.dataset.selectTurma || "").trim();
          if (!turma || turma === selectedTurma) {
            return;
          }
          selectedTurma = turma;
          optionsOpen = false;
          renderSelectedTurmaResults();
        });
      });

      bindTeacherResultCardActions();

      if (options.preserveSearchFocus) {
        const nextSearchInput = teacherUsersList.querySelector("[data-student-search]");
        if (nextSearchInput) {
          const start = Number(options.searchSelectionStart);
          const end = Number(options.searchSelectionEnd);
          requestAnimationFrame(() => {
            nextSearchInput.focus();
            if (Number.isFinite(start) && Number.isFinite(end)) {
              nextSearchInput.setSelectionRange(start, end);
            }
          });
        }
      }
    };

    renderSelectedTurmaResults();
    teacherMessage.textContent = "";
  } catch (error) {
    teacherMessage.textContent = "Erro ao carregar painel do professor.";
    console.error(error);
  }
}

// ────────────────────────────────────────────────────────────────────────────
// Relatório de Turma
// ────────────────────────────────────────────────────────────────────────────

const classroomReportModal = document.getElementById("classroom-report-modal");
const classroomReportClose = document.getElementById("classroom-report-close");
const classroomReportQuizSelect = document.getElementById("classroom-report-quiz-select");
const classroomReportLoading = document.getElementById("classroom-report-loading");
const classroomReportBody = document.getElementById("classroom-report-body");
const classroomReportRefresh = document.getElementById("classroom-report-refresh");

if (classroomReportClose) {
  classroomReportClose.addEventListener("click", () => {
    if (classroomReportModal) classroomReportModal.classList.add("hidden");
  });
}

if (classroomReportModal) {
  classroomReportModal.addEventListener("click", (e) => {
    if (e.target === classroomReportModal) classroomReportModal.classList.add("hidden");
  });
}

let _classroomReportMembros = [];
let _classroomReportTurma = "";

if (classroomReportQuizSelect) {
  classroomReportQuizSelect.addEventListener("change", async () => {
    const quizId = classroomReportQuizSelect.value;
    if (!quizId) {
      if (classroomReportBody) classroomReportBody.classList.add("hidden");
      if (classroomReportLoading) classroomReportLoading.classList.add("hidden");
      return;
    }
    await renderClassroomReport(_classroomReportTurma, _classroomReportMembros, quizId);
  });
}

if (classroomReportRefresh) {
  classroomReportRefresh.addEventListener("click", async () => {
    const quizId = classroomReportQuizSelect?.value;
    if (!quizId) return;
    await renderClassroomReport(_classroomReportTurma, _classroomReportMembros, quizId);
  });
}

function openClassroomReportModal(turmaNome, membros) {
  if (!classroomReportModal || !classroomReportQuizSelect) return;

  _classroomReportTurma = turmaNome;
  _classroomReportMembros = membros;

  const modalTitle = document.getElementById("classroom-report-title");
  if (modalTitle) modalTitle.textContent = `Relatório — ${turmaNome}`;

  const quizzes = getAllQuizzes();
  classroomReportQuizSelect.innerHTML = '<option value="">-- Selecione um quiz --</option>' +
    quizzes.map((q) => `<option value="${escapeHtml(q.id)}">${escapeHtml(q.title)}</option>`).join("");
  if (classroomReportBody) classroomReportBody.classList.add("hidden");
  if (classroomReportLoading) classroomReportLoading.classList.add("hidden");

  classroomReportModal.classList.remove("hidden");
}

async function renderClassroomReport(turmaNome, membros, quizId) {
  if (!classroomReportBody || !classroomReportLoading) return;

  classroomReportBody.classList.add("hidden");
  classroomReportLoading.classList.remove("hidden");
  try {
    const quiz = getAllQuizzes().find((q) => q.id === quizId);
    const quizTitle = quiz ? quiz.title : quizId;

    const quizQuestionsLookup = {};
    if (quiz && Array.isArray(quiz.questions)) {
      quiz.questions.forEach((q) => {
        quizQuestionsLookup[String(q.id)] = q;
      });
    }

    const attemptResults = await Promise.all(membros.map(async ({ slug, nome }) => {
      const attempt = await loadLatestAttemptBySlugQuiz(slug, quizId);
      return { slug, nome, attempt };
    }));

    const concluidos = attemptResults.filter((r) => r.attempt && r.attempt.status === "completed");
    const bloqueados = attemptResults.filter((r) => r.attempt && r.attempt.status === "cancelled");
    const semTentativa = attemptResults.filter((r) => !r.attempt);
    const total = concluidos.length;

    let mediaTurma = 0;
    if (total > 0) {
      mediaTurma = Math.round(concluidos.reduce((sum, r) => sum + (r.attempt.result?.percent ?? 0), 0) / total);
    }

    const faixas = [
      { label: "0–49%", min: 0, max: 49, count: 0 },
      { label: "50–69%", min: 50, max: 69, count: 0 },
      { label: "70–89%", min: 70, max: 89, count: 0 },
      { label: "90–100%", min: 90, max: 100, count: 0 }
    ];
    concluidos.forEach(({ attempt }) => {
      const p = attempt.result?.percent ?? 0;
      for (const f of faixas) {
        if (p >= f.min && p <= f.max) { f.count++; break; }
      }
    });

    const questaoMap = {};
    concluidos.forEach(({ attempt }) => {
      const qr = Array.isArray(attempt.questionResults) ? attempt.questionResults
        : Array.isArray(attempt.result?.questionResults) ? attempt.result.questionResults : [];
      qr.forEach((q) => {
        const qid = String(q.questionId || q.id || "");
        if (!qid) return;
        if (!questaoMap[qid]) {
          const quizQ = quizQuestionsLookup[qid];
          questaoMap[qid] = {
            title: q.questionTitle || quizQ?.title || qid,
            options: quizQ?.options || [],
            image: quizQ?.image || "",
            answerToken: quizQ?.answerToken || "",
            acertos: 0, erros: 0, brancos: 0
          };
        }
        if (q.isCorrect) questaoMap[qid].acertos++;
        else if (q.selectedValue) questaoMap[qid].erros++;
        else questaoMap[qid].brancos++;
      });
    });

    const questoes = Object.values(questaoMap);

    const statCards = `
    <div class="cr-stat-row">
      <div class="cr-stat-card"><span class="cr-stat-label">Concluídos</span><strong class="cr-stat-value">${total}</strong></div>
      <div class="cr-stat-card"><span class="cr-stat-label">Bloqueados</span><strong class="cr-stat-value">${bloqueados.length}</strong></div>
      <div class="cr-stat-card"><span class="cr-stat-label">Sem tentativa</span><strong class="cr-stat-value">${semTentativa.length}</strong></div>
      <div class="cr-stat-card cr-stat-card-highlight"><span class="cr-stat-label">Média da turma</span><strong class="cr-stat-value">${total > 0 ? mediaTurma + "%" : "-"}</strong></div>
    </div>`;

    const barras = faixas.map((f) => {
      const pct = total > 0 ? Math.round((f.count / total) * 100) : 0;
      return `
        <div class="cr-bar-row">
          <span class="cr-bar-label">${f.label}</span>
          <div class="cr-bar-track"><div class="cr-bar-fill" style="width:${pct}%"></div></div>
          <span class="cr-bar-pct">${pct}%</span>
          <span class="cr-bar-count">(${f.count})</span>
        </div>`;
    }).join("");

    const questoesRows = questoes.length > 0
      ? questoes.map((q, idx) => {
        const tot = q.acertos + q.erros + q.brancos;
        const acertoPct = tot > 0 ? Math.round((q.acertos / tot) * 100) : 0;
        const erroPct = tot > 0 ? Math.round((q.erros / tot) * 100) : 0;
        const altHtml = q.options.length > 0
          ? '<ul class="cr-alternatives-list">' +
          q.options.map((opt) => {
            const isCorrect = String(opt.value) === String(q.answerToken);
            return `<li class="cr-alt${isCorrect ? " cr-alt-correct" : ""}">${escapeHtml(String(opt.value))}) ${escapeHtml(String(opt.label))}${isCorrect ? ' <span class="cr-alt-check">&#10003;</span>' : ""}</li>`;
          }).join("") +
          "</ul>"
          : "";
        const imgHtml = q.image ? `<img src="${q.image}" style="max-width: 100%; max-height: 200px; border-radius: 8px; margin: 6px auto 0; display: block;" alt="Imagem da questão" />` : "";
        return `
          <div class="cr-question-row">
            <div class="cr-question-header">
              <span class="cr-question-num">Q${idx + 1}</span>
              <div class="cr-question-bars">
                <div class="cr-qbar cr-qbar-correct" style="width:${acertoPct}%" title="Acertos: ${acertoPct}%"></div>
                <div class="cr-qbar cr-qbar-wrong" style="width:${erroPct}%" title="Erros: ${erroPct}%"></div>
              </div>
              <div class="cr-question-pcts">
                <span class="cr-pct-correct">&#10003; ${acertoPct}%</span>
                <span class="cr-pct-wrong">&#10007; ${erroPct}%</span>
              </div>
            </div>
            <p class="cr-question-title">${escapeHtml(String(q.title))}</p>
            ${imgHtml}
            ${altHtml}
          </div>`;
      }).join("")
      : '<p class="cr-empty">Sem dados de questões disponíveis.</p>';

    classroomReportBody.innerHTML = `
      <div class="cr-topline">
        <p class="cr-quiz-title">Quiz: <strong>${escapeHtml(quizTitle)}</strong></p>
        <span class="cr-updated-at">Dados atualizados</span>
      </div>
      <div class="cr-summary-block">
        ${statCards}
      </div>
      <div class="cr-section cr-section-panel">
        <h4 class="cr-section-title">Distribuição de desempenho</h4>
        ${total > 0 ? barras : '<p class="cr-empty">Sem dados suficientes.</p>'}
      </div>
      <div class="cr-section cr-section-panel">
        <h4 class="cr-section-title">Desempenho por questão</h4>
        <div class="cr-questions-list">${questoesRows}</div>
      </div>`;

    classroomReportBody.classList.remove("hidden");
  } catch (error) {
    console.error("Erro ao renderizar relatório da turma:", error);
    classroomReportBody.innerHTML = '<p class="cr-empty">Não foi possível carregar o relatório. Tente novamente.</p>';
    classroomReportBody.classList.remove("hidden");
  } finally {
    classroomReportLoading.classList.add("hidden");
  }
}

async function openTeacherPanel() {
  if (!state.isTeacher) {
    alert("Acesso negado ao painel do professor.");
    return;
  }

  showOnlyScreen(null);
  homeScreen.classList.add("hidden");
  if (teacherClassroomsScreen) {
    teacherClassroomsScreen.classList.add("hidden");
  }
  teacherScreen.classList.remove("hidden");
  await loadTeacherPanelData();
}

async function openTeacherClassroomsScreen() {
  if (!state.isTeacher) {
    alert("Acesso negado às turmas.");
    return;
  }

  showOnlyScreen(null);
  homeScreen.classList.add("hidden");
  teacherScreen.classList.add("hidden");
  if (teacherClassroomsScreen) {
    teacherClassroomsScreen.classList.remove("hidden");
  }
  await loadTeacherClassroomData();
}

function getBuilderQuestionCards() {
  return builderQuestions ? Array.from(builderQuestions.querySelectorAll(".builder-question-card")) : [];
}

function addBuilderOptionRow(optionsEl, value = "", isCorrect = false) {
  const row = document.createElement("div");
  row.className = "builder-option-row";
  row.innerHTML = `
    <label class="builder-correct-toggle">
      <input type="radio" name="" ${isCorrect ? "checked" : ""} />
      <span>Correta</span>
    </label>
    <input type="text" class="builder-option-label-input" placeholder="Texto da alternativa" value="${value.replace(/"/g, "&quot;")}" required />
    <button type="button" class="builder-remove-option" title="Remover alternativa" aria-label="Remover alternativa">
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
        <path d="M9 3v1H4v2h16V4h-5V3H9z" fill="currentColor" />
        <path d="M6 7l1 14a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-14H6z" fill="currentColor" />
      </svg>
    </button>
  `;

  const removeButton = row.querySelector("button");
  removeButton.addEventListener("click", () => {
    row.remove();
    updateBuilderRadioGroups();
  });

  optionsEl.appendChild(row);
  updateBuilderRadioGroups();
}

function updateBuilderRadioGroups() {
  getBuilderQuestionCards().forEach((card, cardIndex) => {
    const radios = card.querySelectorAll(".builder-option-row input[type='radio']");
    radios.forEach((radio) => {
      radio.name = `builder-correct-${cardIndex}`;
    });
  });
}

function fileToCompressedDataUrl(file, maxDimension = 1280, quality = 0.8) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        let { width, height } = img;
        if (width > height && width > maxDimension) {
          height = Math.round((height * maxDimension) / width);
          width = maxDimension;
        } else if (height > maxDimension) {
          width = Math.round((width * maxDimension) / height);
          height = maxDimension;
        }

        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        if (!ctx) {
          reject(new Error("Canvas indisponivel."));
          return;
        }

        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.onerror = () => reject(new Error("Falha ao carregar imagem."));
      img.src = String(reader.result || "");
    };
    reader.onerror = () => reject(new Error("Falha ao ler arquivo."));
    reader.readAsDataURL(file);
  });
}
function addBuilderQuestionCard(seed = {}) {
  if (!builderQuestions) {
    return;
  }

  const card = document.createElement("article");
  card.className = "builder-question-card";
  card.innerHTML = `
    <div class="builder-question-head">
      <h4>Questão</h4>
      <button type="button" class="builder-remove-question" title="Remover questão" aria-label="Remover questão">
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
          <path d="M9 3v1H4v2h16V4h-5V3H9z" fill="currentColor" />
          <path d="M6 7l1 14a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-14H6z" fill="currentColor" />
        </svg>
      </button>
    </div>
    <input type="text" class="builder-question-title" placeholder="Enunciado da questão" value="${(seed.title || "").replace(/"/g, "&quot;")}" required />
    <input type="text" class="builder-question-description" placeholder="Descrição (opcional)" value="${(seed.description || "").replace(/"/g, "&quot;")}" />
    <textarea class="builder-question-explanation" placeholder="Explicação da questão para a revisão (opcional)">${(seed.explanation || "").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</textarea>
    <div class="builder-image-field">
      <label class="builder-field-label">Imagem (opcional)</label>
      <div class="builder-image-upload">
        <input type="file" class="builder-image-input" accept="image/*" />
        <div class="builder-image-preview"></div>
        <input type="hidden" class="builder-image-url" value="${seed.image || ""}" />
      </div>
    </div>
    <div class="builder-question-meta-grid">
      <label class="builder-field-label">
        Pontuação
        <input type="number" class="builder-question-points" min="1" value="${seed.points || 1}" />
      </label>
      <label class="builder-field-label">
        Tempo
        <select class="builder-question-timer">
          <option value="0" ${seed.timer === 0 || !seed.timer ? "selected" : ""}>Sem tempo</option>
          <option value="15" ${seed.timer === 15 ? "selected" : ""}>15 segundos</option>
          <option value="30" ${seed.timer === 30 ? "selected" : ""}>30 segundos</option>
          <option value="45" ${seed.timer === 45 ? "selected" : ""}>45 segundos</option>
          <option value="60" ${seed.timer === 60 ? "selected" : ""}>1 minuto</option>
          <option value="90" ${seed.timer === 90 ? "selected" : ""}>1 minuto e 30 segundos</option>
          <option value="120" ${seed.timer === 120 ? "selected" : ""}>2 minutos</option>
          <option value="180" ${seed.timer === 180 ? "selected" : ""}>3 minutos</option>
        </select>
      </label>
    </div>
    <div class="builder-options"></div>
    <button type="button" class="btn builder-add-option btn-circle-add" title="Adicionar alternativa" aria-label="Adicionar alternativa">+</button>
  `;

  const removeQuestionButton = card.querySelector(".builder-remove-question");
  removeQuestionButton.addEventListener("click", () => {
    card.remove();
    updateBuilderQuestionTitles();
    updateBuilderRadioGroups();
  });

  const optionsEl = card.querySelector(".builder-options");
  const addOptionButton = card.querySelector(".builder-add-option");
  addOptionButton.addEventListener("click", () => addBuilderOptionRow(optionsEl));

  // Image upload handler
  const imageInput = card.querySelector(".builder-image-input");
  const imagePreview = card.querySelector(".builder-image-preview");
  const imageUrlHidden = card.querySelector(".builder-image-url");

  if (seed.image) {
    const previewImg = document.createElement("img");
    previewImg.src = seed.image;
    previewImg.style.maxWidth = "100%";
    previewImg.style.maxHeight = "150px";
    previewImg.style.borderRadius = "8px";
    previewImg.style.marginTop = "8px";
    imagePreview.appendChild(previewImg);
  }

  if (imageInput) {
    imageInput.addEventListener("change", async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;

      if (!file.type.startsWith("image/")) {
        alert("Por favor, selecione uma imagem valida.");
        return;
      }

      if (file.size > 5 * 1024 * 1024) {
        alert("Imagem muito grande. Maximo 5MB.");
        return;
      }

      imageInput.disabled = true;
      imagePreview.innerHTML = "<p style=\"font-size: 0.8rem; color: #666;\">Processando imagem...</p>";

      try {
        let dataUrl = await fileToCompressedDataUrl(file, 1280, 0.8);
        if (dataUrl.length > 1500000) {
          dataUrl = await fileToCompressedDataUrl(file, 960, 0.7);
        }

        imageUrlHidden.value = dataUrl;
        imagePreview.innerHTML = "";

        const previewImg = document.createElement("img");
        previewImg.src = dataUrl;
        previewImg.style.maxWidth = "100%";
        previewImg.style.maxHeight = "150px";
        previewImg.style.borderRadius = "8px";
        previewImg.style.marginTop = "8px";
        imagePreview.appendChild(previewImg);
      } catch (error) {
        console.error("Erro ao processar imagem:", error);
        imagePreview.innerHTML = "<p style=\"color: #c33; font-size: 0.8rem;\">Erro ao processar imagem. Tente novamente.</p>";
      } finally {
        imageInput.disabled = false;
      }
    });
  }

  const seedOptions = Array.isArray(seed.options) && seed.options.length >= 2
    ? seed.options
    : [{ label: "" }, { label: "" }, { label: "" }, { label: "" }];

  seedOptions.forEach((option) => {
    addBuilderOptionRow(optionsEl, option.label || "", option.value === seed.correctAnswer);
  });

  builderQuestions.appendChild(card);
  updateBuilderQuestionTitles();
  updateBuilderRadioGroups();
}

function setBuilderSelectValue(selectEl, nextValue, fallbackValue = "") {
  if (!selectEl) {
    return;
  }

  const normalized = String(nextValue || "").trim();
  const hasOption = Array.from(selectEl.options || []).some((option) => String(option.value) === normalized);
  if (hasOption) {
    selectEl.value = normalized;
    return;
  }

  if (fallbackValue) {
    selectEl.value = fallbackValue;
  }
}

function normalizeGeneratedQuestionSeed(question = {}, index = 0) {
  const rawOptions = Array.isArray(question.options) ? question.options : [];
  const options = rawOptions
    .map((option, optionIndex) => {
      if (typeof option === "string") {
        return {
          value: String.fromCharCode(97 + optionIndex),
          label: option
        };
      }

      if (!option || typeof option.label !== "string") {
        return null;
      }

      return {
        value: String(option.value || String.fromCharCode(97 + optionIndex)),
        label: String(option.label || "").trim()
      };
    })
    .filter((option) => option && option.label);

  const correctAnswer = String(question.correctAnswer || question.correctOptionValue || "").trim();
  const correctAnswerIndex = Number.parseInt(question.correctAnswerIndex, 10);
  let resolvedCorrectAnswer = correctAnswer;

  if (!resolvedCorrectAnswer && Number.isFinite(correctAnswerIndex) && options[correctAnswerIndex]) {
    resolvedCorrectAnswer = options[correctAnswerIndex].value;
  }

  if (!resolvedCorrectAnswer && options[0]) {
    resolvedCorrectAnswer = options[0].value;
  }

  return {
    id: question.id || `q${index + 1}`,
    title: String(question.title || `Questão ${index + 1}`),
    description: String(question.description || "Selecione apenas uma alternativa."),
    explanation: String(question.explanation || ""),
    points: Math.max(1, Number.parseInt(question.points, 10) || 1),
    timer: Math.max(0, Number.parseInt(question.timer, 10) || 0),
    options,
    correctAnswer: resolvedCorrectAnswer,
    image: typeof question.image === "string" ? question.image : ""
  };
}

function getStoredGeminiApiKey() {
  try {
    return String(window.localStorage?.getItem(GEMINI_API_KEY_STORAGE_KEY) || "").trim();
  } catch (_error) {
    return "";
  }
}

function setStoredGeminiApiKey(value = "") {
  const normalized = String(value || "").trim();
  try {
    if (!normalized) {
      window.localStorage?.removeItem(GEMINI_API_KEY_STORAGE_KEY);
      return;
    }
    window.localStorage?.setItem(GEMINI_API_KEY_STORAGE_KEY, normalized);
  } catch (_error) {
    // Ignora bloqueios de armazenamento do navegador.
  }
}

function normalizeAiDifficulty(value) {
  const normalized = String(value || "facil").trim().toLowerCase();
  if (["facil", "medio", "dificil"].includes(normalized)) {
    return normalized;
  }
  return "facil";
}

function normalizeAiDuration(value) {
  const allowed = new Set(["5 min", "8 min", "10 min", "12 min", "15 min", "20 min"]);
  const normalized = String(value || "8 min").trim();
  return allowed.has(normalized) ? normalized : "8 min";
}

function buildGeminiGenerationPrompt({ prompt, questionCount, difficulty, schoolLevel, duration, maxAttempts, alternativesVisibilityMode }) {
  return [
    "Você deve gerar um questionário em JSON puro, sem markdown, sem comentários e sem texto fora do JSON.",
    "Responda em português do Brasil.",
    "O JSON deve seguir exatamente este formato:",
    JSON.stringify({
      title: "Titulo do quiz",
      description: "Descricao curta do quiz",
      duration: "8 min",
      maxAttempts: 1,
      alternativesVisibilityMode: "revealed",
      questions: [
        {
          title: "Enunciado da questao",
          description: "Orientacao curta da questao",
          explanation: "Explicacao curta para revisao",
          points: 1,
          timer: 0,
          options: [
            { label: "Alternativa A" },
            { label: "Alternativa B" },
            { label: "Alternativa C" },
            { label: "Alternativa D" }
          ],
          correctAnswerIndex: 0
        }
      ]
    }),
    `Quantidade de questoes: ${questionCount}.`,
    `Dificuldade: ${difficulty}.`,
    `Nivel escolar: ${schoolLevel || "nao informado"}.`,
    `Duracao sugerida: ${duration}.`,
    `Numero maximo de tentativas: ${maxAttempts}.`,
    `Visualizacao das alternativas: ${alternativesVisibilityMode}.`,
    "Cada questao deve ter entre 4 alternativas, apenas uma correta.",
    "Nao invente campos extras.",
    `Tema e instrucoes do professor: ${prompt}`
  ].join("\n");
}

function sanitizeGeminiTextPayload(text = "") {
  return String(text || "")
    .replace(/^\uFEFF/, "")
    .replace(/```json\s*/gi, "")
    .replace(/```\s*/g, "")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .trim();
}

function tryParseJsonCandidate(candidate = "") {
  const normalized = String(candidate || "").trim();
  if (!normalized) {
    return null;
  }

  const parseAttempts = [
    normalized,
    normalized.replace(/,\s*([}\]])/g, "$1")
  ];

  for (const attempt of parseAttempts) {
    try {
      return JSON.parse(attempt);
    } catch (_error) {
      // Continua tentando outras variações.
    }
  }

  return null;
}

function extractJsonPayload(text = "") {
  const normalized = sanitizeGeminiTextPayload(text);
  if (!normalized) {
    throw new Error("A Gemini nao retornou conteudo utilizavel.");
  }

  const direct = tryParseJsonCandidate(normalized);
  if (direct && typeof direct === "object") {
    return direct;
  }

  const firstBrace = normalized.indexOf("{");
  const lastBrace = normalized.lastIndexOf("}");
  if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
    const slicedObject = tryParseJsonCandidate(normalized.slice(firstBrace, lastBrace + 1));
    if (slicedObject && typeof slicedObject === "object") {
      return slicedObject;
    }
  }

  const firstBracket = normalized.indexOf("[");
  const lastBracket = normalized.lastIndexOf("]");
  if (firstBracket !== -1 && lastBracket !== -1 && lastBracket > firstBracket) {
    const slicedArray = tryParseJsonCandidate(normalized.slice(firstBracket, lastBracket + 1));
    if (Array.isArray(slicedArray) && slicedArray.length && typeof slicedArray[0] === "object") {
      return slicedArray[0];
    }
  }

  throw new Error("A Gemini retornou um formato invalido para o quiz.");
}

async function requestQuizFromGemini({ apiKey, prompt }) {
  const apiVersions = ["v1", "v1beta"];
  const preferredModels = [
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
    "gemini-1.5-flash-latest",
    "gemini-1.5-flash"
  ];

  async function listModelsForVersion(version) {
    const endpoint = `https://generativelanguage.googleapis.com/${version}/models?key=${encodeURIComponent(apiKey)}`;
    const response = await fetch(endpoint);
    if (!response.ok) {
      return [];
    }

    const json = await response.json();
    const models = Array.isArray(json?.models) ? json.models : [];
    return models
      .filter((model) => Array.isArray(model?.supportedGenerationMethods) && model.supportedGenerationMethods.includes("generateContent"))
      .map((model) => String(model?.name || "").replace(/^models\//, "").trim())
      .filter(Boolean);
  }

  let listedModels = [];
  try {
    const listedFromV1 = await listModelsForVersion("v1");
    const listedFromV1Beta = await listModelsForVersion("v1beta");
    listedModels = [...listedFromV1, ...listedFromV1Beta];
  } catch (_error) {
    listedModels = [];
  }

  const modelsToTry = Array.from(new Set([...preferredModels, ...listedModels])).slice(0, 12);
  const nonFatalErrors = [];

  const requestBody = {
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 8192
    },
    contents: [
      {
        role: "user",
        parts: [{ text: prompt }]
      }
    ]
  };

  for (const model of modelsToTry) {
    for (const version of apiVersions) {
      const endpoint = `https://generativelanguage.googleapis.com/${version}/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`;
      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify(requestBody)
      });

      if (response.ok) {
        const json = await response.json();
        const text = json?.candidates?.[0]?.content?.parts?.map((part) => part?.text || "").join("") || "";
        try {
          return extractJsonPayload(text);
        } catch (parseError) {
          nonFatalErrors.push(`[${version}/${model}] parse-failed`);
          continue;
        }
      }

      const errorBody = await response.text();
      const lowerBody = String(errorBody || "").toLowerCase();
      const modelNotFound = response.status === 404 || lowerBody.includes("not found") || lowerBody.includes("not supported for generatecontent");

      if (modelNotFound) {
        nonFatalErrors.push(`[${version}/${model}] HTTP ${response.status}`);
        continue;
      }

      let errorMessage = "Falha ao gerar quiz com a Gemini API.";
      if (response.status === 400) {
        errorMessage = "Requisição inválida para a Gemini API. Revise o prompt e os campos.";
      } else if (response.status === 401 || response.status === 403) {
        errorMessage = "Chave da Gemini inválida ou sem permissão.";
      } else if (response.status === 429) {
        errorMessage = "Limite de uso da Gemini atingido no momento.";
      }

      throw new Error(`${errorMessage} [HTTP ${response.status}] ${errorBody.slice(0, 220)}`);
    }
  }

  const fallbackInfo = nonFatalErrors.length
    ? ` Modelos testados sem sucesso: ${nonFatalErrors.slice(0, 5).join(" | ")}.`
    : "";
  throw new Error(`Falha ao gerar quiz: nenhum modelo Gemini compatível foi encontrado para sua chave/projeto.${fallbackInfo}`);
}

async function requestQuizFromGeminiInBatches({
  apiKey,
  prompt,
  questionCount,
  difficulty,
  schoolLevel,
  duration,
  maxAttempts,
  alternativesVisibilityMode
}) {
  const safeCount = Math.min(20, Math.max(1, Number(questionCount) || 1));
  const batchSize = safeCount > 10 ? 6 : safeCount;
  const totalBatches = Math.ceil(safeCount / batchSize);
  const normalizedQuizzes = [];

  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex += 1) {
    const remaining = safeCount - (batchIndex * batchSize);
    const currentBatchCount = Math.min(batchSize, remaining);
    const batchPrompt = buildGeminiGenerationPrompt({
      prompt: `${prompt}\n\nLote ${batchIndex + 1}/${totalBatches}. Gere perguntas diferentes das anteriores e sem repetir enunciados.`,
      questionCount: currentBatchCount,
      difficulty,
      schoolLevel,
      duration,
      maxAttempts,
      alternativesVisibilityMode
    });

    const rawBatchQuiz = await requestQuizFromGemini({
      apiKey,
      prompt: batchPrompt
    });

    const normalizedBatchQuiz = normalizeGeneratedQuizPayload(rawBatchQuiz, {
      prompt,
      duration,
      maxAttempts,
      alternativesVisibilityMode
    });

    normalizedQuizzes.push(normalizedBatchQuiz);
  }

  const mergedQuestions = normalizedQuizzes
    .flatMap((quiz) => Array.isArray(quiz.questions) ? quiz.questions : [])
    .slice(0, safeCount)
    .map((question, index) => ({ ...question, id: `q${index + 1}` }));

  if (!mergedQuestions.length) {
    throw new Error("A IA nao retornou questoes validas.");
  }

  const firstQuiz = normalizedQuizzes[0] || {};
  return {
    title: String(firstQuiz.title || "Quiz gerado com IA").trim(),
    description: String(firstQuiz.description || "Quiz gerado com IA.").trim(),
    duration: normalizeAiDuration(firstQuiz.duration || duration),
    maxAttempts: Math.max(1, Number.parseInt(firstQuiz.maxAttempts || maxAttempts, 10) || 1),
    alternativesVisibilityMode: firstQuiz.alternativesVisibilityMode === "hidden" ? "hidden" : "revealed",
    questions: mergedQuestions
  };
}

function normalizeGeneratedQuizPayload(rawQuiz, fallback = {}) {
  const generatedQuestions = Array.isArray(rawQuiz?.questions) ? rawQuiz.questions : [];
  const questions = generatedQuestions
    .map((question, index) => normalizeGeneratedQuestionSeed(question, index))
    .filter((question) => question.title && Array.isArray(question.options) && question.options.length >= 2);

  if (!questions.length) {
    throw new Error("A IA nao retornou questoes validas.");
  }

  return {
    title: String(rawQuiz?.title || "Quiz gerado com IA").trim(),
    description: String(
      rawQuiz?.description || (fallback.prompt ? `Quiz gerado a partir do prompt: ${fallback.prompt}` : "Quiz gerado com IA.")
    ).trim(),
    duration: normalizeAiDuration(rawQuiz?.duration || fallback.duration),
    maxAttempts: Math.max(1, Number.parseInt(rawQuiz?.maxAttempts || fallback.maxAttempts, 10) || 1),
    alternativesVisibilityMode:
      rawQuiz?.alternativesVisibilityMode === "hidden" || fallback.alternativesVisibilityMode === "hidden"
        ? "hidden"
        : "revealed",
    questions
  };
}

function applyGeneratedQuizToBuilder(generatedQuiz = {}) {
  if (!builderQuestions) {
    return;
  }

  state.editingQuizId = null;
  if (builderSubmitButton) {
    builderSubmitButton.textContent = "Salvar no banco";
  }

  if (builderTitleInput) {
    builderTitleInput.value = String(generatedQuiz.title || "").trim();
  }
  if (builderDescriptionInput) {
    builderDescriptionInput.value = String(generatedQuiz.description || "").trim();
  }
  if (builderMaxAttemptsInput) {
    builderMaxAttemptsInput.value = String(Math.max(1, Number.parseInt(generatedQuiz.maxAttempts, 10) || 1));
  }
  setBuilderSelectValue(builderDurationInput, String(generatedQuiz.duration || "").trim(), "8 min");
  setBuilderSelectValue(
    builderAlternativesVisibilityInput,
    generatedQuiz.alternativesVisibilityMode === "hidden" ? "hidden" : "revealed",
    "revealed"
  );

  builderQuestions.innerHTML = "";
  const generatedQuestions = Array.isArray(generatedQuiz.questions) ? generatedQuiz.questions : [];
  generatedQuestions
    .map((question, index) => normalizeGeneratedQuestionSeed(question, index))
    .filter((question) => question.title && Array.isArray(question.options) && question.options.length >= 2)
    .forEach((question) => addBuilderQuestionCard(question));

  state.builderDraftGeneratedWithAi = true;
  state.builderDraftAiPrompt = String(builderAiPromptInput?.value || "").trim();

  if (!builderQuestions.children.length) {
    addBuilderQuestionCard();
  }

  if (builderMessage) {
    builderMessage.textContent = "Rascunho gerado com IA. Revise antes de salvar.";
  }
}

async function handleBuilderGenerateWithAi() {
  if (!state.isTeacher) {
    if (builderAiMessage) {
      builderAiMessage.textContent = "Apenas professores podem gerar quizzes com IA.";
    }
    return;
  }

  const prompt = builderAiPromptInput?.value.trim() || "";
  if (!prompt) {
    if (builderAiMessage) {
      builderAiMessage.textContent = "Descreva o tema antes de gerar o quiz.";
    }
    return;
  }

  const apiKey = String(builderAiApiKeyInput?.value || getStoredGeminiApiKey() || "").trim();
  if (!apiKey) {
    if (builderAiMessage) {
      builderAiMessage.textContent = "Informe sua Gemini API key para gerar o quiz.";
    }
    return;
  }

  setStoredGeminiApiKey(apiKey);

  const requestedQuestionCount = Math.min(20, Math.max(1, Number.parseInt(builderAiQuestionCountInput?.value || "5", 10) || 5));
  const difficulty = normalizeAiDifficulty(String(builderAiDifficultyInput?.value || "facil"));
  const schoolLevel = String(builderAiSchoolLevelInput?.value || "").trim();
  const duration = normalizeAiDuration(builderDurationInput?.value || "8 min");
  const maxAttempts = Math.max(1, Number.parseInt(builderMaxAttemptsInput?.value || "1", 10) || 1);
  const alternativesVisibilityMode = builderAlternativesVisibilityInput?.value === "hidden" ? "hidden" : "revealed";

  if (builderAiGenerateButton) {
    builderAiGenerateButton.disabled = true;
    builderAiGenerateButton.textContent = "Gerando...";
  }
  if (builderAiMessage) {
    builderAiMessage.textContent = "Gerando questionário com IA...";
  }

  try {
    let generatedQuiz;
    if (requestedQuestionCount > 10) {
      if (builderAiMessage) {
        builderAiMessage.textContent = `Gerando em lotes (${requestedQuestionCount} questões) para evitar truncamento...`;
      }

      generatedQuiz = await requestQuizFromGeminiInBatches({
        apiKey,
        prompt,
        questionCount: requestedQuestionCount,
        difficulty,
        schoolLevel,
        duration,
        maxAttempts,
        alternativesVisibilityMode
      });
    } else {
      const generationPrompt = buildGeminiGenerationPrompt({
        prompt,
        questionCount: requestedQuestionCount,
        difficulty,
        schoolLevel,
        duration,
        maxAttempts,
        alternativesVisibilityMode
      });

      const rawQuiz = await requestQuizFromGemini({
        apiKey,
        prompt: generationPrompt
      });

      generatedQuiz = normalizeGeneratedQuizPayload(rawQuiz, {
        prompt,
        duration,
        maxAttempts,
        alternativesVisibilityMode
      });
    }

    state.builderDraftGeneratedWithAi = true;
    state.builderDraftAiPrompt = prompt;
    applyGeneratedQuizToBuilder(generatedQuiz);
    if (builderAiMessage) {
      builderAiMessage.textContent = "Quiz gerado com sucesso. Revise e ajuste o conteúdo antes de salvar.";
    }
    showBottomNotice("Quiz gerado com IA e carregado no criador.", "success", 2600);
  } catch (error) {
    console.error("Erro ao gerar quiz com IA:", error);
    if (builderAiMessage) {
      builderAiMessage.textContent = error?.message || "Não foi possível gerar o quiz com IA agora.";
    }
    showBottomNotice("Falha ao gerar quiz com IA.", "error", 3200);
  } finally {
    if (builderAiGenerateButton) {
      builderAiGenerateButton.disabled = false;
      builderAiGenerateButton.textContent = "Gerar com IA";
    }
  }
}

function updateBuilderQuestionTitles() {
  getBuilderQuestionCards().forEach((card, index) => {
    const title = card.querySelector(".builder-question-head h4");
    if (title) {
      title.textContent = `Questão ${index + 1}`;
    }
  });
}

function resetBuilderForm() {
  state.editingQuizId = null;
  state.builderDraftGeneratedWithAi = false;
  state.builderDraftAiPrompt = "";
  if (builderForm) {
    builderForm.reset();
  }
  if (builderQuestions) {
    builderQuestions.innerHTML = "";
  }
  if (builderSubmitButton) {
    builderSubmitButton.textContent = "Salvar no banco";
  }
  if (builderAiMessage) {
    builderAiMessage.textContent = "";
  }
  addBuilderQuestionCard();
}

async function startEditingQuiz(quizId) {
  const quiz = getAllQuizzes().find((item) => item.id === quizId);
  if (!quiz || !builderQuestions) {
    return;
  }

  const answerKey = await getQuizAnswerKeySecure(quiz.id);

  state.editingQuizId = quiz.id;
  openBuilderScreen(false);
  builderTitleInput.value = quiz.title || "";
  builderDescriptionInput.value = quiz.description || "";
  builderDurationInput.value = quiz.duration || "8 min";
  if (builderMaxAttemptsInput) {
    builderMaxAttemptsInput.value = String(Math.max(1, Number.parseInt(quiz.maxAttempts, 10) || 1));
  }
  if (builderAlternativesVisibilityInput) {
    builderAlternativesVisibilityInput.value = quiz.alternativesVisibilityMode === "hidden" ? "hidden" : "revealed";
  }
  if (builderAiPromptInput) {
    builderAiPromptInput.value = String(quiz.aiGenerationPrompt || "");
  }
  state.builderDraftGeneratedWithAi = Boolean(String(quiz.aiGenerationPrompt || "").trim());
  state.builderDraftAiPrompt = String(quiz.aiGenerationPrompt || "").trim();
  builderQuestions.innerHTML = "";
  (quiz.questions || []).forEach((question, index) => {
    const questionId = question.id || `q${index + 1}`;
    const decodedAnswer = decodeAnswerToken(quiz.id, questionId, question.answerToken);
    addBuilderQuestionCard({
      ...question,
      correctAnswer: String(answerKey[questionId] || decodedAnswer || "")
    });
  });
  if (builderSubmitButton) {
    builderSubmitButton.textContent = "Salvar alteracoes";
  }
  if (builderMessage) {
    builderMessage.textContent = "Editando quiz existente.";
  }
}

function buildQuestionsFromBuilder() {
  const questions = [];
  const answerKey = {};

  for (const card of getBuilderQuestionCards()) {
    const title = card.querySelector(".builder-question-title")?.value.trim() || "";
    const description = card.querySelector(".builder-question-description")?.value.trim() || "";
    const explanation = card.querySelector(".builder-question-explanation")?.value.trim() || "";
    const image = card.querySelector(".builder-image-url")?.value.trim() || "";
    const points = Number(card.querySelector(".builder-question-points")?.value || 1);
    const optionRows = Array.from(card.querySelectorAll(".builder-option-row"));
    const timer = Number(card.querySelector(".builder-question-timer")?.value || 0);

    if (!title || optionRows.length < 2) {
      return { error: "Cada questão precisa de enunciado e pelo menos 2 alternativas." };
    }

    const options = [];
    let correctAnswer = "";

    optionRows.forEach((row, index) => {
      const label = row.querySelector(".builder-option-label-input")?.value.trim() || "";
      const isCorrect = row.querySelector("input[type='radio']")?.checked;
      const value = String.fromCharCode(97 + index);

      if (label) {
        options.push({ value, label });
        if (isCorrect) {
          correctAnswer = value;
        }
      }
    });

    if (options.length < 2 || !correctAnswer) {
      return { error: "Cada questão precisa de pelo menos 2 alternativas preenchidas e uma correta marcada." };
    }

    const questionId = `q${questions.length + 1}`;

    questions.push({
      id: questionId,
      type: "single-choice",
      title,
      description: description || "Selecione apenas uma alternativa.",
      explanation,
      points: Number.isFinite(points) && points > 0 ? points : 1,
      options,
      timer: Number.isFinite(timer) && timer > 0 ? timer : 0,
      image: image || null,
      _correctAnswer: correctAnswer
    });

    answerKey[questionId] = correctAnswer;
  }

  if (questions.length === 0) {
    return { error: "Adicione ao menos uma questão." };
  }

  return { questions, answerKey };
}

function openBuilderScreen(shouldReset = true) {
  showOnlyScreen("builder");
  if (teacherScreen) {
    teacherScreen.classList.add("hidden");
  }
  if (teacherClassroomsScreen) {
    teacherClassroomsScreen.classList.add("hidden");
  }
  if (homeScreen) {
    homeScreen.classList.add("hidden");
  }
  if (builderMessage) {
    builderMessage.textContent = "";
  }
  if (shouldReset) {
    resetBuilderForm();
  }
}

async function handleBuilderCreateQuiz(event) {
  event.preventDefault();

  if (!state.isTeacher) {
    return;
  }

  if (!window.firebaseDB || !window.firebaseSetDoc || !window.firebaseDoc) {
    if (builderMessage) {
      builderMessage.textContent = "Firebase indisponível para criar quiz.";
    }
    return;
  }

  const title = builderTitleInput?.value.trim() || "";
  const description = builderDescriptionInput?.value.trim() || "";
  const duration = builderDurationInput?.value.trim() || "10 min";
  const maxAttempts = Math.max(1, Number.parseInt(builderMaxAttemptsInput?.value || "1", 10) || 1);
  const alternativesVisibilityMode = builderAlternativesVisibilityInput?.value === "hidden" ? "hidden" : "revealed";

  if (!title || !description) {
    if (builderMessage) {
      builderMessage.textContent = "Preencha título e descrição.";
    }
    return;
  }

  const built = buildQuestionsFromBuilder();
  if (built.error) {
    if (builderMessage) {
      builderMessage.textContent = built.error;
    }
    return;
  }

  const slugBase = normalizeNameSlug(title);
  const quizId = state.editingQuizId || `custom-${slugBase}-${Date.now()}`;
  const questionsWithTokens = built.questions.map((question, index) => {
    const questionId = String(question.id || `q${index + 1}`);
    const correctAnswer = String(question._correctAnswer || "");

    return {
      id: questionId,
      type: "single-choice",
      title: question.title,
      description: question.description,
      explanation: String(question.explanation || ""),
      points: question.points,
      options: question.options,
      timer: question.timer,
      image: question.image || null,
      answerToken: buildAnswerToken(quizId, questionId, correctAnswer)
    };
  });
  const questionsRevision = new Date().toISOString();

  const payload = {
    title,
    description,
    duration,
    maxAttempts,
    alternativesVisibilityMode,
    questionCount: questionsWithTokens.length,
    questionsRevision,
    ativo: true,
    criadoPorUid: state.user?.uid || "",
    criadoPorEmail: state.user?.email || "",
    criadoEmIso: new Date().toISOString(),
    atualizadoEmIso: new Date().toISOString()
  };

  const aiPromptToPersist = state.builderDraftGeneratedWithAi
    ? String(state.builderDraftAiPrompt || builderAiPromptInput?.value || "").trim()
    : "";

  if (aiPromptToPersist) {
    payload.aiGenerationPrompt = aiPromptToPersist;
    payload.aiGenerationProvider = "gemini";
  } else {
    payload.aiGenerationPrompt = "";
  }

  if (builderMessage) {
    builderMessage.textContent = state.editingQuizId ? "Salvando alteracoes..." : "Salvando no banco de quizzes...";
  }

  try {
    const quizRef = window.firebaseDoc(window.firebaseDB, "quizzes_custom", quizId);
    const answersRef = window.firebaseDoc(window.firebaseDB, "quizzes_answer_keys", quizId);
    const quizQuestionsPromise = replaceQuizQuestionsInFirestore(quizId, questionsWithTokens, questionsRevision);
    await Promise.all([
      window.firebaseSetDoc(quizRef, payload),
      window.firebaseSetDoc(answersRef, {
        quizId,
        answers: built.answerKey,
        atualizadoEmIso: new Date().toISOString(),
        atualizadoPorUid: state.user?.uid || "",
        atualizadoPorEmail: state.user?.email || ""
      }, { merge: true }),
      quizQuestionsPromise
    ]);
    if (builderMessage) {
      builderMessage.textContent = state.editingQuizId ? "Quiz atualizado com sucesso." : "Quiz salvo no banco com sucesso. Agora poste para uma turma na home.";
    }
    await refreshCustomQuizzes();
    resetBuilderForm();
  } catch (error) {
    if (builderMessage) {
      builderMessage.textContent = state.editingQuizId ? "Erro ao atualizar quiz." : "Erro ao salvar quiz no banco.";
    }
    console.error("Erro ao criar quiz customizado:", error);
  }
}

function getSelectedQuiz() {
  return getAllQuizzes().find((quiz) => quiz.id === state.selectedQuizId) || null;
}

function getStorageKey(quizId) {
  const uid = state.user?.uid || "anon";
  return `quiz_attempt_${uid}_${quizId}`;
}

function getAttemptCountStorageKey(quizId) {
  const uid = state.user?.uid || "anon";
  return `quiz_attempt_count_${uid}_${quizId}`;
}

function getQuizAccessLockStorageKey(quizId) {
  const uid = state.user?.uid || "anon";
  return `quiz_access_lock_${uid}_${quizId}`;
}

function getStoredAttempt(quizId) {
  const raw = localStorage.getItem(getStorageKey(quizId));
  return raw ? JSON.parse(raw) : null;
}

function getStoredAttemptCount(quizId) {
  const raw = localStorage.getItem(getAttemptCountStorageKey(quizId));
  const parsed = Number.parseInt(String(raw || "0"), 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
}

function setStoredAttemptCount(quizId, value) {
  const normalized = Math.max(0, Number.parseInt(String(value || 0), 10) || 0);
  if (normalized === 0) {
    localStorage.removeItem(getAttemptCountStorageKey(quizId));
    return;
  }
  localStorage.setItem(getAttemptCountStorageKey(quizId), String(normalized));
}

function isQuizAccessLocked(quizId) {
  return localStorage.getItem(getQuizAccessLockStorageKey(quizId)) === "1";
}

function lockQuizAccess(quizId) {
  if (!quizId) {
    return;
  }

  localStorage.setItem(getQuizAccessLockStorageKey(quizId), "1");
}

function clearQuizAccessLock(quizId) {
  if (!quizId) {
    return;
  }

  localStorage.removeItem(getQuizAccessLockStorageKey(quizId));
}

function getKnownAttempt(quizId) {
  // Sempre prioriza o estado remoto, se carregado
  if (state.remoteAttemptsLoaded) {
    // Se não existe tentativa remota, remove localStorage para garantir sincronização
    if (!state.remoteAttemptsByQuiz[quizId]) {
      localStorage.removeItem(getStorageKey(quizId));
      setStoredAttemptCount(quizId, Number(state.remoteAttemptCountsByQuiz?.[quizId] || 0));
      return null;
    }
    return state.remoteAttemptsByQuiz[quizId];
  }
  // Se não carregou remoto ainda, usa local
  const localAttempt = getStoredAttempt(quizId);
  return localAttempt || null;
}

function resolveQuizMaxAttempts(quiz) {
  return Math.max(1, Number.parseInt(quiz?.maxAttempts, 10) || 1);
}

function getKnownAttemptCount(quizId) {
  if (state.remoteAttemptsLoaded) {
    return Number(state.remoteAttemptCountsByQuiz?.[quizId] || 0);
  }

  const localCount = getStoredAttemptCount(quizId);
  if (localCount > 0) {
    return localCount;
  }

  return getStoredAttempt(quizId) ? 1 : 0;
}

function incrementKnownAttemptCount(quizId) {
  if (!quizId) {
    return;
  }

  const localCurrent = getStoredAttemptCount(quizId);
  setStoredAttemptCount(quizId, localCurrent + 1);

  if (!state.remoteAttemptsLoaded) {
    return;
  }

  const current = Number(state.remoteAttemptCountsByQuiz?.[quizId] || 0);
  state.remoteAttemptCountsByQuiz[quizId] = current + 1;
}

function formatAttemptsRemainingMessage(quiz, usedAttempts = 0) {
  const maxAttempts = resolveQuizMaxAttempts(quiz);
  const remaining = Math.max(0, maxAttempts - usedAttempts);
  return `${usedAttempts}/${maxAttempts} tentativas usadas (${remaining} restantes).`;
}

function saveStoredAttempt(quizId, payload) {
  localStorage.setItem(getStorageKey(quizId), JSON.stringify(payload));
  if (state.remoteAttemptsLoaded) {
    state.remoteAttemptsByQuiz[quizId] = payload;
  }
}

function syncAttemptCacheForQuiz(quizId, payload, attemptsUsed) {
  if (!quizId) {
    return;
  }

  saveStoredAttempt(quizId, payload);
  setStoredAttemptCount(quizId, attemptsUsed);
  if (state.remoteAttemptsLoaded) {
    state.remoteAttemptsByQuiz[quizId] = payload;
    state.remoteAttemptCountsByQuiz[quizId] = attemptsUsed;
  }
}

function cloneAnswersSnapshot() {
  return { ...state.answers };
}

function buildAttemptResult(quiz, answersSnapshot = cloneAnswersSnapshot(), at = new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" })) {
  const maxPoints = (quiz.questions || []).reduce((sum, question) => {
    const points = Number(question?.points || 1);
    return sum + (Number.isFinite(points) && points > 0 ? points : 1);
  }, 0);

  return {
    quizId: quiz.id,
    quizTitle: quiz.title,
    studentName: state.studentName,
    earnedPoints: 0,
    maxPoints,
    percent: 0,
    date: at
  };
}

function buildCancelledAttemptResult(quiz, answersSnapshot = cloneAnswersSnapshot(), at = new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" })) {
  const sourceQuestions = Array.isArray(state.activeQuestions) && state.activeQuestions.length
    ? state.activeQuestions
    : (Array.isArray(quiz?.questions) ? quiz.questions : []);

  const questionResults = buildCancelledQuestionResults(quiz, answersSnapshot);
  const pointsByQuestionId = {};
  let maxPoints = 0;

  sourceQuestions.forEach((question, index) => {
    const questionId = String(question?.id || `q${index + 1}`);
    const points = Number(question?.points || 1);
    const normalizedPoints = Number.isFinite(points) && points > 0 ? points : 1;
    pointsByQuestionId[questionId] = normalizedPoints;
    maxPoints += normalizedPoints;
  });

  const earnedPoints = questionResults.reduce((sum, item) => {
    if (!item?.isCorrect) {
      return sum;
    }
    const questionPoints = Number(pointsByQuestionId[String(item.questionId || "")] || 1);
    return sum + (Number.isFinite(questionPoints) && questionPoints > 0 ? questionPoints : 1);
  }, 0);

  const percent = maxPoints > 0 ? Math.round((earnedPoints / maxPoints) * 100) : 0;
  const durationSeconds = computeAttemptDurationSeconds(state.attemptStartedAtMs, Date.now());

  return {
    result: {
      quizId: quiz?.id,
      quizTitle: quiz?.title,
      studentName: state.studentName,
      earnedPoints,
      maxPoints,
      percent,
      date: at,
      startedAt: state.attemptStartedAtLabel || "",
      endedAt: at,
      durationSeconds,
      durationLabel: formatDurationFromSeconds(durationSeconds),
      questionResults
    },
    questionResults
  };
}

function buildCancelledQuestionResults(quiz, answersSnapshot = cloneAnswersSnapshot()) {
  const sourceQuestions = Array.isArray(state.activeQuestions) && state.activeQuestions.length
    ? state.activeQuestions
    : (Array.isArray(quiz?.questions) ? quiz.questions : []);

  return sourceQuestions.map((question, index) => {
    const questionId = String(question?.id || `q${index + 1}`);
    const selectedValue = String(answersSnapshot?.[questionId] || "");
    const options = Array.isArray(question?.options) ? question.options : [];
    const selectedOption = options.find((option) => String(option?.value || "") === selectedValue);
    const correctValue = resolveQuestionCorrectValue(quiz?.id, question, index);
    const correctOption = options.find((option) => String(option?.value || "") === correctValue);

    return {
      questionId,
      questionTitle: question?.title || `Questão ${index + 1}`,
      questionExplanation: String(question?.explanation || ""),
      selectedValue,
      selectedLabel: selectedOption ? String(selectedOption.label || "") : "",
      correctValue,
      correctLabel: correctOption ? String(correctOption.label || "") : "",
      isCorrect: Boolean(selectedValue) && Boolean(correctValue) && selectedValue === correctValue,
      options: options.map((option) => ({
        value: String(option?.value || ""),
        label: String(option?.label || "")
      }))
    };
  });
}

function getAttemptForSelectedQuiz() {
  const quiz = getSelectedQuiz();
  if (!quiz) {
    return null;
  }
  return getKnownAttempt(quiz.id);
}

async function syncLatestAttemptForQuiz(quizId) {
  if (!quizId) {
    return null;
  }

  const remoteAttempt = await getRemoteAttemptForQuiz(quizId);
  if (remoteAttempt) {
    saveStoredAttempt(quizId, remoteAttempt);
    return remoteAttempt;
  }

  return getKnownAttempt(quizId);
}

function getPendingFocusWaitContext() {
  if (state.isTeacher) {
    return null;
  }

  const quizzes = getAllQuizzes();
  let selected = null;

  quizzes.forEach((quiz) => {
    const attempt = getKnownAttempt(quiz.id);
    if (!attempt || attempt.status !== "completed" || !attempt.result) {
      return;
    }
    if (quiz.allowFocusExitAfterFinish) {
      return;
    }
    if (attempt.focusExitPenaltyApplied || attempt?.result?.focusExitPenaltyApplied) {
      return;
    }

    const ts = parsePtBrDate(attempt.result?.date || "");
    if (!selected || ts >= selected.ts) {
      selected = { quiz, attempt, ts };
    }
  });

  return selected;
}

function isFocusWaitModeActive() {
  return Boolean(getPendingFocusWaitContext());
}

function isFocusWaitGraceActive() {
  return Number(focusWaitGraceDeadlineMs || 0) > Date.now();
}

function isFocusWaitRecoverySatisfied() {
  return document.visibilityState === "visible" && document.hasFocus() && isFullscreenActive();
}

function ensureFocusWaitAlertElement() {
  let box = document.getElementById("focus-wait-alert");
  if (box) {
    return box;
  }

  box = document.createElement("section");
  box.id = "focus-wait-alert";
  box.className = "hidden";
  box.innerHTML = `
    <div class="focus-wait-alert-card">
      <p>${FOCUS_WAIT_PROCESS_NAME}</p>
      <p id="focus-wait-alert-text"></p>
      <button type="button" id="focus-wait-back-btn" class="btn btn-primary">
        Voltar para tela cheia
        <small id="focus-wait-alert-count" class="focus-wait-alert-counter"></small>
      </button>
    </div>
  `;

  document.body.appendChild(box);

  const backButton = document.getElementById("focus-wait-back-btn");
  if (backButton) {
    backButton.addEventListener("click", async () => {
      await requestFullscreenMode();
      if (isFocusWaitRecoverySatisfied()) {
        resolveFocusWaitGraceWithoutPenalty();
      }
    });
  }

  return box;
}

function ensureFocusWaitPersistentNoticeElement() {
  let notice = document.getElementById("focus-wait-persistent-notice");
  if (notice) {
    return notice;
  }

  notice = document.createElement("div");
  notice.id = "focus-wait-persistent-notice";
  notice.className = "focus-wait-persistent-notice hidden";
  notice.setAttribute("role", "status");
  notice.setAttribute("aria-live", "polite");
  notice.innerHTML = `
    <strong>${FOCUS_WAIT_PROCESS_NAME}</strong>
    <span>Modo de espera ativo. Nao feche, nao recarregue e nao saia da tela. Aguarde a liberacao do professor.</span>
  `;

  document.body.appendChild(notice);
  return notice;
}

function renderFocusWaitPersistentNotice() {
  const notice = ensureFocusWaitPersistentNoticeElement();
  const isActive = !state.isTeacher && isFocusWaitModeActive();

  if (isActive) {
    notice.classList.remove("hidden");
    return;
  }

  notice.classList.add("hidden");
}

function hideFocusWaitAlert() {
  const box = document.getElementById("focus-wait-alert");
  if (box) {
    box.classList.add("hidden");
  }
}

function updateFocusWaitAlert() {
  const box = ensureFocusWaitAlertElement();
  const text = document.getElementById("focus-wait-alert-text");
  const count = document.getElementById("focus-wait-alert-count");
  const remainingMs = Math.max(0, Number(focusWaitGraceDeadlineMs || 0) - Date.now());
  const remainingSeconds = Math.ceil(remainingMs / 1000);

  if (text) {
    text.textContent = `${focusWaitGraceReason} Retorne ao foco em tela cheia para evitar penalidade.`;
  }
  if (count) {
    count.textContent = `${Math.max(0, remainingSeconds)}s`;
  }
  box.classList.remove("hidden");
}

function clearFocusWaitGraceTimers() {
  if (focusWaitGraceTimeoutId) {
    clearTimeout(focusWaitGraceTimeoutId);
    focusWaitGraceTimeoutId = null;
  }
  if (focusWaitGraceIntervalId) {
    clearInterval(focusWaitGraceIntervalId);
    focusWaitGraceIntervalId = null;
  }
}

function clearFocusWaitGrace() {
  clearFocusWaitGraceTimers();
  focusWaitGraceDeadlineMs = 0;
  focusWaitGraceReason = "";
  hideFocusWaitAlert();
}

function resolveFocusWaitGraceWithoutPenalty() {
  clearFocusWaitGrace();
  showBottomNotice("Aviso removido: você voltou ao foco em tela cheia dentro do tempo.", "success", 3000);
}

function normalizeLobbySegment(value, fallback = "sala") {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 48) || fallback;
}

function buildWaitClassLobbyKey(quiz) {
  const classSeed = String(quiz?.postedToClassroomId || state.classroomId || quiz?.postedToClassroomName || "turma");
  const quizSeed = String(quiz?.sourceQuizId || quiz?.id || "quiz");
  const classSegment = normalizeLobbySegment(classSeed, "turma");
  const quizSegment = normalizeLobbySegment(quizSeed, "quiz");
  return `${WAIT_NONOGRAM_LOBBY_PREFIX}-${classSegment}-${quizSegment}`;
}

function getWaitNonogramDateKey() {
  return new Date().toISOString().slice(0, 10);
}

function hashStringToSeed(text) {
  let hash = 2166136261;
  const source = String(text || "seed");
  for (let i = 0; i < source.length; i += 1) {
    hash ^= source.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0) || 123456789;
}

function createSeededRng(seedText) {
  let seed = hashStringToSeed(seedText);
  return () => {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return seed / 4294967296;
  };
}

function createEmptyNonogramMarks(size) {
  return Array.from({ length: size }, () => Array.from({ length: size }, () => 0));
}

const NONOGRAM_ICON_LIBRARY = [
  {
    name: "Coracao",
    rows: [
      "00100100",
      "01111110",
      "11111111",
      "11111111",
      "01111110",
      "00111100",
      "00011000",
      "00000000"
    ]
  },
  {
    name: "Estrela",
    rows: [
      "00011000",
      "01011010",
      "01111110",
      "11111111",
      "01111110",
      "01011010",
      "10000001",
      "00000000"
    ]
  },
  {
    name: "Casa",
    rows: [
      "00011000",
      "00111100",
      "01111110",
      "11011011",
      "10000001",
      "10011001",
      "11111111",
      "00000000"
    ]
  },
  {
    name: "Lampada",
    rows: [
      "00111100",
      "01111110",
      "01111110",
      "00111100",
      "00011000",
      "00011000",
      "00111100",
      "00111100"
    ]
  },
  {
    name: "Trofeu",
    rows: [
      "11111111",
      "01111110",
      "11011011",
      "00111100",
      "00011000",
      "00111100",
      "00111100",
      "01111110"
    ]
  },
  {
    name: "Raio",
    rows: [
      "00011000",
      "00110000",
      "01100000",
      "11111100",
      "00011000",
      "00110000",
      "01100000",
      "11000000"
    ]
  },
  {
    name: "Smiley",
    rows: [
      "00111100",
      "01000010",
      "10100101",
      "10000001",
      "10100101",
      "10011001",
      "01000010",
      "00111100"
    ]
  },
  {
    name: "Livro",
    rows: [
      "11000011",
      "11100111",
      "11100111",
      "11011011",
      "11011011",
      "11100111",
      "11100111",
      "11000011"
    ]
  }
];

function iconRowsToMatrix(rows, size) {
  return rows.slice(0, size).map((rowText) => {
    const normalized = String(rowText || "").padEnd(size, "0").slice(0, size);
    return normalized.split("").map((char) => (char === "1" ? 1 : 0));
  });
}

function iconMatrixToRows(matrix, size) {
  return Array.from({ length: size }, (_, row) => {
    const rowValues = Array.isArray(matrix?.[row]) ? matrix[row] : [];
    return Array.from({ length: size }, (_, col) => (Number(rowValues[col] || 0) ? "1" : "0")).join("");
  });
}

function normalizeNonogramTemplate(rawTemplate, size) {
  if (!rawTemplate || !Array.isArray(rawTemplate.rows)) {
    return null;
  }

  const rows = rawTemplate.rows
    .slice(0, size)
    .map((rowText) => String(rowText || "").replace(/[^01]/g, "").padEnd(size, "0").slice(0, size));

  if (rows.length !== size) {
    return null;
  }

  const filled = rows.join("").split("").reduce((acc, char) => acc + (char === "1" ? 1 : 0), 0);
  if (filled === 0) {
    return null;
  }

  return {
    id: String(rawTemplate.id || "").trim() || `tpl_${Date.now().toString(36)}`,
    name: String(rawTemplate.name || "Desenho").trim() || "Desenho",
    rows,
    createdAtMs: Number(rawTemplate.createdAtMs || Date.now())
  };
}

function getCachedNonogramTemplatesByQuiz(quizId) {
  const key = NONOGRAM_TEMPLATE_GLOBAL_KEY;
  return Array.isArray(state.nonogramTemplatesByQuiz[key]) ? state.nonogramTemplatesByQuiz[key] : [];
}

async function loadCustomNonogramTemplatesForQuiz(quizId, options = {}) {
  const key = NONOGRAM_TEMPLATE_GLOBAL_KEY;

  if (!options.force && Array.isArray(state.nonogramTemplatesByQuiz[key])) {
    return state.nonogramTemplatesByQuiz[key];
  }

  if (!window.firebaseDB || !window.firebaseDoc || !window.firebaseGetDoc) {
    if (!Array.isArray(state.nonogramTemplatesByQuiz[key])) {
      state.nonogramTemplatesByQuiz[key] = [];
    }
    return state.nonogramTemplatesByQuiz[key];
  }

  try {
    const ref = window.firebaseDoc(window.firebaseDB, "nonogram_templates", NONOGRAM_TEMPLATE_GLOBAL_KEY);
    const snap = await window.firebaseGetDoc(ref);
    if (!snap?.exists?.()) {
      state.nonogramTemplatesByQuiz[key] = [];
      return [];
    }

    const data = snap.data() || {};
    const normalized = (Array.isArray(data.templates) ? data.templates : [])
      .map((template) => normalizeNonogramTemplate(template, WAIT_NONOGRAM_GRID_SIZE))
      .filter(Boolean);

    state.nonogramTemplatesByQuiz[key] = normalized;
    return normalized;
  } catch (error) {
    console.error("Erro ao carregar templates de nonograma:", error);
    if (!Array.isArray(state.nonogramTemplatesByQuiz[key])) {
      state.nonogramTemplatesByQuiz[key] = [];
    }
    return state.nonogramTemplatesByQuiz[key];
  }
}

async function saveCustomNonogramTemplateForQuiz(quizId, template) {
  const key = NONOGRAM_TEMPLATE_GLOBAL_KEY;
  if (!key || !template || !window.firebaseDB || !window.firebaseDoc || !window.firebaseSetDoc) {
    return { ok: false, reason: "unavailable" };
  }

  const existing = await loadCustomNonogramTemplatesForQuiz(key);
  const sameShape = existing.some((item) => item.rows.join("") === template.rows.join(""));
  if (sameShape) {
    return { ok: false, reason: "duplicate" };
  }

  const next = [template, ...existing].slice(0, 60);

  try {
    const ref = window.firebaseDoc(window.firebaseDB, "nonogram_templates", NONOGRAM_TEMPLATE_GLOBAL_KEY);
    await window.firebaseSetDoc(ref, {
      scope: "global",
      templates: next.map((item) => ({
        id: item.id,
        name: item.name,
        rows: item.rows,
        createdAtMs: item.createdAtMs
      })),
      updatedAtMs: Date.now(),
      updatedBy: state.profileSlug || state.user?.uid || ""
    }, { merge: true });
    state.nonogramTemplatesByQuiz[key] = next;
    return { ok: true, reason: "saved" };
  } catch (error) {
    console.error("Erro ao salvar template de nonograma:", error);
    const code = String(error?.code || "").toLowerCase();
    if (code.includes("permission-denied")) {
      // Keep a local session cache so the teacher can continue creating even before rules are deployed.
      state.nonogramTemplatesByQuiz[key] = next;
      return { ok: true, reason: "local-only" };
    }
    return { ok: false, reason: "save-error" };
  }
}

function getActiveNonogramLibrary(quizId) {
  const custom = getCachedNonogramTemplatesByQuiz(quizId);
  return custom.length ? custom : NONOGRAM_ICON_LIBRARY;
}

function buildNonogramPuzzle(size, rng, library = NONOGRAM_ICON_LIBRARY) {
  const sourceLibrary = Array.isArray(library) && library.length ? library : NONOGRAM_ICON_LIBRARY;

  if (size === 8 && Array.isArray(sourceLibrary) && sourceLibrary.length) {
    const index = Math.floor(rng() * sourceLibrary.length) % sourceLibrary.length;
    const icon = sourceLibrary[index];
    const matrix = iconRowsToMatrix(icon.rows, size);

    if (matrix.length === size && matrix.every((row) => row.length === size)) {
      return {
        matrix,
        name: String(icon?.name || "Desenho").trim() || "Desenho"
      };
    }
  }

  // Fallback for unexpected sizes.
  return {
    matrix: Array.from({ length: size }, () => Array.from({ length: size }, () => (rng() > 0.56 ? 1 : 0))),
    name: "Desenho surpresa"
  };
}

function getNonogramDrawingSuggestionsText() {
  const libraryNames = NONOGRAM_ICON_LIBRARY
    .map((item) => String(item?.name || "").trim())
    .filter(Boolean);
  const names = [...new Set([...libraryNames, ...NONOGRAM_DRAWING_SUGGESTIONS])];
  return names.slice(0, 10).join(", ");
}

function buildLineClues(line) {
  const clues = [];
  let streak = 0;
  line.forEach((value) => {
    if (value) {
      streak += 1;
    } else if (streak > 0) {
      clues.push(streak);
      streak = 0;
    }
  });
  if (streak > 0) {
    clues.push(streak);
  }
  return clues.length ? clues : [0];
}

function buildNonogramClues(solution) {
  const size = solution.length;
  const rowClues = solution.map((row) => buildLineClues(row));
  const colClues = Array.from({ length: size }, (_, col) => {
    const column = Array.from({ length: size }, (_, row) => solution[row][col]);
    return buildLineClues(column);
  });

  return { rowClues, colClues };
}

function calcWaitNonogramStats() {
  let correctFilled = 0;
  let wrongFilled = 0;
  let totalFilled = 0;

  for (let row = 0; row < waitNonogramState.size; row += 1) {
    for (let col = 0; col < waitNonogramState.size; col += 1) {
      const expected = Number(waitNonogramState.solution[row]?.[col] || 0);
      const mark = Number(waitNonogramState.marks[row]?.[col] || 0);
      if (expected === 1) {
        totalFilled += 1;
      }
      if (mark === 1 && expected === 1) {
        correctFilled += 1;
      }
      if (mark === 1 && expected === 0) {
        wrongFilled += 1;
      }
    }
  }

  const progress = totalFilled > 0 ? Math.round((correctFilled / totalFilled) * 100) : 0;
  return { correctFilled, wrongFilled, totalFilled, progress };
}

function computeWaitNonogramScore({ correctFilled, errors, completed }) {
  const base = correctFilled * 100;
  const penalty = errors * 50;
  const completionBonus = completed ? 900 : 0;
  return Math.max(0, base + completionBonus - penalty);
}

function resetWaitNonogramRuntime() {
  if (waitNonogramState.autoAdvanceTimeoutId) {
    clearTimeout(waitNonogramState.autoAdvanceTimeoutId);
    waitNonogramState.autoAdvanceTimeoutId = null;
  }
  waitNonogramState.active = false;
  waitNonogramState.submitted = false;
}

function clearWaitNonogramBoard() {
  if (!waitNonogramState.active || waitNonogramState.completed || waitNonogramState.submitted) {
    return;
  }

  waitNonogramState.marks = createEmptyNonogramMarks(waitNonogramState.size);
  waitNonogramState.errors = 0;
  renderWaitNonogramBoard();
  updateWaitNonogramHud();
}

function toggleTeacherRevealMode() {
  if (!state.isTeacher) {
    return;
  }

  waitNonogramState.teacherRevealMode = !waitNonogramState.teacherRevealMode;
  const toggleButton = document.getElementById("wait-nonogram-reveal");
  if (toggleButton) {
    toggleButton.textContent = waitNonogramState.teacherRevealMode ? "Ocultar desenho (professor)" : "Revelar desenho (professor)";
  }
  renderWaitNonogramBoard();
  updateWaitNonogramHud();
}

function ensureNonogramCreatorMatrix() {
  if (!Array.isArray(waitNonogramState.creatorMatrix) || waitNonogramState.creatorMatrix.length !== waitNonogramState.size) {
    waitNonogramState.creatorMatrix = createEmptyNonogramMarks(waitNonogramState.size);
  }
}

function renderWaitNonogramCreatorGrid() {
  const grid = document.getElementById("wait-nonogram-creator-grid");
  if (!grid || !state.isTeacher) {
    return;
  }

  ensureNonogramCreatorMatrix();
  grid.innerHTML = "";
  const size = waitNonogramState.size;

  for (let row = 0; row < size; row += 1) {
    for (let col = 0; col < size; col += 1) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "wait-nonogram-creator-cell";
      if (Number(waitNonogramState.creatorMatrix[row]?.[col] || 0) === 1) {
        button.classList.add("is-on");
      }

      button.addEventListener("click", () => {
        const current = Number(waitNonogramState.creatorMatrix[row][col] || 0);
        waitNonogramState.creatorMatrix[row][col] = current === 1 ? 0 : 1;
        renderWaitNonogramCreatorGrid();
      });

      grid.appendChild(button);
    }
  }
}

function setWaitNonogramCreatorMessage(text, tone = "info") {
  const msg = document.getElementById("wait-nonogram-creator-msg");
  if (!msg) {
    return;
  }
  msg.textContent = String(text || "");
  msg.classList.remove("is-success", "is-error");
  if (tone === "success") {
    msg.classList.add("is-success");
  } else if (tone === "error") {
    msg.classList.add("is-error");
  }
}

async function saveCurrentNonogramCreatorTemplate(quiz) {
  if (!state.isTeacher || !quiz?.id) {
    return;
  }

  ensureNonogramCreatorMatrix();
  const rows = iconMatrixToRows(waitNonogramState.creatorMatrix, waitNonogramState.size);
  const filled = rows.join("").split("").reduce((acc, char) => acc + (char === "1" ? 1 : 0), 0);
  if (filled < 8) {
    setWaitNonogramCreatorMessage("Desenho muito vazio. Preencha mais células antes de salvar.", "error");
    return;
  }

  const nameInput = document.getElementById("wait-nonogram-creator-name");
  const templateName = String(nameInput?.value || waitNonogramState.creatorName || "").trim() || `Desenho ${Date.now().toString().slice(-4)}`;
  const template = normalizeNonogramTemplate({
    id: `tpl_${Date.now().toString(36)}`,
    name: templateName,
    rows,
    createdAtMs: Date.now()
  }, waitNonogramState.size);

  if (!template) {
    setWaitNonogramCreatorMessage("Não foi possível validar o desenho.", "error");
    return;
  }

  const result = await saveCustomNonogramTemplateForQuiz(quiz.id, template);
  if (!result?.ok) {
    if (result?.reason === "duplicate") {
      setWaitNonogramCreatorMessage("Esse desenho já existe para este quiz.", "error");
    } else if (result?.reason === "unavailable") {
      setWaitNonogramCreatorMessage("Salvamento indisponível agora. Verifique conexão e recarregue.", "error");
    } else {
      setWaitNonogramCreatorMessage("Não foi possível salvar no momento. Tente novamente.", "error");
    }
    return;
  }

  if (result.reason === "local-only") {
    setWaitNonogramCreatorMessage("Desenho salvo apenas nesta sessão (publique as regras do Firestore para salvar no banco).", "success");
  } else {
    setWaitNonogramCreatorMessage("Desenho salvo! Ele já pode aparecer nos próximos nonogramas da turma.", "success");
  }
}

function setHomeNonogramCreatorMessage(text, tone = "info") {
  const msg = document.getElementById("home-nonogram-creator-msg");
  if (!msg) {
    return;
  }
  msg.textContent = String(text || "");
  msg.classList.remove("is-success", "is-error");
  if (tone === "success") {
    msg.classList.add("is-success");
  } else if (tone === "error") {
    msg.classList.add("is-error");
  }
}

function renderHomeNonogramCreatorGrid() {
  const grid = document.getElementById("home-nonogram-creator-grid");
  if (!grid) {
    return;
  }

  ensureNonogramCreatorMatrix();
  grid.innerHTML = "";
  const size = waitNonogramState.size;

  for (let row = 0; row < size; row += 1) {
    for (let col = 0; col < size; col += 1) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "wait-nonogram-creator-cell";
      if (Number(waitNonogramState.creatorMatrix[row]?.[col] || 0) === 1) {
        button.classList.add("is-on");
      }

      button.addEventListener("click", () => {
        const current = Number(waitNonogramState.creatorMatrix[row][col] || 0);
        waitNonogramState.creatorMatrix[row][col] = current === 1 ? 0 : 1;
        renderHomeNonogramCreatorGrid();
      });

      grid.appendChild(button);
    }
  }
}

function renderHomeNonogramTemplateList() {
  const list = document.getElementById("home-nonogram-template-list");
  const count = document.getElementById("home-nonogram-template-count");
  if (!list) {
    return;
  }

  const templates = getCachedNonogramTemplatesByQuiz("global");
  if (count) {
    count.textContent = String(templates.length);
  }

  if (!templates.length) {
    list.innerHTML = `<li class="wait-nonogram-leader-empty">Nenhum nonograma criado ainda.</li>`;
    return;
  }

  list.innerHTML = templates
    .slice(0, 60)
    .map((template, index) => `<li class="wait-nonogram-leader-item"><span>${index + 1}. ${String(template.name || "Desenho")}</span></li>`)
    .join("");
}

async function saveHomeNonogramCreatorTemplate() {
  if (!state.isTeacher) {
    return;
  }

  ensureNonogramCreatorMatrix();
  const rows = iconMatrixToRows(waitNonogramState.creatorMatrix, waitNonogramState.size);
  const filled = rows.join("").split("").reduce((acc, char) => acc + (char === "1" ? 1 : 0), 0);
  if (filled < 8) {
    setHomeNonogramCreatorMessage("Desenho muito vazio. Preencha mais células antes de salvar.", "error");
    return;
  }

  const nameInput = document.getElementById("home-nonogram-creator-name");
  const templateName = String(nameInput?.value || waitNonogramState.creatorName || "").trim() || `Desenho ${Date.now().toString().slice(-4)}`;
  const template = normalizeNonogramTemplate({
    id: `tpl_${Date.now().toString(36)}`,
    name: templateName,
    rows,
    createdAtMs: Date.now()
  }, waitNonogramState.size);

  if (!template) {
    setHomeNonogramCreatorMessage("Nao foi possível validar o desenho.", "error");
    return;
  }

  const result = await saveCustomNonogramTemplateForQuiz("global", template);
  if (!result?.ok) {
    if (result?.reason === "duplicate") {
      setHomeNonogramCreatorMessage("Esse desenho ja existe na biblioteca global.", "error");
    } else if (result?.reason === "unavailable") {
      setHomeNonogramCreatorMessage("Salvamento indisponivel agora. Verifique conexao e recarregue.", "error");
    } else {
      setHomeNonogramCreatorMessage("Nao foi possível salvar no momento. Tente novamente.", "error");
    }
    return;
  }

  if (result.reason === "local-only") {
    setHomeNonogramCreatorMessage("Desenho salvo apenas nesta sessao (publique as regras para salvar no banco).", "success");
  } else {
    setHomeNonogramCreatorMessage("Desenho salvo na biblioteca global!", "success");
  }
  renderHomeNonogramTemplateList();
}

function closeHomeNonogramEditorModal() {
  const modal = document.getElementById("home-nonogram-editor-modal");
  if (modal) {
    modal.classList.add("hidden");
  }
}

function ensureHomeNonogramEditorModal() {
  let modal = document.getElementById("home-nonogram-editor-modal");
  if (modal) {
    return modal;
  }

  modal = document.createElement("div");
  modal.id = "home-nonogram-editor-modal";
  modal.className = "modal-overlay hidden";
  modal.innerHTML = `
    <div class="modal-card home-nonogram-editor-card">
      <h3>Editor de nonograma</h3>
      <p>Desenhe um icone 8x8 para a biblioteca global de nonogramas.</p>
      <label class="home-nonogram-editor-field" for="home-nonogram-creator-name">
        Nome do desenho
        <input id="home-nonogram-creator-name" type="text" maxlength="40" placeholder="Ex.: Foguete" />
      </label>
      <p class="wait-nonogram-creator-msg">Sugestoes: ${getNonogramDrawingSuggestionsText()}</p>
      <div id="home-nonogram-creator-grid" class="wait-nonogram-creator-grid"></div>
      <p id="home-nonogram-creator-msg" class="wait-nonogram-creator-msg"></p>
      <section class="wait-nonogram-leaderboard-wrap">
        <h4>Nonogramas criados pelo professor (<span id="home-nonogram-template-count">0</span>)</h4>
        <ol id="home-nonogram-template-list" class="wait-nonogram-leaderboard"></ol>
      </section>
      <div class="modal-actions">
        <button id="home-nonogram-clear" class="btn btn-ghost" type="button">Limpar</button>
        <button id="home-nonogram-save" class="btn btn-primary" type="button">Salvar desenho</button>
        <button id="home-nonogram-close" class="btn btn-ghost" type="button">Fechar</button>
      </div>
    </div>
  `;

  document.body.appendChild(modal);

  modal.addEventListener("click", (event) => {
    if (event.target === modal) {
      closeHomeNonogramEditorModal();
    }
  });

  const closeButton = modal.querySelector("#home-nonogram-close");
  if (closeButton) {
    closeButton.addEventListener("click", () => {
      closeHomeNonogramEditorModal();
    });
  }

  const clearButton = modal.querySelector("#home-nonogram-clear");
  if (clearButton) {
    clearButton.addEventListener("click", () => {
      waitNonogramState.creatorMatrix = createEmptyNonogramMarks(waitNonogramState.size);
      renderHomeNonogramCreatorGrid();
      setHomeNonogramCreatorMessage("Desenho limpo.");
    });
  }

  const saveButton = modal.querySelector("#home-nonogram-save");
  if (saveButton) {
    saveButton.addEventListener("click", async () => {
      const nameInput = document.getElementById("home-nonogram-creator-name");
      waitNonogramState.creatorName = String(nameInput?.value || "").trim();

      saveButton.disabled = true;
      await saveHomeNonogramCreatorTemplate();
      saveButton.disabled = false;
    });
  }

  return modal;
}

function openHomeNonogramEditorModal() {
  if (!state.isTeacher) {
    return;
  }

  const modal = ensureHomeNonogramEditorModal();

  waitNonogramState.size = WAIT_NONOGRAM_GRID_SIZE;
  waitNonogramState.creatorMatrix = createEmptyNonogramMarks(waitNonogramState.size);
  waitNonogramState.creatorName = "";

  const nameInput = document.getElementById("home-nonogram-creator-name");
  if (nameInput) {
    nameInput.value = "";
  }

  setHomeNonogramCreatorMessage("Desenhe e salve na biblioteca global.");
  renderHomeNonogramCreatorGrid();
  loadCustomNonogramTemplatesForQuiz("global", { force: true }).then(() => {
    renderHomeNonogramTemplateList();
  });
  modal.classList.remove("hidden");
}

function ensureWaitNonogramContainer() {
  const resultScreen = document.getElementById("result-screen");
  if (!resultScreen) {
    return null;
  }

  let container = document.getElementById("result-wait-nonogram");
  if (!container) {
    container = document.createElement("section");
    container.id = "result-wait-nonogram";
    container.className = "wait-nonogram hidden";
    resultScreen.appendChild(container);
  }

  return container;
}

function hideWaitNonogram() {
  resetWaitNonogramRuntime();
  const container = document.getElementById("result-wait-nonogram");
  if (container) {
    container.classList.add("hidden");
    container.innerHTML = "";
  }
}

function updateWaitNonogramHud() {
  const progressEl = document.getElementById("wait-nonogram-progress");
  const errorsEl = document.getElementById("wait-nonogram-errors");
  const scoreEl = document.getElementById("wait-nonogram-score");
  const solvedEl = document.getElementById("wait-nonogram-solved");
  const levelEl = document.getElementById("wait-nonogram-level");
  const statusEl = document.getElementById("wait-nonogram-status");

  const stats = calcWaitNonogramStats();
  waitNonogramState.progress = stats.progress;
  waitNonogramState.roundScore = computeWaitNonogramScore({
    correctFilled: stats.correctFilled,
    errors: waitNonogramState.errors,
    completed: waitNonogramState.completed
  });
  const liveTotalScore = waitNonogramState.totalScore + (waitNonogramState.submitted ? 0 : waitNonogramState.roundScore);
  waitNonogramState.score = liveTotalScore;
  waitNonogramState.level = 1 + Math.floor(Number(waitNonogramState.solvedCount || 0) / 3);

  if (progressEl) progressEl.textContent = `${waitNonogramState.progress}%`;
  if (errorsEl) errorsEl.textContent = String(waitNonogramState.errors);
  if (scoreEl) scoreEl.textContent = String(liveTotalScore);
  if (solvedEl) solvedEl.textContent = String(waitNonogramState.solvedCount);
  if (levelEl) levelEl.textContent = String(waitNonogramState.level);

  if (statusEl) {
    if (waitNonogramState.completed) {
      statusEl.textContent = `Voce concluiu: ${waitNonogramState.currentPuzzleName || "Desenho"}. Carregando o proximo nonograma...`;
    } else if (state.isTeacher && waitNonogramState.teacherRevealMode) {
      statusEl.textContent = `Nivel ${waitNonogramState.level} ativo. Modo professor: desenho revelado para conferencia.`;
    } else {
      statusEl.textContent = `Nivel ${waitNonogramState.level} ativo. Sem limite de tempo: avance no seu ritmo.`;
    }
  }
}

function renderWaitNonogramBoard() {
  const board = document.getElementById("wait-nonogram-board");
  if (!board) {
    return;
  }

  board.innerHTML = "";
  const size = waitNonogramState.size;

  for (let row = 0; row < size; row += 1) {
    const rowWrap = document.createElement("div");
    rowWrap.className = "wait-nonogram-row";

    const clue = document.createElement("div");
    clue.className = "wait-nonogram-row-clue";
    clue.textContent = (waitNonogramState.rowClues[row] || [0]).join(" / ");
    rowWrap.appendChild(clue);

    for (let col = 0; col < size; col += 1) {
      const cell = document.createElement("button");
      cell.type = "button";
      cell.className = "wait-nonogram-cell";
      const mark = Number(waitNonogramState.marks[row]?.[col] || 0);
      const expected = Number(waitNonogramState.solution[row]?.[col] || 0);
      if (mark === 1) {
        cell.classList.add("is-filled");
      } else if (mark === -1) {
        cell.classList.add("is-crossed");
        cell.textContent = "x";
      } else if (state.isTeacher && waitNonogramState.teacherRevealMode && expected === 1) {
        cell.classList.add("is-revealed");
        cell.textContent = "•";
      }

      cell.addEventListener("click", () => {
        if (!waitNonogramState.active || waitNonogramState.completed || waitNonogramState.submitted) {
          return;
        }

        const current = Number(waitNonogramState.marks[row][col] || 0);
        const next = current === 1 ? 0 : 1;
        waitNonogramState.marks[row][col] = next;
        if (current !== 1 && next === 1 && waitNonogramState.solution[row][col] === 0) {
          waitNonogramState.errors += 1;
        }

        const stats = calcWaitNonogramStats();
        if (stats.correctFilled === stats.totalFilled && stats.wrongFilled === 0) {
          waitNonogramState.completed = true;
          submitWaitNonogramScore("completed");
        }

        renderWaitNonogramBoard();
        updateWaitNonogramHud();
      });

      cell.addEventListener("contextmenu", (event) => {
        event.preventDefault();
        if (!waitNonogramState.active || waitNonogramState.completed || waitNonogramState.submitted) {
          return;
        }

        const current = Number(waitNonogramState.marks[row][col] || 0);
        waitNonogramState.marks[row][col] = current === -1 ? 0 : -1;
        renderWaitNonogramBoard();
        updateWaitNonogramHud();
      });

      rowWrap.appendChild(cell);
    }

    board.appendChild(rowWrap);
  }
}

function renderWaitNonogramLeaderboard(rows) {
  const list = document.getElementById("wait-nonogram-leaderboard");
  if (!list) {
    return;
  }

  if (!rows.length) {
    list.innerHTML = `<li class="wait-nonogram-leader-empty">Sem resultados ainda nesta rodada.</li>`;
    return;
  }

  list.innerHTML = rows.map((entry, index) => {
    const isSelf = String(entry.slug || "") && String(entry.slug) === String(state.profileSlug || "");
    const cls = isSelf ? "wait-nonogram-leader-item is-self" : "wait-nonogram-leader-item";
    return `<li class="${cls}"><span>${index + 1}. ${entry.nome || "Aluno"}</span><strong>${entry.score || 0}</strong></li>`;
  }).join("");
}

async function fetchWaitNonogramLeaderboard() {
  if (!waitNonogramState.lobbyKey || !window.firebaseDB || !window.firebaseCollection || !window.firebaseGetDocs) {
    return;
  }

  try {
    const ref = window.firebaseCollection(window.firebaseDB, "nonogram_scores");
    let snap;
    if (window.firebaseQuery && window.firebaseWhere) {
      const q = window.firebaseQuery(
        ref,
        window.firebaseWhere("lobbyKey", "==", waitNonogramState.lobbyKey),
        window.firebaseWhere("dateKey", "==", waitNonogramState.dateKey)
      );
      snap = await window.firebaseGetDocs(q);
    } else {
      snap = await window.firebaseGetDocs(ref);
    }

    const rows = [];
    snap.forEach((docSnap) => {
      const data = docSnap.data() || {};
      if (String(data.lobbyKey || "") !== waitNonogramState.lobbyKey) {
        return;
      }
      if (String(data.dateKey || "") !== waitNonogramState.dateKey) {
        return;
      }
      rows.push({
        nome: String(data.nome || "Aluno"),
        slug: String(data.slug || ""),
        score: Number(data.score || 0),
        errors: Number(data.errors || 0),
        updatedAtMs: Number(data.updatedAtMs || 0)
      });
    });

    rows.sort((a, b) => {
      if (b.score !== a.score) {
        return b.score - a.score;
      }
      if (a.errors !== b.errors) {
        return a.errors - b.errors;
      }
      return b.updatedAtMs - a.updatedAtMs;
    });

    renderWaitNonogramLeaderboard(rows.slice(0, 10));
  } catch (error) {
    console.error("Erro ao atualizar placar do nonograma:", error);
  }
}

async function submitWaitNonogramScore(reason) {
  if (waitNonogramState.submitted) {
    return;
  }

  waitNonogramState.submitted = true;
  waitNonogramState.active = false;
  const completedRound = String(reason || "") === "completed" || waitNonogramState.completed;
  if (completedRound && !waitNonogramState.roundCounted) {
    waitNonogramState.totalScore += waitNonogramState.roundScore;
    waitNonogramState.solvedCount += 1;
    waitNonogramState.roundCounted = true;
  }
  updateWaitNonogramHud();

  const saveButton = document.getElementById("wait-nonogram-restart");
  if (saveButton) {
    saveButton.classList.remove("hidden");
  }

  if (!window.firebaseDB || !window.firebaseDoc || !window.firebaseSetDoc) {
    return;
  }

  try {
    const profile = normalizeLobbySegment(state.profileSlug || state.user?.uid || "anon", "anon");
    const docId = `${waitNonogramState.lobbyKey}__${profile}`;
    const ref = window.firebaseDoc(window.firebaseDB, "nonogram_scores", docId);

    await window.firebaseSetDoc(ref, {
      lobbyKey: waitNonogramState.lobbyKey,
      dateKey: waitNonogramState.dateKey,
      turmaId: state.classroomId || "",
      turmaNome: state.classroomName || "",
      quizId: waitNonogramState.quizId,
      puzzleIndex: waitNonogramState.puzzleIndex,
      slug: state.profileSlug || "",
      nome: state.studentName || state.user?.displayName || "Aluno",
      score: waitNonogramState.totalScore,
      roundScore: waitNonogramState.roundScore,
      solvedCount: waitNonogramState.solvedCount,
      progress: waitNonogramState.progress,
      errors: waitNonogramState.errors,
      completed: Boolean(waitNonogramState.completed),
      endReason: String(reason || "completed"),
      updatedAtMs: Date.now()
    }, { merge: true });

    fetchWaitNonogramLeaderboard();

    if (completedRound) {
      const quiz = getSelectedQuiz();
      if (quiz && String(quiz.id || "") === waitNonogramState.quizId) {
        if (waitNonogramState.autoAdvanceTimeoutId) {
          clearTimeout(waitNonogramState.autoAdvanceTimeoutId);
        }
        waitNonogramState.autoAdvanceTimeoutId = setTimeout(() => {
          startWaitNonogramRound(quiz, { advance: true });
        }, 1100);
      }
    }
  } catch (error) {
    console.error("Erro ao salvar score do nonograma:", error);
  }
}

async function startWaitNonogramRound(quiz, options = {}) {
  if (waitNonogramState.autoAdvanceTimeoutId) {
    clearTimeout(waitNonogramState.autoAdvanceTimeoutId);
    waitNonogramState.autoAdvanceTimeoutId = null;
  }
  if (options.advance) {
    waitNonogramState.puzzleIndex += 1;
  }

  const lobbyKey = buildWaitClassLobbyKey(quiz);
  const dateKey = getWaitNonogramDateKey();
  await loadCustomNonogramTemplatesForQuiz(quiz?.id);
  const activeLibrary = getActiveNonogramLibrary(quiz?.id);
  const rng = createSeededRng(`${lobbyKey}::${dateKey}::${waitNonogramState.puzzleIndex}`);
  const puzzle = buildNonogramPuzzle(WAIT_NONOGRAM_GRID_SIZE, rng, activeLibrary);
  const solution = puzzle.matrix;
  const clues = buildNonogramClues(solution);

  waitNonogramState.active = true;
  waitNonogramState.quizId = String(quiz?.id || "");
  waitNonogramState.lobbyKey = lobbyKey;
  waitNonogramState.dateKey = dateKey;
  waitNonogramState.size = WAIT_NONOGRAM_GRID_SIZE;
  waitNonogramState.solution = solution;
  waitNonogramState.currentPuzzleName = String(puzzle.name || "Desenho");
  waitNonogramState.marks = createEmptyNonogramMarks(WAIT_NONOGRAM_GRID_SIZE);
  waitNonogramState.rowClues = clues.rowClues;
  waitNonogramState.colClues = clues.colClues;
  waitNonogramState.errors = 0;
  waitNonogramState.completed = false;
  waitNonogramState.submitted = false;
  waitNonogramState.roundCounted = false;
  waitNonogramState.teacherRevealMode = false;
  waitNonogramState.progress = 0;
  waitNonogramState.roundScore = 0;
  waitNonogramState.score = waitNonogramState.totalScore;
  waitNonogramState.level = 1 + Math.floor(Number(waitNonogramState.solvedCount || 0) / 3);
  waitNonogramState.creatorName = "";
  waitNonogramState.creatorMatrix = createEmptyNonogramMarks(WAIT_NONOGRAM_GRID_SIZE);

  const container = ensureWaitNonogramContainer();
  if (!container) {
    return;
  }

  container.classList.remove("hidden");
  container.innerHTML = `
    <div class="wait-nonogram-head">
      <h3>Nonograma da Turma · Fase ${waitNonogramState.puzzleIndex + 1} · Nivel ${waitNonogramState.level}</h3>
      <p>Todos da turma recebem o mesmo puzzle do dia. Modo progressivo sem limite de tempo.</p>
    </div>
    <div class="wait-nonogram-metrics">
      <span><strong id="wait-nonogram-level">${waitNonogramState.level}</strong> nivel</span>
      <span><strong id="wait-nonogram-progress">0%</strong> progresso</span>
      <span><strong id="wait-nonogram-errors">0</strong> erros</span>
      <span><strong id="wait-nonogram-solved">${waitNonogramState.solvedCount}</strong> completos</span>
      <span><strong id="wait-nonogram-score">${waitNonogramState.totalScore}</strong> pontos</span>
    </div>
    <details class="wait-nonogram-help" open>
      <summary>Como jogar</summary>
      <ol>
        <li>Os números ao lado e no topo mostram blocos contínuos de quadrados preenchidos.</li>
        <li>Exemplo: "3 / 1" significa um bloco de 3 e depois um bloco de 1, com pelo menos um espaço entre eles.</li>
        <li>Clique para preencher um quadrado que você acha correto.</li>
        <li>Use botão direito para marcar "X" onde você acha que deve ficar vazio.</li>
        <li>Complete o ícone inteiro para registrar sua pontuação no placar.</li>
      </ol>
    </details>
    <div class="wait-nonogram-col-clues" id="wait-nonogram-col-clues"></div>
    <div class="wait-nonogram-board" id="wait-nonogram-board"></div>
    <p class="wait-nonogram-status" id="wait-nonogram-status">Nivel ${waitNonogramState.level} ativo. Sem limite de tempo: avance no seu ritmo.</p>
    <div class="wait-nonogram-actions">
      <button type="button" class="btn btn-ghost btn-sm hidden" id="wait-nonogram-restart">Próximo nonograma</button>
      <button type="button" class="btn btn-ghost btn-sm" id="wait-nonogram-clear">Limpar grade</button>
      ${state.isTeacher ? '<button type="button" class="btn btn-ghost btn-sm" id="wait-nonogram-reveal">Revelar desenho (professor)</button>' : ""}
      <button type="button" class="btn btn-ghost btn-sm" id="wait-nonogram-refresh">Atualizar placar</button>
    </div>
    <section class="wait-nonogram-leaderboard-wrap">
      <h4>Placar da turma (top 10)</h4>
      <ol id="wait-nonogram-leaderboard" class="wait-nonogram-leaderboard"></ol>
    </section>
    ${state.isTeacher ? `
      <section class="wait-nonogram-creator">
        <h4>Criador de nonograma (professor)</h4>
        <p>Crie um desenho 8x8. Os alunos poderão receber esse desenho nas próximas fases.</p>
        <p class="wait-nonogram-creator-msg">Sugestoes: ${getNonogramDrawingSuggestionsText()}</p>
        <input id="wait-nonogram-creator-name" type="text" maxlength="40" placeholder="Nome do desenho (ex.: Foguete)" />
        <div id="wait-nonogram-creator-grid" class="wait-nonogram-creator-grid"></div>
        <div class="wait-nonogram-actions">
          <button type="button" class="btn btn-ghost btn-sm" id="wait-nonogram-creator-clear">Limpar desenho</button>
          <button type="button" class="btn btn-primary btn-sm" id="wait-nonogram-creator-save">Salvar desenho</button>
        </div>
        <p id="wait-nonogram-creator-msg" class="wait-nonogram-creator-msg"></p>
      </section>
    ` : ""}
  `;

  const colClues = document.getElementById("wait-nonogram-col-clues");
  if (colClues) {
    colClues.innerHTML = waitNonogramState.colClues
      .map((clue) => `<span class="wait-nonogram-col-clue">${clue.map((value) => `<span>${value}</span>`).join("")}</span>`)
      .join("");
  }

  const restartButton = document.getElementById("wait-nonogram-restart");
  if (restartButton) {
    restartButton.addEventListener("click", () => {
      startWaitNonogramRound(quiz, { advance: true });
    });
  }

  const clearButton = document.getElementById("wait-nonogram-clear");
  if (clearButton) {
    clearButton.addEventListener("click", () => {
      clearWaitNonogramBoard();
    });
  }

  const revealButton = document.getElementById("wait-nonogram-reveal");
  if (revealButton) {
    revealButton.addEventListener("click", () => {
      toggleTeacherRevealMode();
    });
  }

  const creatorNameInput = document.getElementById("wait-nonogram-creator-name");
  if (creatorNameInput) {
    creatorNameInput.addEventListener("input", () => {
      waitNonogramState.creatorName = creatorNameInput.value;
    });
  }

  const creatorClearButton = document.getElementById("wait-nonogram-creator-clear");
  if (creatorClearButton) {
    creatorClearButton.addEventListener("click", () => {
      waitNonogramState.creatorMatrix = createEmptyNonogramMarks(waitNonogramState.size);
      renderWaitNonogramCreatorGrid();
      setWaitNonogramCreatorMessage("Desenho limpo.");
    });
  }

  const creatorSaveButton = document.getElementById("wait-nonogram-creator-save");
  if (creatorSaveButton) {
    creatorSaveButton.addEventListener("click", async () => {
      creatorSaveButton.disabled = true;
      await saveCurrentNonogramCreatorTemplate(quiz);
      creatorSaveButton.disabled = false;
    });
  }

  const refreshButton = document.getElementById("wait-nonogram-refresh");
  if (refreshButton) {
    refreshButton.addEventListener("click", async () => {
      refreshButton.disabled = true;
      await fetchWaitNonogramLeaderboard();
      setTimeout(() => {
        refreshButton.disabled = false;
      }, 900);
    });
  }

  renderWaitNonogramBoard();
  renderWaitNonogramCreatorGrid();
  updateWaitNonogramHud();
  fetchWaitNonogramLeaderboard();
}

function renderWaitNonogramForWaitMode(quiz, shouldShow) {
  if (!shouldShow || !quiz?.id) {
    hideWaitNonogram();
    return;
  }

  const lobbyKey = buildWaitClassLobbyKey(quiz);
  const dateKey = getWaitNonogramDateKey();
  const sameRound = waitNonogramState.active
    && waitNonogramState.quizId === String(quiz.id)
    && waitNonogramState.lobbyKey === lobbyKey
    && waitNonogramState.dateKey === dateKey;

  if (sameRound) {
    const container = ensureWaitNonogramContainer();
    if (container) {
      container.classList.remove("hidden");
    }
    return;
  }

  waitNonogramState.puzzleIndex = 0;
  waitNonogramState.totalScore = 0;
  waitNonogramState.solvedCount = 0;
  resetWaitNonogramRuntime();
  startWaitNonogramRound(quiz);
}

function triggerFocusWaitGrace(reasonText) {
  if (!isFocusWaitModeActive() || state.isTeacher || focusWaitPenaltyInFlight) {
    return;
  }

  if (isFocusWaitRecoverySatisfied()) {
    return;
  }

  const normalizedReason = String(reasonText || "Você saiu do foco antes da liberação do professor.").trim();

  if (!isFocusWaitGraceActive()) {
    focusWaitGraceReason = normalizedReason;
    focusWaitGraceDeadlineMs = Date.now() + 10000;
    showBottomNotice("Atenção: volte ao foco em tela cheia em até 10s para evitar penalidade.", "info", 4200);
  }

  updateFocusWaitAlert();
  clearFocusWaitGraceTimers();

  focusWaitGraceIntervalId = setInterval(() => {
    if (!isFocusWaitModeActive()) {
      clearFocusWaitGrace();
      return;
    }

    if (isFocusWaitRecoverySatisfied()) {
      resolveFocusWaitGraceWithoutPenalty();
      return;
    }

    updateFocusWaitAlert();
  }, 200);

  focusWaitGraceTimeoutId = setTimeout(() => {
    const reason = focusWaitGraceReason || normalizedReason;
    clearFocusWaitGrace();
    handleFocusWaitPenalty(reason);
  }, Math.max(0, Number(focusWaitGraceDeadlineMs || 0) - Date.now()));
}

function mergeFocusWaitPenaltyIntoAttempt(attempt, penalty) {
  if (!attempt || !penalty) {
    return attempt;
  }

  const nextResult = {
    ...(attempt.result || {}),
    focusExitPenaltyApplied: true,
    focusExitPenaltyReason: penalty.reason,
    focusExitPenaltyAt: penalty.at
  };

  return {
    ...attempt,
    focusExitPenaltyApplied: true,
    focusExitPenaltyReason: penalty.reason,
    focusExitPenaltyAt: penalty.at,
    result: nextResult
  };
}

function handleFocusWaitPenalty(reasonText) {
  const now = Date.now();
  if (focusWaitPenaltyInFlight || now - lastFocusWaitPenaltyAt < 1200) {
    return;
  }

  const context = getPendingFocusWaitContext();
  if (!context?.quiz || !context?.attempt) {
    return;
  }

  const reason = String(reasonText || "Saiu do foco antes da liberação do professor.").trim();
  const at = new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" });
  const penalty = {
    applied: true,
    reason,
    at,
    processName: FOCUS_WAIT_PROCESS_NAME
  };

  lastFocusWaitPenaltyAt = now;
  focusWaitPenaltyInFlight = true;

  const updatedAttempt = mergeFocusWaitPenaltyIntoAttempt(context.attempt, penalty);
  saveStoredAttempt(context.quiz.id, updatedAttempt);
  if (state.remoteAttemptsLoaded) {
    state.remoteAttemptsByQuiz[context.quiz.id] = updatedAttempt;
  }

  if (state.selectedQuizId === context.quiz.id && updatedAttempt.result) {
    showResult(updatedAttempt.result, "", { isBlocked: false });
    updateReviewButtons(context.quiz, updatedAttempt);
  }

  showBottomNotice("Penalidade registrada: você saiu do foco antes da liberação do professor.", "error", 4200);
  renderHome();

  persistFocusWaitPenaltyToFirestore(context.quiz.id, penalty)
    .catch((error) => {
      console.error("Erro ao salvar penalidade de espera em foco:", error);
    })
    .finally(() => {
      focusWaitPenaltyInFlight = false;
    });
}

function persistFocusWaitPenaltyOnUnload() {
  const context = getPendingFocusWaitContext();
  if (!context?.quiz || !context?.attempt) {
    return;
  }

  const reason = "Fechou a pagina durante o modo de espera sem liberacao do professor.";
  const at = new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" });

  savePendingFocusWaitPenalty({
    quizId: context.quiz.id,
    reason,
    at,
    processName: FOCUS_WAIT_PROCESS_NAME
  });
}

async function flushPendingFocusWaitPenalty() {
  const pending = readPendingFocusWaitPenalty();
  if (!pending?.quizId) {
    return;
  }

  const navEntry = (performance.getEntriesByType && performance.getEntriesByType("navigation")[0]) || null;
  const navType = String(navEntry?.type || "").toLowerCase();
  if (navType === "reload") {
    clearPendingFocusWaitPenalty();
    return;
  }

  const penalty = {
    applied: true,
    reason: String(pending.reason || "Saiu do foco sem liberacao do professor."),
    at: String(pending.at || new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" })),
    processName: FOCUS_WAIT_PROCESS_NAME
  };

  const existingAttempt = getKnownAttempt(String(pending.quizId));
  if (existingAttempt && !existingAttempt.focusExitPenaltyApplied && !existingAttempt?.result?.focusExitPenaltyApplied) {
    const updatedAttempt = mergeFocusWaitPenaltyIntoAttempt(existingAttempt, penalty);
    saveStoredAttempt(String(pending.quizId), updatedAttempt);
    if (state.remoteAttemptsLoaded) {
      state.remoteAttemptsByQuiz[String(pending.quizId)] = updatedAttempt;
    }
  }

  try {
    await persistFocusWaitPenaltyToFirestore(String(pending.quizId), penalty);
    clearPendingFocusWaitPenalty();
  } catch (error) {
    console.error("Erro ao reenviar penalidade pendente de espera em foco:", error);
  }
}

async function persistFocusWaitPenaltyToFirestore(quizId, penalty) {
  if (!quizId || !state.profileSlug || !window.firebaseDB || !window.firebaseCollection || !window.firebaseQuery || !window.firebaseWhere || !window.firebaseGetDocs || !window.firebaseSetDoc) {
    return;
  }

  const attemptsRef = window.firebaseCollection(window.firebaseDB, "usuarios", state.profileSlug, "tentativas");
  const q = window.firebaseQuery(
    attemptsRef,
    window.firebaseWhere("quizId", "==", quizId)
  );
  const snap = await window.firebaseGetDocs(q);
  if (snap.empty) {
    return;
  }

  let latestDoc = null;
  let latestTs = -1;

  snap.forEach((docSnap) => {
    const data = docSnap.data() || {};
    const ts = parsePtBrDate(String(data.data || ""));
    if (!latestDoc || ts >= latestTs) {
      latestDoc = docSnap;
      latestTs = ts;
    }
  });

  if (!latestDoc) {
    return;
  }

  await window.firebaseSetDoc(latestDoc.ref, {
    penalidadeSaidaFoco: true,
    penalidadeSaidaFocoMotivo: String(penalty.reason || ""),
    penalidadeSaidaFocoEm: String(penalty.at || ""),
    processoControleFoco: FOCUS_WAIT_PROCESS_NAME,
    atualizadoEmIso: new Date().toISOString()
  }, { merge: true });
}

function resolveReviewOptionLabel(quiz, item, index, value) {
  const normalizedValue = String(value || "");
  if (!normalizedValue) {
    return "";
  }

  const itemOptions = Array.isArray(item?.options) ? item.options : [];
  const optionFromItem = itemOptions.find((option) => String(option?.value || "") === normalizedValue);
  if (optionFromItem?.label) {
    return String(optionFromItem.label);
  }

  const byId = item?.questionId
    ? (quiz?.questions || []).find((question) => String(question?.id || "") === String(item.questionId))
    : null;
  const byIndex = (quiz?.questions || [])[index];
  const question = byId || byIndex;
  const questionOptions = Array.isArray(question?.options) ? question.options : [];
  const optionFromQuestion = questionOptions.find((option) => String(option?.value || "") === normalizedValue);
  if (optionFromQuestion?.label) {
    return String(optionFromQuestion.label);
  }

  return "";
}

function getReviewOptionsForQuestion(quiz, item, index) {
  const itemOptions = Array.isArray(item?.options)
    ? item.options
      .map((option) => {
        if (!option) {
          return null;
        }
        return {
          value: String(option.value || ""),
          label: String(option.label || "")
        };
      })
      .filter((option) => option && option.value && option.label)
    : [];

  if (itemOptions.length > 0) {
    return itemOptions;
  }

  const quizQuestions = Array.isArray(quiz?.questions) ? quiz.questions : [];
  const normalizedItemTitle = String(item?.questionTitle || "").trim().toLowerCase();
  const byId = item?.questionId
    ? quizQuestions.find((question) => String(question?.id || "") === String(item.questionId))
    : null;
  const byTitle = normalizedItemTitle
    ? quizQuestions.find((question) => String(question?.title || "").trim().toLowerCase() === normalizedItemTitle)
    : null;
  const byIndex = quizQuestions[index];
  const question = byId || byTitle || byIndex;
  const questionOptions = Array.isArray(question?.options) ? question.options : [];

  return questionOptions
    .map((option) => {
      if (!option) {
        return null;
      }
      return {
        value: String(option.value || ""),
        label: String(option.label || "")
      };
    })
    .filter((option) => option && option.value && option.label);
}

function getReviewExplanationForQuestion(quiz, item, index) {
  if (typeof item?.questionExplanation === "string" && item.questionExplanation.trim()) {
    return item.questionExplanation.trim();
  }

  const itemExplanationKeys = ["explanation", "justification", "justificativa", "feedback"];
  for (const key of itemExplanationKeys) {
    const value = item?.[key];
    if (typeof value === "string" && value.trim()) {
      return value.trim();
    }
  }

  const quizQuestions = Array.isArray(quiz?.questions) ? quiz.questions : [];
  const normalizedItemTitle = String(item?.questionTitle || "").trim().toLowerCase();
  const byId = item?.questionId
    ? quizQuestions.find((question) => String(question?.id || "") === String(item.questionId))
    : null;
  const byTitle = normalizedItemTitle
    ? quizQuestions.find((question) => String(question?.title || "").trim().toLowerCase() === normalizedItemTitle)
    : null;
  const byIndex = quizQuestions[index];
  const question = byId || byTitle || byIndex;

  const questionExplanationKeys = ["explanation", "questionExplanation", "justification", "justificativa"];
  for (const key of questionExplanationKeys) {
    const value = question?.[key];
    if (typeof value === "string" && value.trim()) {
      return value.trim();
    }
  }

  return "";
}

function renderReviewOptionsHtml(quiz, item, index) {
  const options = getReviewOptionsForQuestion(quiz, item, index);
  if (!options.length) {
    return '<p class="review-options-empty">Alternativas não disponíveis neste registro.</p>';
  }

  const selectedValue = String(item?.selectedValue || "");
  const correctValue = String(item?.correctValue || "");

  return `
    <ul class="review-options-list">
      ${options.map((option) => {
    const isSelected = selectedValue && option.value === selectedValue;
    const isCorrect = correctValue && option.value === correctValue;
    const classes = ["review-option-item"];
    if (isSelected) classes.push("is-selected");
    if (isCorrect) classes.push("is-correct");
    const optionCode = String(option.value || "").trim().toUpperCase();

    const badges = [
      isSelected ? '<span class="review-option-badge badge-selected">Marcada</span>' : "",
      isCorrect ? '<span class="review-option-badge badge-correct">Correta</span>' : ""
    ].join("");

    return `
          <li class="${classes.join(" ")}">
            <span class="review-option-main">
              ${optionCode ? `<span class="review-option-code">${escapeHtml(optionCode)})</span>` : ""}
              <span class="review-option-label">${escapeHtml(option.label)}</span>
            </span>
            <span class="review-option-badges">${badges}</span>
          </li>
        `;
  }).join("")}
    </ul>
  `;
}

function updateReviewButtons(quiz, attempt) {
  const canShow = Boolean(quiz?.allowStudentReview && attempt?.result);

  if (resultReviewButton) {
    resultReviewButton.classList.toggle("hidden", !canShow || attempt?.status !== "completed");
  }

  if (cancelReviewButton) {
    cancelReviewButton.classList.toggle("hidden", !canShow || attempt?.status !== "cancelled");
  }
}

function openReviewScreen(backScreen = "result") {
  const quiz = getSelectedQuiz();
  const attempt = getAttemptForSelectedQuiz();
  if (!quiz || !attempt?.result || !reviewList || !reviewSummaryMeta) {
    return;
  }

  state.reviewBackScreen = backScreen;
  reviewSummaryMeta.textContent = `${attempt.result.earnedPoints} / ${attempt.result.maxPoints} pontos • ${attempt.result.percent}% de acertos`;

  const questionResults = attempt?.questionResults || attempt?.result?.questionResults || [];
  if (questionResults.length > 0) {
    reviewList.innerHTML = questionResults.map((item, index) => {
      const statusClass = item.isCorrect ? "review-card-correct" : (item.selectedValue ? "review-card-wrong" : "review-card-blank");
      const statusText = item.selectedValue ? (item.isCorrect ? "Acertou" : "Errou") : "Em branco";
      const allOptionsHtml = renderReviewOptionsHtml(quiz, item, index);
      const questionExplanation = getReviewExplanationForQuestion(quiz, item, index);

      return `
        <article class="builder-question-card review-card ${statusClass}">
          <div class="builder-question-head">
            <h4>Questão ${index + 1}</h4>
            <span class="quiz-status ${item.isCorrect ? "available" : "cancelled"}">${statusText}</span>
          </div>
          <p class="review-question-title">${item.questionTitle || `Questão ${index + 1}`}</p>
          <div class="review-options-block">
            <p class="review-options-title">Alternativas</p>
            ${allOptionsHtml}
          </div>
          ${questionExplanation ? `<div class="review-explanation-block"><p class="review-options-title">Explicação</p><p class="review-question-explanation">${escapeHtml(questionExplanation)}</p></div>` : ""}
        </article>
      `;
    }).join("");
  } else {
    reviewList.innerHTML = `<article class="builder-question-card review-card review-card-blank"><p>Este registro é antigo e não possui detalhamento por questão. Novas tentativas mostrarão o resumo completo.</p></article>`;
  }

  showOnlyScreen("review");
}

async function refreshRemoteAttempts(options = {}) {
  if (!window.firebaseDB || !window.firebaseCollection || !window.firebaseGetDocs || !state.profileSlug) {
    return;
  }

  try {
    const attemptsRef = window.firebaseCollection(window.firebaseDB, "usuarios", state.profileSlug, "tentativas");
    const snap = await window.firebaseGetDocs(attemptsRef);
    const nextMap = {};
    const nextCounts = {};

    snap.forEach((docSnap) => {
      const data = docSnap.data();
      if (!data?.quizId) {
        return;
      }

      nextCounts[data.quizId] = Number(nextCounts[data.quizId] || 0) + 1;

      const current = nextMap[data.quizId];
      const nextTs = parsePtBrDate(data.data || "");
      const currentTs = current ? parsePtBrDate(current.result?.date || "") : -1;
      if (current && nextTs < currentTs) {
        return;
      }

      nextMap[data.quizId] = {
        status: data.status || "completed",
        cancelReason: data.motivoCancelamento || "",
        blockedByViolation: Boolean(data.bloqueadoPorViolacao),
        focusExitPenaltyApplied: Boolean(data.penalidadeSaidaFoco),
        focusExitPenaltyReason: String(data.penalidadeSaidaFocoMotivo || ""),
        focusExitPenaltyAt: String(data.penalidadeSaidaFocoEm || ""),
        result: {
          quizId: data.quizId,
          quizTitle: data.quizTitulo,
          studentName: data.nome,
          earnedPoints: data.pontos ?? 0,
          maxPoints: data.total ?? 0,
          percent: data.percentual ?? 0,
          date: data.data,
          startedAt: data.inicioAvaliacao || "",
          endedAt: data.fimAvaliacao || data.data,
          durationSeconds: Number(data.duracaoSegundos || 0),
          durationLabel: data.duracaoTexto || "",
          questionResults: data.correcaoQuestoes || [],
          focusExitPenaltyApplied: Boolean(data.penalidadeSaidaFoco),
          focusExitPenaltyReason: String(data.penalidadeSaidaFocoMotivo || ""),
          focusExitPenaltyAt: String(data.penalidadeSaidaFocoEm || "")
        },
        answers: data.respostas || {},
        questionResults: data.correcaoQuestoes || []
      };
    });

    state.remoteAttemptsByQuiz = nextMap;
    state.remoteAttemptCountsByQuiz = nextCounts;
    state.remoteAttemptsLoaded = true;

    getAllQuizzes().forEach((quiz) => {
      const storageKey = getStorageKey(quiz.id);
      if (nextMap[quiz.id]) {
        localStorage.setItem(storageKey, JSON.stringify(nextMap[quiz.id]));
      } else {
        localStorage.removeItem(storageKey);
      }

      setStoredAttemptCount(quiz.id, Number(nextCounts[quiz.id] || 0));
    });

    if (options.render !== false) {
      renderHome();
    }
  } catch (error) {
    console.error("Erro ao carregar tentativas remotas:", error);
  }
}

function renderHome() {
  renderFocusWaitPersistentNotice();

  const allQuizzes = getAllQuizzes();
  const postedClassroomsMap = getPostedClassroomsByQuizId();
  const filterText = quizSearch.value.trim().toLowerCase();
  const list = allQuizzes.filter((quiz) => {
    const data = `${quiz.title} ${quiz.description}`.toLowerCase();
    return data.includes(filterText);
  });

  if (list.length === 0) {
    const emptySignature = `empty|${filterText}|${state.isTeacher ? "teacher" : "student"}`;
    if (lastHomeRenderSignature === emptySignature) {
      return;
    }
    quizGrid.innerHTML = `<article class="card"><p>Nenhum quiz encontrado.</p></article>`;
    lastHomeRenderSignature = emptySignature;
    return;
  }

  const fragment = document.createDocumentFragment();
  const signatureParts = [];

  list.forEach((quiz) => {
    let attempt = getKnownAttempt(quiz.id);
    const attemptsUsed = getKnownAttemptCount(quiz.id);
    const maxAttempts = resolveQuizMaxAttempts(quiz);
    const accessLocked = isQuizAccessLocked(quiz.id);
    const hasRemainingAttempts = state.isTeacher || (!accessLocked && attemptsUsed < maxAttempts);
    // Se o professor liberou nova tentativa, remove status cancelado/completed
    if (isRetakeReleased(quiz.id)) {
      // Remove do localStorage e do estado
      localStorage.removeItem(getStorageKey(quiz.id));
      if (state.remoteAttemptsByQuiz) delete state.remoteAttemptsByQuiz[quiz.id];
      attempt = null;
    }
    const canStartQuiz = state.isTeacher || (Boolean(quiz.allowStudentStart) && (isRetakeReleased(quiz.id) || hasRemainingAttempts));
    const canOpenQuiz = canStartQuiz || Boolean(attempt);
    const mainButtonLabel = canStartQuiz ? "Abrir quiz" : (attempt ? "Ver detalhes" : "Início bloqueado");
    const status = attempt?.status || (isRetakeReleased(quiz.id) ? "available" : "available");
    const statusLabel =
      status === "completed" ? "Concluído" :
        status === "cancelled" ? "Bloqueado" :
          "Disponível";

    signatureParts.push([
      quiz.id,
      status,
      canOpenQuiz ? "1" : "0",
      quiz.allowStudentStart ? "1" : "0",
      quiz.allowStudentReview ? "1" : "0",
      quiz.allowFocusExitAfterFinish ? "1" : "0",
      accessLocked ? "1" : "0",
      `${attemptsUsed}/${maxAttempts}`,
      attempt?.result?.date || "",
      state.isTeacher ? (postedClassroomsMap[quiz.id] || []).map((item) => item.id).join(",") : "",
      state.isTeacher ? "1" : "0"
    ].join("|"));

    const postedClassrooms = postedClassroomsMap[quiz.id] || [];
    const postedByClassroomId = {};
    postedClassrooms.forEach((item) => {
      postedByClassroomId[item.id] = item;
    });

    const chipsHtml = [];

    if (state.classrooms.length > 0) {
      state.classrooms.forEach((classroom) => {
        const activePost = postedByClassroomId[classroom.id] || null;
        if (activePost) {
          chipsHtml.push(`
            <span class="quiz-access-chip is-active">
              <span class="quiz-access-chip-label">${escapeHtml(classroom.nome)}</span>
              <button type="button" class="quiz-access-chip-action is-remove" data-remove-post="${escapeHtml(activePost.postId)}" data-remove-classroom-name="${escapeHtml(classroom.nome)}" data-remove-quiz-title="${escapeHtml(quiz.title)}" title="Remover acesso da turma" aria-label="Remover acesso da turma ${escapeHtml(classroom.nome)}">x</button>
            </span>
          `);
          return;
        }

        chipsHtml.push(`
          <span class="quiz-access-chip is-inactive">
            <span class="quiz-access-chip-label">${escapeHtml(classroom.nome)}</span>
            <button type="button" class="quiz-access-chip-action is-add" data-add-classroom="${escapeHtml(classroom.id)}" data-add-classroom-name="${escapeHtml(classroom.nome)}" data-add-quiz-title="${escapeHtml(quiz.title)}" title="Adicionar acesso para esta turma" aria-label="Adicionar acesso para a turma ${escapeHtml(classroom.nome)}">+</button>
          </span>
        `);
      });
    }

    postedClassrooms.forEach((item) => {
      if (state.classrooms.some((classroom) => classroom.id === item.id)) {
        return;
      }

      chipsHtml.push(`
        <span class="quiz-access-chip is-active">
          <span class="quiz-access-chip-label">${escapeHtml(item.nome)}</span>
          <button type="button" class="quiz-access-chip-action is-remove" data-remove-post="${escapeHtml(item.postId)}" data-remove-classroom-name="${escapeHtml(item.nome)}" data-remove-quiz-title="${escapeHtml(quiz.title)}" title="Remover acesso da turma" aria-label="Remover acesso da turma ${escapeHtml(item.nome)}">x</button>
        </span>
      `);
    });

    const postedClassroomsHtml = chipsHtml.length
      ? `<span class="quiz-access-list">${chipsHtml.join("")}</span>`
      : "<span>Nenhuma turma cadastrada</span>";
    const card = document.createElement("article");
    card.className = "card quiz-card";
    card.innerHTML = `
      <div class="quiz-card-head">
        <h3>${quiz.title}</h3>
        <span class="quiz-status ${status}">${statusLabel}</span>
      </div>
      <p>${quiz.description}</p>
      <ul class="quiz-meta">
        <li>Questões: ${quiz.questions.length}</li>
        <li>Tentativas: ${state.isTeacher ? maxAttempts : `${attemptsUsed}/${maxAttempts}`}</li>
      </ul>
      ${state.isTeacher ? `<p class="quiz-visibility-note"><strong>Visível para:</strong> ${postedClassroomsHtml}</p>` : ""}
      <div class="quiz-card-actions">
        <button class="btn btn-primary btn-sm ${canOpenQuiz ? "" : "btn-start-blocked"}" data-quiz="${quiz.id}" title="${canOpenQuiz ? "" : "Atualizar"}">
          ${mainButtonLabel}
        </button>
        ${attempt?.result ? (quiz.allowStudentReview ? `<button class="btn btn-primary btn-sm" data-quick-review="${quiz.id}">Resumo</button>` : `<button class="btn btn-primary btn-sm btn-summary-blocked" data-summary-blocked="${quiz.id}" title="Atualizar">Resumo</button>`) : ""}
        ${state.isTeacher ? `<button class="btn btn-ghost btn-sm" data-toggle-start="${quiz.id}">${quiz.allowStudentStart ? "Bloquear início" : "Liberar início"}</button>` : ""}
        ${state.isTeacher ? `<button class="btn btn-ghost btn-sm" data-toggle-review="${quiz.id}">${quiz.allowStudentReview ? "Bloquear resumo" : "Liberar resumo"}</button>` : ""}
        ${state.isTeacher ? `<button class="btn btn-ghost btn-sm" data-toggle-focus-exit="${quiz.id}">${quiz.allowFocusExitAfterFinish ? "Ativar modo de espera" : FOCUS_WAIT_RELEASE_BUTTON_LABEL}</button>` : ""}
        ${state.isTeacher ? `
          <div class="quiz-card-manage-actions">
            <button class="btn btn-ghost btn-sm" data-edit-quiz="${quiz.id}" title="Editar quiz" aria-label="Editar quiz">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
                <path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25z" fill="currentColor" />
                <path d="M20.71 7.04a1 1 0 0 0 0-1.41l-2.34-2.34a1 1 0 0 0-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z" fill="currentColor" />
              </svg>
              Editar
            </button>
            <button class="btn btn-danger-soft btn-sm" data-remove-quiz="${quiz.id}" title="Excluir quiz" aria-label="Excluir quiz">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
                <path d="M9 3v1H4v2h16V4h-5V3H9z" fill="currentColor" />
                <path d="M6 7l1 14a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-14H6z" fill="currentColor" />
              </svg>
              Excluir
            </button>
          </div>
        ` : ""}
      </div>
      ${!state.isTeacher && !attempt && !quiz.allowStudentStart ? `<p class="quiz-review-note">Início bloqueado no momento. Aguarde o professor liberar este quiz.</p>` : ""}
      ${!state.isTeacher && !isRetakeReleased(quiz.id) && attemptsUsed >= maxAttempts ? `<p class="quiz-review-note">Limite de tentativas atingido (${attemptsUsed}/${maxAttempts}).</p>` : ""}
      ${attempt && !quiz.allowStudentReview ? `<p class="quiz-review-note">Resumo bloqueado no momento.</p>` : ""}
      ${!state.isTeacher && attempt?.status === "completed" && !quiz.allowFocusExitAfterFinish && !attempt?.focusExitPenaltyApplied ? `<p class="quiz-review-note">${FOCUS_WAIT_PROCESS_NAME} ativo: permaneça no site até o professor clicar em '${FOCUS_WAIT_RELEASE_BUTTON_LABEL}'.</p>` : ""}
      ${!state.isTeacher && attempt?.focusExitPenaltyApplied ? `<p class="quiz-review-note">Penalidade aplicada: saiu do foco antes da liberação do professor.</p>` : ""}
    `;

    const quickReviewButton = card.querySelector("button[data-quick-review]");
    if (quickReviewButton) {
      quickReviewButton.addEventListener("click", async () => {
        state.selectedQuizId = quiz.id;
        const syncedAttempt = await syncLatestAttemptForQuiz(quiz.id);
        if (!syncedAttempt?.result) {
          return;
        }
        openReviewScreen("home");
      });
    }

    const blockedSummaryButton = card.querySelector("button[data-summary-blocked]");
    if (blockedSummaryButton) {
      blockedSummaryButton.addEventListener("mouseenter", () => {
        blockedSummaryButton.textContent = "Atualizar";
      });
      blockedSummaryButton.addEventListener("mouseleave", () => {
        blockedSummaryButton.textContent = "Resumo";
      });
      blockedSummaryButton.addEventListener("click", () => {
        window.location.reload();
      });
    }

    const mainButton = card.querySelector('button[data-quiz]');
    if (mainButton) {
      if (!canOpenQuiz && !attempt) {
        mainButton.addEventListener("mouseenter", () => {
          mainButton.textContent = "Atualizar";
        });
        mainButton.addEventListener("mouseleave", () => {
          mainButton.textContent = "Início bloqueado";
        });
      }

      mainButton.addEventListener('click', async () => {
        if (!canOpenQuiz && !attempt) {
          window.location.reload();
          return;
        }

        if (canStartQuiz) {
          await openQuiz(quiz.id);
          return;
        }

        if (attempt) {
          // Show result screen for this attempt
          state.selectedQuizId = quiz.id;
          const syncedAttempt = await syncLatestAttemptForQuiz(quiz.id);
          if (!syncedAttempt?.result) {
            return;
          }
          if (syncedAttempt.status === "cancelled") {
            const reason = syncedAttempt.cancelReason || "Avaliação bloqueada por violação de regras.";
            showResult(syncedAttempt.result, "", { isBlocked: true, reason });
          } else {
            showResult(syncedAttempt.result, "");
          }
          updateReviewButtons(quiz, syncedAttempt);
          showOnlyScreen("result");
        }
      });
    }

    const editButton = card.querySelector("button[data-edit-quiz]");
    if (editButton) {
      editButton.addEventListener("click", async () => {
        await startEditingQuiz(quiz.id);
      });
    }

    const removeButton = card.querySelector("button[data-remove-quiz]");
    if (removeButton) {
      removeButton.addEventListener("click", async () => {
        await removeQuiz(quiz.id);
      });
    }

    const toggleReviewButton = card.querySelector("button[data-toggle-review]");
    if (toggleReviewButton) {
      toggleReviewButton.addEventListener("click", async () => {
        await toggleQuizReview(quiz.id);
      });
    }

    const toggleStartButton = card.querySelector("button[data-toggle-start]");
    if (toggleStartButton) {
      toggleStartButton.addEventListener("click", async () => {
        await toggleQuizStart(quiz.id);
      });
    }

    const toggleFocusExitButton = card.querySelector("button[data-toggle-focus-exit]");
    if (toggleFocusExitButton) {
      toggleFocusExitButton.addEventListener("click", async () => {
        await toggleQuizFocusExitRelease(quiz.id);
      });
    }

    card.querySelectorAll("button[data-add-classroom]").forEach((button) => {
      button.addEventListener("click", async (event) => {
        event.preventDefault();
        event.stopPropagation();
        const classroomId = String(button.dataset.addClassroom || "").trim();
        if (!classroomId) {
          return;
        }

        button.disabled = true;
        button.textContent = "...";
        await postQuizToClassroom(quiz.id, classroomId);
        button.disabled = false;
        button.textContent = "+";
      });
    });

    card.querySelectorAll("button[data-remove-post]").forEach((button) => {
      button.addEventListener("click", async (event) => {
        event.preventDefault();
        event.stopPropagation();
        const postId = String(button.dataset.removePost || "").trim();
        const classroomName = String(button.dataset.removeClassroomName || "").trim() || "Turma";
        const quizTitle = String(button.dataset.removeQuizTitle || quiz.title || "Quiz").trim();
        await removeQuizPostFromClassroom(postId, quizTitle, classroomName);
      });
    });

    fragment.appendChild(card);
  });

  const nextSignature = `${filterText}|${state.isTeacher ? "teacher" : "student"}|${signatureParts.join("||")}`;
  if (lastHomeRenderSignature === nextSignature) {
    return;
  }

  quizGrid.replaceChildren(fragment);
  lastHomeRenderSignature = nextSignature;

  if (!state.isTeacher && isFocusWaitModeActive() && document.visibilityState === "visible" && !isFullscreenActive()) {
    triggerFocusWaitGrace("O modo foco exige tela cheia até o professor liberar o modo de espera.");
  }
}


async function getRemoteAttemptForQuiz(quizId) {
  if (!window.firebaseDB || !window.firebaseCollection || !window.firebaseQuery || !window.firebaseWhere || !window.firebaseGetDocs) {
    return null;
  }

  if (!state.profileSlug) {
    return null;
  }

  try {
    const attemptsRef = window.firebaseCollection(window.firebaseDB, "usuarios", state.profileSlug, "tentativas");
    const q = window.firebaseQuery(
      attemptsRef,
      window.firebaseWhere("quizId", "==", quizId)
    );
    const snap = await window.firebaseGetDocs(q);

    if (snap.empty) {
      return null;
    }

    let latest = null;
    snap.forEach((docSnap) => {
      const data = docSnap.data();
      if (!latest || parsePtBrDate(data.data || "") >= parsePtBrDate(latest.data || "")) {
        latest = data;
      }
    });

    if (!latest) {
      return null;
    }

    return {
      status: latest.status || "completed",
      cancelReason: latest.motivoCancelamento || "",
      blockedByViolation: Boolean(latest.bloqueadoPorViolacao),
      focusExitPenaltyApplied: Boolean(latest.penalidadeSaidaFoco),
      focusExitPenaltyReason: String(latest.penalidadeSaidaFocoMotivo || ""),
      focusExitPenaltyAt: String(latest.penalidadeSaidaFocoEm || ""),
      result: {
        quizId: latest.quizId,
        quizTitle: latest.quizTitulo,
        studentName: latest.nome,
        earnedPoints: latest.pontos ?? 0,
        maxPoints: latest.total ?? 0,
        percent: latest.percentual ?? 0,
        date: latest.data,
        startedAt: latest.inicioAvaliacao || "",
        endedAt: latest.fimAvaliacao || latest.data,
        durationSeconds: Number(latest.duracaoSegundos || 0),
        durationLabel: latest.duracaoTexto || "",
        questionResults: latest.correcaoQuestoes || [],
        focusExitPenaltyApplied: Boolean(latest.penalidadeSaidaFoco),
        focusExitPenaltyReason: String(latest.penalidadeSaidaFocoMotivo || ""),
        focusExitPenaltyAt: String(latest.penalidadeSaidaFocoEm || "")
      },
      answers: latest.respostas || {},
      questionResults: latest.correcaoQuestoes || []
    };
  } catch (error) {
    console.error("Erro ao consultar tentativa no Firestore:", error);
    return null;
  }
}

async function openQuiz(quizId) {
  state.selectedQuizId = quizId;
  const quiz = getSelectedQuiz();
  let attempt = getKnownAttempt(quizId);

  if (!attempt) {
    const remoteAttempt = await getRemoteAttemptForQuiz(quizId);
    if (remoteAttempt) {
      attempt = remoteAttempt;
      saveStoredAttempt(quizId, remoteAttempt);
    }
  }

  if (!quiz) {
    return;
  }

  const retakeReleased = isRetakeReleased(quizId);
  const isStartAllowed = state.isTeacher || Boolean(quiz.allowStudentStart);
  const attemptsUsed = getKnownAttemptCount(quizId);
  const maxAttempts = resolveQuizMaxAttempts(quiz);
  const accessLocked = isQuizAccessLocked(quizId);
  const hasRemainingAttempts = retakeReleased || state.isTeacher || (!accessLocked && attemptsUsed < maxAttempts);

  if (!hasRemainingAttempts && attempt?.status === "completed") {
    resultAlreadyCompleted.classList.remove("hidden");
    showResult(attempt.result, `Você já realizou este questionário. Limite de tentativas atingido (${attemptsUsed}/${maxAttempts}).`);
    updateReviewButtons(quiz, attempt);
    return;
  }

  if (!hasRemainingAttempts && attempt?.status === "cancelled") {
    const base = attempt.cancelReason || "Avaliação bloqueada por violação de regras.";
    showResult(attempt.result, `Limite de tentativas atingido (${attemptsUsed}/${maxAttempts}).`, { isBlocked: true, reason: base });
    updateReviewButtons(quiz, attempt);
    return;
  }

  selectedQuizTitle.textContent = quiz.title;
  selectedQuizDescription.textContent = quiz.description;
  const hasTimedQuestion = quiz.questions.some((question) => Number(question.timer || 0) > 0);
  if (studentNameDisplay) {
    studentNameDisplay.textContent = state.studentName || "-";
  }
  selectedQuizMeta.innerHTML =
    `<li>Duração estimada: ${quiz.duration}</li>` +
    `<li>Total de questões: ${quiz.questions.length}</li>` +
    `<li>Tentativas permitidas: ${maxAttempts}</li>` +
    `<li>Tentativas usadas: ${attemptsUsed}</li>`;

  if (timedWarningDisplay) {
    if (hasTimedQuestion) {
      timedWarningDisplay.textContent = "⏱ Atenção: este quiz possui questões com tempo limite.";
      timedWarningDisplay.classList.remove("hidden");
    } else {
      timedWarningDisplay.classList.add("hidden");
    }
  }

  if (startRulesList) {
    startRulesList.innerHTML =
      `<li>Não é possível voltar para questões anteriores.</li>` +
      `<li>Você pode pular uma questão para responder depois (até 3 vezes por tentativa).</li>` +
      `<li>As questões e alternativas são embaralhadas.</li>` +
      `<li>Não recarregue a página e não troque de tela; isso bloqueia a avaliação.</li>` +
      `<li>Ao bloquear o quiz, você perde o acesso à tentativa.</li>` +
      `<li>${formatAttemptsRemainingMessage(quiz, attemptsUsed)}</li>` +
      (isStartAllowed ? "" : "<li><strong>Início bloqueado pelo professor no momento.</strong></li>");
  }

  if (startQuizButton) {
    const canStartNow = isStartAllowed && hasRemainingAttempts;
    startQuizButton.disabled = !canStartNow;
    startQuizButton.textContent = canStartNow ? "Iniciar questionário" : "Início bloqueado";
  }
  if (startReloadButton) {
    if (isStartAllowed) {
      startReloadButton.classList.add("hidden");
    } else {
      startReloadButton.classList.remove("hidden");
    }
  }

  showOnlyScreen("start");
}

async function refreshHomeNavigationData() {
  await refreshTeacherAccess({ render: false });
  await Promise.allSettled([
    refreshQuizConfigs({ render: false }),
    refreshCustomQuizzes({ render: false }),
    refreshRetakeReleases({ render: false }),
    refreshRemoteAttempts({ render: false }),
    refreshClassrooms()
  ]);
}

async function showHome(options = {}) {
  const shouldRefresh = options.refresh !== false;
  clearSessionState();
  showOnlyScreen(null);
  teacherScreen.classList.add("hidden");
  if (teacherClassroomsScreen) {
    teacherClassroomsScreen.classList.add("hidden");
  }
  if (builderScreen) {
    builderScreen.classList.add("hidden");
  }
  homeScreen.classList.remove("hidden");
  resultAlreadyCompleted.classList.add("hidden");

  if (shouldRefresh) {
    await refreshHomeNavigationData();
  }

  renderHome();
}

function showOnlyScreen(screenName) {
  Object.values(screens).forEach((screen) => screen.classList.add("hidden"));
  document.body.classList.toggle("quiz-mode", screenName === "quiz");

  if (screenName !== "quiz") {
    hideQuestionImageViewer();
  }

  if (appHeader) {
    appHeader.classList.toggle("hidden", screenName === "quiz");
  }

  if (screenName) {
    homeScreen.classList.add("hidden");
    teacherScreen.classList.add("hidden");
  } else if (state.user) {
    homeScreen.classList.remove("hidden");
  }

  if (screenName && screens[screenName]) {
    screens[screenName].classList.remove("hidden");
  }
}

function clearSessionState() {
  state.currentQuestionIndex = 0;
  state.activeQuestions = [];
  state.answers = {};
  state.skipQuestionUsesRemaining = 3;
  state.isActive = false;
  state.isCancelled = false;
  state.timedOutQuestionId = null;
  state.reviewBackScreen = "result";
  state.violationCount = 0;
  state.isViolationGraceActive = false;
  state.attemptStartedAtMs = 0;
  state.attemptStartedAtLabel = "";
  clearQuestionTimerState();
  clearViolationTimers();
  hidePolicyWarning();
  clearFocusWaitGrace();
  hideWaitNonogram();
  hideQuestionImageViewer();
  renderFocusWaitPersistentNotice();
  hideCancelModal();
  applyZoomCounterScale();
}

function hideCancelModal() {
  pendingModalAction = null;
  if (cancelModal) {
    cancelModal.classList.add("hidden");
  }
}

function openActionModal({ title, text, confirmLabel, onConfirm }) {
  if (!cancelModal || !cancelModalTitle || !cancelModalText || !cancelModalConfirm) {
    if (typeof onConfirm === "function") {
      onConfirm();
    }
    return;
  }

  cancelModalTitle.textContent = title;
  cancelModalText.textContent = text;
  cancelModalConfirm.textContent = confirmLabel;
  pendingModalAction = onConfirm;
  cancelModal.classList.remove("hidden");
}

function tryHandleReloadShortcut(event) {
  if (!event || !state.isActive || state.isTeacher || !getSelectedQuiz()) {
    return false;
  }

  const key = String(event.key || "").toLowerCase();
  const isReloadShortcut = key === "f5" || ((event.ctrlKey || event.metaKey) && key === "r");
  if (!isReloadShortcut) {
    return false;
  }

  event.preventDefault();
  openActionModal({
    title: "Recarregar página?",
    text: "Se você recarregar agora, este questionário será bloqueado e você perderá o acesso a ele.",
    confirmLabel: "Sim, recarregar",
    onConfirm: async () => {
      await cancelQuiz("Questionário bloqueado ao recarregar a página durante a avaliação.", { blockedByViolation: true, skipScreen: true });
      window.location.reload();
    }
  });

  return true;
}

function updateNextButtonVisibility() {
  if (!nextButton) {
    return;
  }

  const selected = questionForm?.querySelector("input[type='radio']:checked");
  nextButton.classList.remove("hidden");
  nextButton.disabled = !selected || Boolean(state.timedOutQuestionId);
  if (clearAnswerButton) {
    clearAnswerButton.classList.remove("hidden");
  }
  updateSkipButtonVisibility();
  // mark selected option visually
  if (questionForm) {
    const allOptions = questionForm.querySelectorAll('.answer-option');
    allOptions.forEach((opt) => opt.classList.remove('selected'));
    if (selected) {
      const parent = selected.closest('.answer-option');
      if (parent) parent.classList.add('selected');
    }
  }
}

function updateSkipButtonVisibility() {
  if (!skipQuestionButton) {
    return;
  }

  if (state.isTeacher) {
    skipQuestionButton.classList.add("hidden");
    return;
  }

  const quiz = getSelectedQuiz();
  const activeQuestions = state.activeQuestions.length ? state.activeQuestions : (quiz?.questions || []);
  const canSkipByCount = Number(state.skipQuestionUsesRemaining || 0) > 0;
  const hasAnotherQuestion = Number(state.currentQuestionIndex || 0) < activeQuestions.length - 1;
  const isLockedByTimeout = Boolean(state.timedOutQuestionId);

  skipQuestionButton.classList.remove("hidden");
  skipQuestionButton.textContent = `Pular (${Math.max(0, Number(state.skipQuestionUsesRemaining || 0))})`;
  skipQuestionButton.disabled = state.isTeacher || !state.isActive || !canSkipByCount || !hasAnotherQuestion || isLockedByTimeout;
}

function clearQuestionTimerState() {
  if (window.questionTimerInterval) {
    clearInterval(window.questionTimerInterval);
    window.questionTimerInterval = null;
  }
  window.questionTimerDeadlineMs = 0;
  clearQuestionAutoAdvanceState();
  if (questionInfoArea) {
    questionInfoArea.innerHTML = "";
  }
  if (questionForm) {
    questionForm.classList.remove("question-timeout-locked");
  }
}

function clearQuestionAutoAdvanceState() {
  if (window.questionAutoAdvanceInterval) {
    clearInterval(window.questionAutoAdvanceInterval);
    window.questionAutoAdvanceInterval = null;
  }
  if (window.questionAutoAdvanceTimeout) {
    clearTimeout(window.questionAutoAdvanceTimeout);
    window.questionAutoAdvanceTimeout = null;
  }
  window.questionAutoAdvanceDeadlineMs = 0;
}

function advanceQuestion(options = {}) {
  clearQuestionAutoAdvanceState();

  const quiz = getSelectedQuiz();
  if (!quiz) {
    return;
  }

  const activeQuestions = state.activeQuestions.length ? state.activeQuestions : quiz.questions;
  const question = activeQuestions[state.currentQuestionIndex];
  if (!question) {
    finishQuiz();
    return;
  }

  const selected = questionForm.querySelector(`input[name="${question.id}"]:checked`);
  if (!options.allowBlank && !selected) {
    showValidationMessage("Selecione uma resposta para continuar.");
    return;
  }

  clearValidationMessage();
  if (selected) {
    state.answers[question.id] = selected.value;
  }
  state.timedOutQuestionId = null;
  state.currentQuestionIndex += 1;

  if (state.currentQuestionIndex >= activeQuestions.length) {
    finishQuiz();
    return;
  }

  renderQuestion();
}

function handleQuestionTimerExpired(question) {
  state.timedOutQuestionId = question.id;
  clearQuestionTimerState();
  clearValidationMessage();

  const selected = questionForm.querySelector(`input[name="${question.id}"]:checked`);
  const hasSelectedAnswer = Boolean(selected);
  if (selected) {
    state.answers[question.id] = selected.value;
  }

  Array.from(questionForm.querySelectorAll("input")).forEach((input) => {
    input.disabled = true;
  });

  questionForm.classList.add("question-timeout-locked");

  nextButton.disabled = true;
  nextButton.classList.remove("hidden");
  if (clearAnswerButton) {
    clearAnswerButton.classList.remove("hidden");
  }

  if (!questionInfoArea) {
    advanceQuestion({ allowBlank: true });
    return;
  }

  const notice = document.createElement("div");
  notice.className = "question-timeout-notice";
  notice.innerHTML = `
    <p>${hasSelectedAnswer ? "O tempo desta questão acabou. Sua resposta foi registrada." : "O tempo desta questão acabou. Sua resposta não foi registrada."}</p>
    <p class="question-timeout-autonext">indo para a próxima questão em <strong id="question-timeout-countdown">10</strong>s</p>
    <button type="button" class="btn btn-primary" id="question-timeout-next">Ir para a próxima questão</button>
  `;
  questionInfoArea.innerHTML = "";
  questionInfoArea.appendChild(notice);

  const timeoutButton = document.getElementById("question-timeout-next");
  const timeoutCountdown = document.getElementById("question-timeout-countdown");
  window.questionAutoAdvanceDeadlineMs = Date.now() + 10000;

  const updateAutoAdvanceCountdown = () => {
    const remainingMs = Math.max(0, Number(window.questionAutoAdvanceDeadlineMs || 0) - Date.now());
    const remainingSeconds = Math.ceil(remainingMs / 1000);

    if (timeoutCountdown) {
      timeoutCountdown.textContent = String(Math.max(0, remainingSeconds));
    }

    if (remainingMs === 0) {
      clearQuestionAutoAdvanceState();
      if (state.timedOutQuestionId === question.id) {
        advanceQuestion({ allowBlank: true });
      }
    }
  };

  updateAutoAdvanceCountdown();
  window.questionAutoAdvanceInterval = setInterval(updateAutoAdvanceCountdown, 200);

  if (timeoutButton) {
    timeoutButton.addEventListener("click", () => {
      advanceQuestion({ allowBlank: true });
    });
  }
}

function renderQuestionTimer(question) {
  clearQuestionTimerState();

  if (!questionInfoArea || !question.timer || question.timer <= 0) {
    return;
  }

  const durationMs = Math.max(0, Number(question.timer || 0) * 1000);
  window.questionTimerDeadlineMs = Date.now() + durationMs;

  const timerBox = document.createElement("div");
  timerBox.className = "question-timer";
  timerBox.innerHTML = `<span>Tempo restante</span> <strong>${Math.ceil(durationMs / 1000)} s</strong>`;
  timerBox.classList.toggle("question-timer-danger", durationMs <= 10000);
  questionInfoArea.appendChild(timerBox);

  const updateTimer = () => {
    const remainingMs = Math.max(0, Number(window.questionTimerDeadlineMs || 0) - Date.now());
    const timeLeft = Math.ceil(remainingMs / 1000);
    const strong = timerBox.querySelector("strong");
    if (strong) {
      strong.textContent = `${Math.max(0, timeLeft)} s`;
    }
    timerBox.classList.toggle("question-timer-danger", timeLeft <= 10);

    if (remainingMs === 0) {
      clearQuestionTimerState();
      handleQuestionTimerExpired(question);
    }
  };

  updateTimer();
  window.questionTimerInterval = setInterval(updateTimer, 200);
}

async function handleStartQuiz(event) {
  event.preventDefault();
  const quiz = getSelectedQuiz();
  if (!quiz) {
    return;
  }

  if (!state.isTeacher) {
    await refreshRemoteAttempts({ render: false });
  }

  const retakeReleased = isRetakeReleased(quiz.id);
  const attemptsUsed = getKnownAttemptCount(quiz.id);
  const maxAttempts = resolveQuizMaxAttempts(quiz);
  const accessLocked = isQuizAccessLocked(quiz.id);
  const hasRemainingAttempts = retakeReleased || state.isTeacher || (!accessLocked && attemptsUsed < maxAttempts);

  if (!state.isTeacher && !quiz.allowStudentStart) {
    alert("Este quiz está bloqueado para início no momento. Aguarde o professor liberar.");
    return;
  }

  if (!state.isTeacher && accessLocked && !retakeReleased) {
    alert("Você perdeu o acesso a novas tentativas deste questionário.");
    return;
  }

  if (!hasRemainingAttempts) {
    alert(`Limite de tentativas atingido (${attemptsUsed}/${maxAttempts}).`);
    await openQuiz(quiz.id);
    return;
  }

  if (!state.studentName) {
    alert("Complete seu perfil antes de iniciar o quiz.");
    return;
  }

  const enteredFullscreen = await requestFullscreenMode();
  if (!enteredFullscreen) {
    alert("E obrigatorio permanecer em tela cheia para realizar o questionario.");
    return;
  }

  clearSessionState();
  state.activeQuestions = buildShuffledQuestionSet(quiz);
  state.skipQuestionUsesRemaining = 3;
  state.isActive = true;
  state.attemptStartedAtMs = Date.now();
  state.attemptStartedAtLabel = new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" });
  lockZoomForQuiz();

  await consumeRetakeRelease(quiz.id);

  showOnlyScreen("quiz");
  cancelButton.classList.add("hidden");
  renderQuestion();
}

function renderQuestion() {
  const quiz = getSelectedQuiz();
  if (!quiz) {
    return;
  }

  const activeQuestions = state.activeQuestions.length ? state.activeQuestions : quiz.questions;
  const question = activeQuestions[state.currentQuestionIndex];
  if (!question) {
    finishQuiz();
    return;
  }

  progress.textContent = `Questao ${state.currentQuestionIndex + 1} de ${activeQuestions.length}`;
  questionTitle.textContent = question.title;
  questionDescription.textContent = question.description || "";

  const existingImage = document.querySelector(".question-inline-image");
  if (existingImage) {
    existingImage.remove();
  }

  if (question.image) {
    const imgEl = document.createElement("img");
    imgEl.className = "question-inline-image";
    imgEl.src = question.image;
    imgEl.alt = "Imagem da questao";
    imgEl.title = "Clique para ampliar a imagem";
    imgEl.setAttribute("role", "button");
    imgEl.tabIndex = 0;
    imgEl.setAttribute("aria-label", "Clique para ampliar a imagem da questão");
    const openViewer = () => openQuestionImageViewer(question.image, question.title || "Imagem da questão");
    imgEl.addEventListener("click", openViewer);
    imgEl.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        openViewer();
      }
    });
    questionDescription.parentNode.insertBefore(imgEl, questionForm);
  }

  questionForm.innerHTML = "";
  renderQuestionTimer(question);

  const alternativesVisibilityMode = quiz.alternativesVisibilityMode === "hidden" ? "hidden" : "revealed";
  const startsRevealed = alternativesVisibilityMode === "revealed";
  questionForm.classList.toggle("masked", !startsRevealed);
  question.options.forEach((option) => {
    const id = `${question.id} -${option.value} `;
    const label = document.createElement("label");
    label.className = "answer-option";
    label.setAttribute("for", id);

    const input = document.createElement("input");
    input.type = "radio";
    input.name = question.id;
    input.id = id;
    input.value = option.value;
    input.required = true;
    input.checked = state.answers[question.id] === option.value;
    input.addEventListener("change", () => {
      state.answers[question.id] = input.value;
      if (startsRevealed) {
        questionForm.classList.add("masked");
      }
      updateNextButtonVisibility();
    });

    const content = document.createElement("div");
    content.className = "answer-option-content";

    const placeholders = document.createElement("div");
    placeholders.className = "answer-placeholders";

    // Determine if the option text will wrap to two lines by measuring
    // its rendered width against the available placeholder width.
    let lineCount = 1;
    try {
      const available = Math.max(120, questionForm.clientWidth * 0.74);
      const canvas = document.createElement("canvas");
      const ctx = canvas.getContext("2d");
      const cs = window.getComputedStyle(questionForm || document.body);
      // prefer the full font shorthand if available
      ctx.font = cs.font || `${cs.fontSize} ${cs.fontFamily} `;
      const measured = ctx.measureText(option.label || "").width;
      if (measured > available) lineCount = 2;
    } catch (e) {
      // fallback to simple length heuristic
      lineCount = (option.label || "").length > 52 ? 2 : 1;
    }

    const primaryBlock = document.createElement("span");
    primaryBlock.className = "answer-placeholder-block";
    const primaryWidth = Math.floor(Math.random() * 24) + 46; // 46% a 69%
    primaryBlock.style.width = `${primaryWidth}% `;
    placeholders.appendChild(primaryBlock);

    if (lineCount > 1) {
      const secondaryBlock = document.createElement("span");
      secondaryBlock.className = "answer-placeholder-block answer-placeholder-block-secondary";
      const secondaryWidth = Math.floor(Math.random() * 22) + 30; // 30% a 51%
      secondaryBlock.style.width = `${secondaryWidth}% `;
      placeholders.appendChild(secondaryBlock);
    }

    const text = document.createElement("span");
    text.className = "answer-text";
    text.textContent = option.label;
    text.dataset.label = option.label;

    content.appendChild(placeholders);
    content.appendChild(text);
    label.appendChild(input);
    label.appendChild(content);
    questionForm.appendChild(label);
  });

  if (startsRevealed && state.answers[question.id]) {
    questionForm.classList.add("masked");
  }

  updateNextButtonVisibility();

  // Adiciona listener para efeito de "máscara circular"
  if (!questionForm._maskHandler) {
    questionForm._maskHandler = function (e) {
      const hoveredOption = e.target.closest(".answer-option");

      Array.from(questionForm.children).forEach((el) => {
        const textEl = el.querySelector(".answer-text");
        if (!textEl) {
          return;
        }

        if (el !== hoveredOption) {
          textEl.style.setProperty("--reveal-x", "-9999px");
          textEl.style.setProperty("--reveal-y", "-9999px");
          return;
        }

        const elRect = el.getBoundingClientRect();
        const localX = e.clientX - elRect.left;
        const localY = e.clientY - elRect.top;
        textEl.style.setProperty("--reveal-x", `${localX} px`);
        textEl.style.setProperty("--reveal-y", `${localY} px`);
      });
    };
    questionForm.addEventListener("mousemove", questionForm._maskHandler);
    questionForm.addEventListener("mouseleave", function () {
      Array.from(questionForm.children).forEach((el) => {
        const textEl = el.querySelector(".answer-text");
        if (textEl) {
          textEl.style.setProperty("--reveal-x", "-9999px");
          textEl.style.setProperty("--reveal-y", "-9999px");
        }
      });
    });
  }
  // fim renderQuestion
}

function handleNextQuestion() {
  advanceQuestion();
}

function handleSkipQuestion() {
  const quiz = getSelectedQuiz();
  if (!quiz || !state.isActive) {
    return;
  }

  if (state.timedOutQuestionId) {
    return;
  }

  if (Number(state.skipQuestionUsesRemaining || 0) <= 0) {
    showValidationMessage("Voce ja usou os 3 pulos deste questionario.");
    updateSkipButtonVisibility();
    return;
  }

  const activeQuestions = state.activeQuestions.length ? state.activeQuestions : quiz.questions;
  if (!Array.isArray(activeQuestions) || Number(state.currentQuestionIndex || 0) >= activeQuestions.length - 1) {
    return;
  }

  const currentIndex = Number(state.currentQuestionIndex || 0);
  const currentQuestion = activeQuestions[currentIndex];
  if (!currentQuestion) {
    return;
  }

  clearValidationMessage();
  const [movedQuestion] = activeQuestions.splice(currentIndex, 1);
  activeQuestions.push(movedQuestion);
  state.skipQuestionUsesRemaining = Math.max(0, Number(state.skipQuestionUsesRemaining || 0) - 1);
  renderQuestion();
}


async function finishQuiz() {
  // Bloqueia finalização se excedeu violações ou quiz foi cancelado
  if (!state.isTeacher && (state.violationCount >= 2 || state.isCancelled)) {
    alert("Você não pode mais concluir este questionário devido a violações de regras ou tempo esgotado.");
    state.isActive = false;
    state.isCancelled = true;
    showOnlyScreen("cancel");
    return;
  }

  const quiz = getSelectedQuiz();
  if (!quiz) {
    return;
  }

  const answersSnapshot = cloneAnswersSnapshot();
  const at = new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" });

  let result;
  let questionResults = [];
  let shouldPersistLocalFallback = false;
  try {
    const graded = await gradeQuizAttemptSecure(quiz, answersSnapshot, at);
    result = graded.result;
    questionResults = graded.questionResults;
  } catch (error) {
    console.error("Correção segura indisponível. Usando correção local ofuscada:", error);
    try {
      const local = gradeQuizAttemptObfuscated(quiz, answersSnapshot, at);
      result = local.result;
      questionResults = local.questionResults;
      shouldPersistLocalFallback = true;
    } catch (localError) {
      console.error("Falha também na correção local ofuscada:", localError);
      const secureMessage = getSecureGradingErrorMessage(error);
      const localMessage = getLocalGradingErrorMessage(localError);
      alert(`${localMessage} \n\nDetalhe técnico: ${secureMessage} `);
      return;
    }
  }

  state.isActive = false;
  state.timedOutQuestionId = null;
  state.isViolationGraceActive = false;
  clearQuestionTimerState();
  clearViolationTimers();
  hidePolicyWarning();

  result.questionResults = questionResults;
  attachAttemptTimingMetadata(result, at);

  saveStoredAttempt(quiz.id, {
    status: "completed",
    result,
    answers: answersSnapshot,
    questionResults
  });
  incrementKnownAttemptCount(quiz.id);

  saveAttemptToFirestore(quiz, result, answersSnapshot, questionResults);

  const baseMessage = result.earnedPoints === result.maxPoints
    ? "Parabéns, você gabaritou todas as questões!"
    : "Resultado registrado com sucesso.";
  showResult(result, baseMessage);
  updateReviewButtons(quiz, { status: "completed", result, answers: answersSnapshot, questionResults });
}

function showResult(result, detailsText, options = {}) {
  const safe = {
    studentName: "-",
    quizTitle: "-",
    quizId: "-",
    earnedPoints: 0,
    maxPoints: 0,
    percent: 0,
    date: "",
    focusExitPenaltyApplied: false,
    focusExitPenaltyReason: "",
    focusExitPenaltyAt: "",
    ...result
  };
  const isBlocked = Boolean(options.isBlocked);
  const blockedReason = String(options.reason || "").trim();
  const endedAt = safe.endedAt || safe.date || "-";
  const durationLabel = safe.durationLabel || formatDurationFromSeconds(safe.durationSeconds || 0);

  resultStudentName.textContent = safe.studentName;
  if (resultQuizTitle) {
    resultQuizTitle.textContent = `Avaliação: ${safe.quizTitle || safe.quizId || "-"}`;
  }
  resultScore.textContent = `${safe.earnedPoints} / ${safe.maxPoints}`;
  resultPercent.textContent = `${safe.percent}% de acertos`;
  const quiz = getSelectedQuiz();
  const focusPenaltyApplied = Boolean(safe.focusExitPenaltyApplied);
  const focusPenaltyReason = String(safe.focusExitPenaltyReason || "").trim();
  const focusPenaltyAt = String(safe.focusExitPenaltyAt || "").trim();

  const prevDisabled = document.getElementById("result-review-disabled");
  if (prevDisabled && prevDisabled.parentNode) {
    prevDisabled.parentNode.removeChild(prevDisabled);
  }

  const existingNote = document.getElementById("result-blocked-note");
  if (existingNote) {
    existingNote.remove();
  }

  resultDetails.innerHTML = "";

  const sectionLabel = document.createElement("p");
  sectionLabel.className = "result-meta-line result-meta-line-key";
  sectionLabel.textContent = "Resumo da tentativa";
  resultDetails.appendChild(sectionLabel);

  const statusLine = document.createElement("p");
  statusLine.className = "result-meta-line result-meta-line-strong";
  statusLine.textContent = isBlocked ? "Status: Avaliação bloqueada" : "Status: Avaliação concluída";
  resultDetails.appendChild(statusLine);

  const endLine = document.createElement("p");
  endLine.className = "result-meta-line";
  endLine.textContent = `Término: ${endedAt}`;
  resultDetails.appendChild(endLine);

  const durationLine = document.createElement("p");
  durationLine.className = "result-meta-line";
  durationLine.textContent = `Duração: ${durationLabel}`;
  resultDetails.appendChild(durationLine);

  const shouldShowWaitNonogram = Boolean(quiz && !isBlocked && (state.isTeacher || !quiz.allowFocusExitAfterFinish));

  if (!state.isTeacher && quiz && !isBlocked && !quiz.allowFocusExitAfterFinish) {
    const waitingLine = document.createElement("p");
    waitingLine.className = "result-meta-line result-meta-line-highlight";
    waitingLine.textContent = `${FOCUS_WAIT_PROCESS_NAME} ativo: aguarde o professor clicar em '${FOCUS_WAIT_RELEASE_BUTTON_LABEL}' para encerrar o modo de espera sem penalidade.`;
    resultDetails.appendChild(waitingLine);

    if (document.visibilityState === "visible" && !isFullscreenActive()) {
      triggerFocusWaitGrace("O modo foco exige tela cheia até o professor liberar o modo de espera.");
    }
  }

  renderWaitNonogramForWaitMode(quiz, shouldShowWaitNonogram);

  if (focusPenaltyApplied) {
    const penaltyLine = document.createElement("p");
    penaltyLine.className = "result-meta-line result-meta-line-highlight";
    const details = [
      "Penalidade registrada por sair do foco antes da liberação do professor.",
      focusPenaltyReason ? `Motivo: ${focusPenaltyReason}` : "",
      focusPenaltyAt ? `Quando: ${focusPenaltyAt}` : ""
    ].filter(Boolean).join(" ");
    penaltyLine.textContent = details;
    resultDetails.appendChild(penaltyLine);
  }

  if (isBlocked && blockedReason) {
    const reasonLine = document.createElement("p");
    reasonLine.className = "result-meta-line result-meta-line-highlight";
    reasonLine.textContent = `Motivo do bloqueio: ${blockedReason}`;
    resultDetails.appendChild(reasonLine);
  }

  if (detailsText) {
    const detailsLine = document.createElement("p");
    detailsLine.className = "result-meta-line";
    detailsLine.textContent = `Observação: ${detailsText}`;
    resultDetails.appendChild(detailsLine);
  }

  if (quiz && !quiz.allowStudentReview) {
    const note = document.createElement("p");
    note.id = "result-blocked-note";
    note.className = "result-blocked-note";
    note.textContent = "Resumo bloqueado no momento.";
    resultDetails.parentNode.appendChild(note);

    if (resultBackHomeButton && resultBackHomeButton.parentNode) {
      const disabledBtn = document.createElement("button");
      disabledBtn.type = "button";
      disabledBtn.id = "result-review-disabled";
      disabledBtn.className = "btn btn-disabled-gray btn-summary-blocked";
      disabledBtn.textContent = "Ver resumo";
      disabledBtn.title = "Atualizar";
      disabledBtn.addEventListener("mouseenter", () => {
        disabledBtn.textContent = "Atualizar";
      });
      disabledBtn.addEventListener("mouseleave", () => {
        disabledBtn.textContent = "Ver resumo";
      });
      disabledBtn.addEventListener("click", () => {
        window.location.reload();
      });
      resultBackHomeButton.parentNode.insertBefore(disabledBtn, resultBackHomeButton);
    }
  }

  const existingRetakeNote = document.getElementById("result-retake-note");
  if (existingRetakeNote) existingRetakeNote.remove();

  if (quiz && isRetakeReleased(quiz.id)) {
    const retakeNote = document.createElement("p");
    retakeNote.id = "result-retake-note";
    retakeNote.className = "result-retake-note";
    retakeNote.textContent = "Novo questionário liberado! Volte à home para iniciar uma nova tentativa.";
    resultDetails.parentNode.appendChild(retakeNote);
  }

  renderFocusWaitPersistentNotice();

  showOnlyScreen("result");
}

async function cancelQuiz(reason, options = {}) {
  const quiz = getSelectedQuiz();
  if (!quiz) {
    return;
  }

  const blockedByViolation = Boolean(options.blockedByViolation);
  const cancelAt = new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" });
  const answersSnapshot = cloneAnswersSnapshot();
  const cancelled = buildCancelledAttemptResult(quiz, answersSnapshot, cancelAt);
  const result = cancelled.result;
  const questionResults = cancelled.questionResults;

  state.isActive = false;
  state.isCancelled = true;
  state.timedOutQuestionId = null;
  state.isViolationGraceActive = false;
  clearQuestionTimerState();
  clearViolationTimers();
  hidePolicyWarning();
  hideCancelModal();

  const attemptPayload = {
    status: "cancelled",
    cancelReason: reason,
    blockedByViolation,
    at: cancelAt,
    result,
    answers: answersSnapshot,
    questionResults
  };

  saveStoredAttempt(quiz.id, attemptPayload);
  incrementKnownAttemptCount(quiz.id);
  if (blockedByViolation) {
    lockQuizAccess(quiz.id);
  }
  savePendingUnloadCancellation({
    quizId: quiz.id,
    quizTitle: quiz.title,
    questionsLength: quiz.questions.length,
    reason,
    blockedByViolation,
    at: cancelAt,
    result,
    answers: answersSnapshot,
    questionResults
  });

  if (!options.skipScreen) {
    showResult(result, "", { isBlocked: true, reason });
    updateReviewButtons(quiz, { status: "cancelled", result, answers: answersSnapshot, questionResults });
  }

  try {
    await saveCancelledAttemptToFirestore(quiz, {
      reason,
      blockedByViolation,
      at: cancelAt,
      result,
      answers: answersSnapshot,
      questionResults
    });
  } catch (error) {
    console.error("Erro ao persistir cancelamento do quiz:", error);
  }

  if (options.skipScreen) {
    return;
  }
}

async function saveCancelledAttemptToFirestore(quiz, cancellation) {
  if (!window.firebaseDB || !window.firebaseSetDoc || !window.firebaseDoc) {
    throw new Error("Firebase indisponível para salvar cancelamento.");
  }

  const uid = state.user?.uid || "anon";
  const slug = state.profileSlug || normalizeNameSlug(state.studentName || state.user?.displayName || "usuario");
  const safeQuizId = quiz.id.replace(/[^a-z0-9-_]/gi, "_");
  const timestamp = new Date().toISOString().replace(/[.:]/g, "-");
  const attemptId = `${safeQuizId}_${timestamp}`;

  // Garante que todos os campos necessários estejam presentes
  const result = cancellation.result || {
    quizId: quiz.id,
    quizTitle: quiz.title,
    studentName: state.studentName || "",
    earnedPoints: 0,
    maxPoints: (quiz.questions || []).length,
    percent: 0,
    date: cancellation.at || new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" })
  };
  const answersSnapshot = cancellation.answers || {};
  const questionResults = cancellation.questionResults || [];

  const userRef = window.firebaseDoc(window.firebaseDB, "usuarios", slug);
  const attemptRef = window.firebaseDoc(window.firebaseDB, "usuarios", slug, "tentativas", attemptId);
  const indexRef = window.firebaseDoc(window.firebaseDB, "usuarios_index", uid);

  const userPayload = {
    uid,
    nome: state.studentName || result.studentName || "",
    slug,
    email: state.user?.email || "",
    ultimaAtividade: result.date,
    atualizadoEmIso: new Date().toISOString()
  };

  const attemptPayload = {
    uid,
    quizId: quiz.id,
    quizTitulo: quiz.title,
    nome: result.studentName,
    pontos: result.earnedPoints,
    total: result.maxPoints,
    percentual: result.percent,
    data: result.date,
    inicioAvaliacao: result.startedAt || "",
    fimAvaliacao: result.endedAt || result.date || cancellation.at || "",
    duracaoSegundos: Number(result.durationSeconds || 0),
    duracaoTexto: result.durationLabel || formatDurationFromSeconds(result.durationSeconds || 0),
    status: "cancelled",
    bloqueadoPorViolacao: Boolean(cancellation.blockedByViolation),
    motivoCancelamento: cancellation.reason || "",
    penalidadeSaidaFoco: false,
    penalidadeSaidaFocoMotivo: "",
    penalidadeSaidaFocoEm: "",
    processoControleFoco: FOCUS_WAIT_PROCESS_NAME,
    respostas: answersSnapshot,
    correcaoQuestoes: Array.isArray(questionResults) ? questionResults : []
  };

  await Promise.all([
    window.firebaseSetDoc(userRef, userPayload, { merge: true }),
    window.firebaseSetDoc(attemptRef, attemptPayload),
    window.firebaseSetDoc(indexRef, {
      uid,
      nome: state.studentName || result.studentName || "",
      slug,
      email: state.user?.email || "",
      atualizadoEmIso: new Date().toISOString()
    }, { merge: true })
  ]);

  state.remoteAttemptsByQuiz[quiz.id] = {
    status: "cancelled",
    cancelReason: cancellation.reason,
    blockedByViolation: Boolean(cancellation.blockedByViolation),
    focusExitPenaltyApplied: false,
    focusExitPenaltyReason: "",
    focusExitPenaltyAt: "",
    result: {
      quizId: quiz.id,
      quizTitle: quiz.title,
      studentName: state.studentName || "",
      earnedPoints: result.earnedPoints,
      maxPoints: result.maxPoints,
      percent: result.percent,
      date: result.date,
      startedAt: result.startedAt || "",
      endedAt: result.endedAt || result.date || "",
      durationSeconds: Number(result.durationSeconds || 0),
      durationLabel: result.durationLabel || formatDurationFromSeconds(result.durationSeconds || 0),
      focusExitPenaltyApplied: false,
      focusExitPenaltyReason: "",
      focusExitPenaltyAt: "",
      questionResults
    },
    answers: answersSnapshot,
    questionResults
  };
  state.remoteAttemptsLoaded = true;
  clearPendingUnloadCancellation();
  renderHome();
}

function saveAttemptToFirestore(quiz, result, answers, questionResults = []) {
  if (!window.firebaseDB || !window.firebaseSetDoc || !window.firebaseDoc) {
    return;
  }

  const uid = state.user?.uid || "anon";
  const slug = state.profileSlug || normalizeNameSlug(state.studentName || state.user?.displayName || "usuario");
  const safeQuizId = quiz.id.replace(/[^a-z0-9-_]/gi, "_");
  const timestamp = new Date().toISOString().replace(/[.:]/g, "-");
  const attemptId = `${safeQuizId}_${timestamp}`;

  // Estrutura organizada por nome escolhido:
  // usuarios/{nomeSlug}
  // usuarios/{nomeSlug}/tentativas/{attemptId}
  const userRef = window.firebaseDoc(window.firebaseDB, "usuarios", slug);
  const attemptRef = window.firebaseDoc(window.firebaseDB, "usuarios", slug, "tentativas", attemptId);
  const indexRef = window.firebaseDoc(window.firebaseDB, "usuarios_index", uid);

  const userPayload = {
    uid,
    nome: state.studentName || result.studentName || "",
    slug,
    email: state.user?.email || "",
    ultimaAtividade: result.date,
    atualizadoEmIso: new Date().toISOString()
  };

  const attemptPayload = {
    uid,
    quizId: quiz.id,
    quizTitulo: quiz.title,
    nome: result.studentName,
    pontos: result.earnedPoints,
    total: result.maxPoints,
    percentual: result.percent,
    data: result.date,
    inicioAvaliacao: result.startedAt || "",
    fimAvaliacao: result.endedAt || result.date || "",
    duracaoSegundos: Number(result.durationSeconds || 0),
    duracaoTexto: result.durationLabel || formatDurationFromSeconds(result.durationSeconds || 0),
    status: "completed",
    bloqueadoPorViolacao: false,
    motivoCancelamento: "",
    penalidadeSaidaFoco: Boolean(result.focusExitPenaltyApplied),
    penalidadeSaidaFocoMotivo: String(result.focusExitPenaltyReason || ""),
    penalidadeSaidaFocoEm: String(result.focusExitPenaltyAt || ""),
    processoControleFoco: FOCUS_WAIT_PROCESS_NAME,
    respostas: answers,
    correcaoQuestoes: Array.isArray(questionResults) ? questionResults : []
  };

  syncAttemptCacheForQuiz(quiz.id, {
    status: "completed",
    result,
    answers,
    questionResults
  }, Number(getStoredAttemptCount(quiz.id) || 0));

  Promise.all([
    window.firebaseSetDoc(userRef, userPayload, { merge: true }),
    window.firebaseSetDoc(attemptRef, attemptPayload),
    window.firebaseSetDoc(indexRef, {
      uid,
      nome: state.studentName || result.studentName || "",
      slug,
      email: state.user?.email || "",
      atualizadoEmIso: new Date().toISOString()
    }, { merge: true })
  ]).catch((error) => {
    console.error("Erro ao salvar no Firestore:", error);
  });
}

function isFullscreenActive() {
  return Boolean(document.fullscreenElement);
}

let _zoomWatchMQ = null;
let _zoomWatchHandler = null;

function registerZoomWatch() {
  if (_zoomWatchMQ && _zoomWatchHandler) {
    _zoomWatchMQ.removeEventListener("change", _zoomWatchHandler);
  }
  const mqString = `(resolution: ${window.devicePixelRatio}dppx)`;
  _zoomWatchMQ = window.matchMedia(mqString);
  _zoomWatchHandler = () => {
    applyZoomCounterScale();
    registerZoomWatch();
  };
  _zoomWatchMQ.addEventListener("change", _zoomWatchHandler);
}

function applyZoomCounterScale() {
  if (!quizScreen) return;
  if (!state.isActive || state.isTeacher) {
    quizScreen.style.transform = "";
    return;
  }
  const zoomFactor = window.devicePixelRatio / (state.baselineDPR || window.devicePixelRatio);
  if (Math.abs(zoomFactor - 1) < 0.02) {
    quizScreen.style.transform = "";
    return;
  }
  quizScreen.style.transform = `scale(${1 / zoomFactor})`;
  quizScreen.style.transformOrigin = "top center";
}

function lockZoomForQuiz() {
  state.baselineDPR = window.devicePixelRatio;
  registerZoomWatch();
  applyZoomCounterScale();
}

function handlePolicyViolation() {
  if (!state.isActive || state.isTeacher || state.isViolationGraceActive) {
    return;
  }

  if (state.violationCount >= 2) {
    cancelQuiz("Questionário bloqueado após múltiplas violações das regras.", { blockedByViolation: true });
    return;
  }

  state.violationCount += 1;
  state.isViolationGraceActive = true;
  state.violationDeadline = Date.now() + 10000;

  showPolicyWarning();
  clearViolationTimers();
  updatePolicyCountdown();
  state.violationIntervalId = setInterval(updatePolicyCountdown, 100);
  state.violationTimeoutId = setTimeout(() => {
    if (state.isActive && state.isViolationGraceActive) {
      cancelQuiz("Questionário bloqueado porque você não retornou em até 10 segundos.", { blockedByViolation: true });
    }
  }, 10000);
}

async function handleReturnToQuiz() {
  if (!state.isActive || state.isCancelled) {
    showOnlyScreen("cancel");
    return;
  }

  if (!isFullscreenActive()) {
    const entered = await requestFullscreenMode();
    if (!entered) {
      return;
    }
  }

  state.isViolationGraceActive = false;
  state.violationDeadline = 0;
  clearViolationTimers();
  hidePolicyWarning();
}

function clearViolationTimers() {
  if (state.violationTimeoutId) {
    clearTimeout(state.violationTimeoutId);
    state.violationTimeoutId = null;
  }

  if (state.violationIntervalId) {
    clearInterval(state.violationIntervalId);
    state.violationIntervalId = null;
  }
}

function showPolicyWarning() {
  if (!policyOverlay || !policyChancesText) {
    return;
  }

  const changesMessage = state.violationCount === 1
    ? "Não abra mais. Esta é sua 1a de 2 chances."
    : "Não abra mais. Esta é sua ultima chance.";

  policyChancesText.textContent = changesMessage;
  quizScreen.classList.add("quiz-locked");
  policyOverlay.classList.remove("hidden");
}

function hidePolicyWarning() {
  if (!policyOverlay || !policyChancesText) {
    return;
  }

  policyCountdown.textContent = "10";
  policyChancesText.textContent = "";
  quizScreen.classList.remove("quiz-locked");
  policyOverlay.classList.add("hidden");
}

function updatePolicyCountdown() {
  if (!state.isViolationGraceActive || !state.violationDeadline) {
    return;
  }

  const remainingMs = Math.max(0, state.violationDeadline - Date.now());
  policyCountdown.textContent = String(Math.ceil(remainingMs / 1000));

  if (remainingMs === 0 && state.isViolationGraceActive) {
    cancelQuiz("Questionário bloqueado porque você não retornou em até 10 segundos.", { blockedByViolation: true });
  }
}

async function requestFullscreenMode() {
  if (isFullscreenActive()) {
    return true;
  }

  if (!document.documentElement.requestFullscreen) {
    return false;
  }

  try {
    await document.documentElement.requestFullscreen();
    return isFullscreenActive();
  } catch (error) {
    console.error(error);
    return false;
  }
}

function showValidationMessage(message) {
  clearValidationMessage();
  const warning = document.createElement("p");
  warning.id = "question-validation-message";
  warning.textContent = message;
  warning.style.color = "#b42318";
  warning.style.margin = "8px 0 0";
  warning.style.fontWeight = "700";
  questionForm.appendChild(warning);
}

function clearValidationMessage() {
  const existing = document.getElementById("question-validation-message");
  if (existing) {
    existing.remove();
  }
}

function clampQuestionImageZoom(value) {
  return Math.min(questionImageZoomState.maxScale, Math.max(questionImageZoomState.minScale, value));
}

function applyQuestionImageZoomTransform() {
  if (!questionImageViewerImage) {
    return;
  }

  questionImageViewerImage.style.transform = `translate(${questionImageZoomState.translateX}px, ${questionImageZoomState.translateY}px) scale(${questionImageZoomState.scale})`;
}

function resetQuestionImageZoom() {
  questionImageZoomState.scale = 1;
  questionImageZoomState.translateX = 0;
  questionImageZoomState.translateY = 0;
  applyQuestionImageZoomTransform();
}

function updateQuestionImageZoom(delta, anchorX = 0.5, anchorY = 0.5) {
  const previousScale = questionImageZoomState.scale;
  const nextScale = clampQuestionImageZoom(previousScale + delta);
  if (nextScale === previousScale) {
    return;
  }

  const factor = nextScale / previousScale;
  const stageRect = questionImageViewerStage?.getBoundingClientRect();
  const stageWidth = stageRect?.width || 1;
  const stageHeight = stageRect?.height || 1;
  const offsetX = (anchorX - 0.5) * stageWidth;
  const offsetY = (anchorY - 0.5) * stageHeight;

  questionImageZoomState.translateX = (questionImageZoomState.translateX - offsetX) * factor + offsetX;
  questionImageZoomState.translateY = (questionImageZoomState.translateY - offsetY) * factor + offsetY;
  questionImageZoomState.scale = nextScale;
  applyQuestionImageZoomTransform();
}

function openQuestionImageViewer(src, title = "Imagem da questão") {
  if (!questionImageViewer || !questionImageViewerImage || !src) {
    return;
  }

  questionImageZoomState = {
    src,
    scale: 1,
    minScale: 1,
    maxScale: 4,
    translateX: 0,
    translateY: 0,
    dragging: false,
    dragStartX: 0,
    dragStartY: 0,
    startTranslateX: 0,
    startTranslateY: 0
  };

  questionImageViewerImage.src = src;
  questionImageViewerImage.alt = title;
  questionImageViewer.classList.remove("hidden");
  document.body.classList.add("question-image-viewer-open");
  if (questionImageViewerTitle) {
    questionImageViewerTitle.textContent = title;
  }
  if (questionImageViewerCaption) {
    questionImageViewerCaption.textContent = "Use a roda do mouse, os botões ou arraste a imagem.";
  }
  resetQuestionImageZoom();
}

function hideQuestionImageViewer() {
  if (!questionImageViewer || !questionImageViewerImage) {
    return;
  }

  questionImageViewer.classList.add("hidden");
  document.body.classList.remove("question-image-viewer-open");
  questionImageViewerImage.removeAttribute("src");
  questionImageZoomState.dragging = false;
}

if (questionImageViewerClose) {
  questionImageViewerClose.addEventListener("click", hideQuestionImageViewer);
}

if (questionImageViewer) {
  questionImageViewer.addEventListener("click", (event) => {
    if (event.target?.matches?.("[data-close-question-image-viewer]")) {
      hideQuestionImageViewer();
    }
  });
}

if (questionImageZoomOut) {
  questionImageZoomOut.addEventListener("click", () => updateQuestionImageZoom(-0.25));
}

if (questionImageZoomReset) {
  questionImageZoomReset.addEventListener("click", resetQuestionImageZoom);
}

if (questionImageZoomIn) {
  questionImageZoomIn.addEventListener("click", () => updateQuestionImageZoom(0.25));
}

if (questionImageViewerStage) {
  questionImageViewerStage.addEventListener("wheel", (event) => {
    if (questionImageViewer.classList.contains("hidden")) {
      return;
    }

    event.preventDefault();
    const delta = event.deltaY > 0 ? -0.2 : 0.2;
    const rect = questionImageViewerStage.getBoundingClientRect();
    const anchorX = rect.width ? (event.clientX - rect.left) / rect.width : 0.5;
    const anchorY = rect.height ? (event.clientY - rect.top) / rect.height : 0.5;
    updateQuestionImageZoom(delta, anchorX, anchorY);
  }, { passive: false });

  questionImageViewerStage.addEventListener("pointerdown", (event) => {
    if (questionImageViewer.classList.contains("hidden") || !questionImageViewerImage || questionImageZoomState.scale <= 1) {
      return;
    }

    questionImageZoomState.dragging = true;
    questionImageZoomState.dragStartX = event.clientX;
    questionImageZoomState.dragStartY = event.clientY;
    questionImageZoomState.startTranslateX = questionImageZoomState.translateX;
    questionImageZoomState.startTranslateY = questionImageZoomState.translateY;
    questionImageViewerStage.classList.add("is-dragging");
    questionImageViewerStage.setPointerCapture(event.pointerId);
  });

  questionImageViewerStage.addEventListener("pointermove", (event) => {
    if (!questionImageZoomState.dragging) {
      return;
    }

    const deltaX = event.clientX - questionImageZoomState.dragStartX;
    const deltaY = event.clientY - questionImageZoomState.dragStartY;
    questionImageZoomState.translateX = questionImageZoomState.startTranslateX + deltaX;
    questionImageZoomState.translateY = questionImageZoomState.startTranslateY + deltaY;
    applyQuestionImageZoomTransform();
  });

  const stopDragging = () => {
    questionImageZoomState.dragging = false;
    questionImageViewerStage.classList.remove("is-dragging");
  };

  questionImageViewerStage.addEventListener("pointerup", stopDragging);
  questionImageViewerStage.addEventListener("pointercancel", stopDragging);
  questionImageViewerStage.addEventListener("pointerleave", stopDragging);
}


