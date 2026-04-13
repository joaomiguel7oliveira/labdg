const state = {
  user: null,
  isTeacher: false,
  profileSlug: "",
  classroomId: "",
  classroomName: "",
  classrooms: [],
  classroomStudents: {},
  customQuizzes: [],
  quizConfigs: {},
  remoteAttemptsByQuiz: {},
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
  isRenamingProfile: false,
  reviewBackScreen: "result",
  violationCount: 0,
  isViolationGraceActive: false,
  violationTimeoutId: null,
  violationIntervalId: null,
  violationDeadline: 0
};

let pendingModalAction = null;
let lastHomeRenderSignature = "";

const authScreen = document.getElementById("auth-screen");
const bootScreen = document.getElementById("boot-screen");
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
const teacherRefreshButton = document.getElementById("teacher-refresh-button");
const teacherUsersList = document.getElementById("teacher-users-list");
const teacherMessage = document.getElementById("teacher-message");
const builderSubmitButton = document.getElementById("builder-submit-button");
const builderForm = document.getElementById("builder-form");
function getPendingUnloadKey() {
  return `quiz_pending_unload_cancellation_${state.user?.uid || "anon"}`;
}
const builderTitleInput = document.getElementById("builder-title");
const builderDescriptionInput = document.getElementById("builder-description");
const builderDurationInput = document.getElementById("builder-duration");
const builderQuestions = document.getElementById("builder-questions");
const builderAddQuestionButton = document.getElementById("builder-add-question-button");
const builderMessage = document.getElementById("builder-message");
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
const cancelButton = document.getElementById("cancel-button");
const cancelModal = document.getElementById("cancel-modal");
const cancelModalTitle = document.getElementById("cancel-modal-title");
const cancelModalText = document.getElementById("cancel-modal-text");
const cancelModalConfirm = document.getElementById("cancel-modal-confirm");
const cancelModalCancel = document.getElementById("cancel-modal-cancel");
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
    clearValidationMessage();
    updateNextButtonVisibility();
  });
}
cancelButton.addEventListener("click", () => {
  openActionModal({
    title: "Cancelar questionário?",
    text: "Se você cancelar, não terá mais acesso a este quiz.",
    confirmLabel: "Sim, cancelar",
    onConfirm: async () => {
      await cancelQuiz("Questionário bloqueado ao cancelar a avaliação.", { blockedByViolation: true });
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
      const base = attempt?.cancelReason || "Questionário cancelado por violação de regras.";
      if (quiz && !quiz.allowStudentReview) {
        cancelMessage.innerHTML = getBlockedReviewNoteHtml(base);
      } else {
        cancelMessage.textContent = base;
      }
      if (quiz && attempt) {
        updateReviewButtons(quiz, attempt);
      }
      showOnlyScreen("cancel");
      return;
    }

    if (attempt?.result) {
      showResult(attempt.result, "Resultado carregado.");
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
        text: "Ao voltar para a home, este questionário será cancelado e você perderá o acesso a ele.",
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
        text: "Se você continuar, o questionário atual será cancelado e você perderá o acesso a ele.",
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

logoutButton.addEventListener("click", async () => {
  if (state.isActive && getSelectedQuiz()) {
    if (state.isTeacher) {
      if (typeof window.firebaseSignOut === "function") await window.firebaseSignOut();
      return;
    }
    openActionModal({
      title: "Sair da plataforma?",
      text: "Ao sair agora, este questionário será cancelado e você perderá o acesso a ele.",
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
  }
});

window.addEventListener("blur", () => {
  setTimeout(() => {
    if (state.isActive && !state.isTeacher && !document.hasFocus()) {
      handlePolicyViolation();
    }
  }, 0);
});

document.addEventListener("fullscreenchange", () => {
  if (state.isActive && !state.isTeacher && !isFullscreenActive()) {
    handlePolicyViolation();
  }
});

window.addEventListener("beforeunload", (event) => {
  if (!state.isActive || state.isTeacher || !getSelectedQuiz()) {
    return;
  }

  persistUnloadCancellation();
  event.preventDefault();
  event.returnValue = "";
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

function clearPendingUnloadCancellation() {
  localStorage.removeItem(getPendingUnloadKey());
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
  const reason = "Questionário cancelado ao recarregar ou fechar a página.";
  const answersSnapshot = cloneAnswersSnapshot();
  const cancelled = buildCancelledAttemptResult(quiz, answersSnapshot, cancelAt);
  const result = cancelled.result;
  const questionResults = cancelled.questionResults;

  // Salva localmente
  saveStoredAttempt(quiz.id, {
    status: "cancelled",
    cancelReason: reason,
    blockedByViolation: false,
    at: cancelAt,
    result,
    answers: answersSnapshot,
    questionResults
  });
  // Salva no Firestore para o painel do professor
  saveCancelledAttemptToFirestore(quiz, {
    reason,
    blockedByViolation: false,
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
    blockedByViolation: false,
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
        allowStudentStart: cfg.allowStudentStart ?? false,
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
      options: options.map(opt => ({ value: String(opt.value), label: String(opt.label || "") }))
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
        points: Number(question.points || 1),
        timer: Number(question.timer || 0),
        options,
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
    allowStudentReview: Boolean(rawQuiz.allowStudentReview),
    questions: normalizedQuestions,
    isCustom: true
  };
}

async function refreshCustomQuizzes(options = {}) {
  if (!window.firebaseDB || !window.firebaseCollection || !window.firebaseGetDocs) {
    state.customQuizzes = [];
    return;
  }

  try {
    const customRef = window.firebaseCollection(window.firebaseDB, "quizzes_custom");
    const snap = await window.firebaseGetDocs(customRef);
    const nextCustom = [];

    snap.forEach((docSnap) => {
      const data = docSnap.data();
      if (data?.ativo === false) {
        return;
      }
      const normalized = normalizeCustomQuiz(data, docSnap.id);
      if (normalized) {
        nextCustom.push(normalized);
      }
    });

    state.customQuizzes = nextCustom;
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
      if (window.firebaseDeleteDoc) {
        await window.firebaseDeleteDoc(quizRef);
      } else {
        await window.firebaseSetDoc(quizRef, { ativo: false }, { merge: true });
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
  refreshRemoteAttempts();
  refreshQuizConfigs();
  refreshRetakeReleases();
  refreshCustomQuizzes();
  refreshClassrooms();
  refreshTeacherAccess();
}

function showUnauthenticatedUI() {
  endBootPhase();
  state.isTeacher = false;
  state.studentName = "";
  state.profileSlug = "";
  state.classroomId = "";
  state.classroomName = "";
  state.classrooms = [];
  state.quizConfigs = {};
  state.remoteAttemptsByQuiz = {};
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
  logoutButton.classList.add("hidden");
  showOnlyScreen(null);
}

async function refreshTeacherAccess(options = {}) {
  state.isTeacher = false;
  teacherPanelButton.classList.add("hidden");
  if (teacherClassroomsButton) {
    teacherClassroomsButton.classList.add("hidden");
  }
  if (teacherBuilderButton) {
    teacherBuilderButton.classList.add("hidden");
  }

  if (!state.user?.email || !window.firebaseDoc || !window.firebaseGetDoc || !window.firebaseDB) {
    return;
  }

  const emailKey = state.user.email.trim().toLowerCase();

  try {
    const teacherRef = window.firebaseDoc(window.firebaseDB, "professores_permitidos", emailKey);
    const teacherSnap = await window.firebaseGetDoc(teacherRef);
    const allowed = teacherSnap.exists() && teacherSnap.data()?.ativo !== false;
    state.isTeacher = Boolean(allowed);

    if (state.isTeacher) {
      teacherPanelButton.classList.remove("hidden");
      if (teacherClassroomsButton) {
        teacherClassroomsButton.classList.remove("hidden");
      }
      if (teacherBuilderButton) {
        teacherBuilderButton.classList.remove("hidden");
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
      if (state.remoteAttemptsByQuiz) {
        delete state.remoteAttemptsByQuiz[quizId];
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
      result: {
        quizId: latest.quizId,
        quizTitle: latest.quizTitulo,
        studentName: latest.nome,
        earnedPoints: latest.pontos ?? 0,
        maxPoints: latest.total ?? 0,
        percent: latest.percentual ?? 0,
        date: latest.data,
        questionResults: latest.correcaoQuestoes || []
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
    await refreshClassrooms();
    const usersRef = window.firebaseCollection(window.firebaseDB, "usuarios");
    const usersSnap = await window.firebaseGetDocs(usersRef);

    if (usersSnap.empty) {
      teacherMessage.textContent = "Nenhum aluno encontrado.";
      return;
    }

    const groups = {};
    for (const userDoc of usersSnap.docs) {
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
            ? ((attempt.blockedByViolation || attempt.bloqueadoPorViolacao) ? "Bloqueado por violação" : "Cancelado")
            : "Concluído"}
                  </span>
                  ${attempt.status === "cancelled" && (attempt.motivoCancelamento || attempt.cancelReason)
            ? `<small class="teacher-status-note ${(attempt.blockedByViolation || attempt.bloqueadoPorViolacao) ? "teacher-status-note-danger" : "teacher-status-note-muted"}">${attempt.motivoCancelamento || attempt.cancelReason}</small>`
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

      if (!groups[turmaNome]) {
        groups[turmaNome] = [];
      }
      groups[turmaNome].push(userCard);
    }

    const orderedTurmas = Object.keys(groups).sort((a, b) => {
      if (a === "Sem turma") return 1;
      if (b === "Sem turma") return -1;
      return a.localeCompare(b, "pt-BR");
    });

    teacherUsersList.innerHTML = orderedTurmas
      .map((turmaNome) => `
        <section class="teacher-classroom-section">
          <h3>Turma: ${turmaNome}</h3>
          <div class="teacher-users-list">
            ${groups[turmaNome].join("")}
          </div>
        </section>
      `)
      .join("");

    teacherMessage.textContent = "";

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
  } catch (error) {
    teacherMessage.textContent = "Erro ao carregar painel do professor.";
    console.error(error);
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
    <input type="text" placeholder="Texto da alternativa" value="${value.replace(/"/g, "&quot;")}" required />
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

  const seedOptions = Array.isArray(seed.options) && seed.options.length >= 2
    ? seed.options
    : [{ label: "" }, { label: "" }, { label: "" }, { label: "" }];

  seedOptions.forEach((option) => addBuilderOptionRow(optionsEl, option.label || "", option.value === seed.correctAnswer));

  builderQuestions.appendChild(card);
  updateBuilderQuestionTitles();
  updateBuilderRadioGroups();
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
  if (builderForm) {
    builderForm.reset();
  }
  if (builderQuestions) {
    builderQuestions.innerHTML = "";
  }
  if (builderSubmitButton) {
    builderSubmitButton.textContent = "Publicar quiz";
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
    const points = Number(card.querySelector(".builder-question-points")?.value || 1);
    const optionRows = Array.from(card.querySelectorAll(".builder-option-row"));
    const timer = Number(card.querySelector(".builder-question-timer")?.value || 0);

    if (!title || optionRows.length < 2) {
      return { error: "Cada questão precisa de enunciado e pelo menos 2 alternativas." };
    }

    const options = [];
    let correctAnswer = "";

    optionRows.forEach((row, index) => {
      const label = row.querySelector("input[type='text']")?.value.trim() || "";
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
      points: Number.isFinite(points) && points > 0 ? points : 1,
      options,
      timer: Number.isFinite(timer) && timer > 0 ? timer : 0,
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
      points: question.points,
      options: question.options,
      timer: question.timer,
      answerToken: buildAnswerToken(quizId, questionId, correctAnswer)
    };
  });

  const payload = {
    title,
    description,
    duration,
    questions: questionsWithTokens,
    ativo: true,
    criadoPorUid: state.user?.uid || "",
    criadoPorEmail: state.user?.email || "",
    criadoEmIso: new Date().toISOString(),
    atualizadoEmIso: new Date().toISOString()
  };

  if (builderMessage) {
    builderMessage.textContent = state.editingQuizId ? "Salvando alteracoes..." : "Publicando quiz...";
  }

  try {
    const quizRef = window.firebaseDoc(window.firebaseDB, "quizzes_custom", quizId);
    const answersRef = window.firebaseDoc(window.firebaseDB, "quizzes_answer_keys", quizId);
    await Promise.all([
      window.firebaseSetDoc(quizRef, payload),
      window.firebaseSetDoc(answersRef, {
        quizId,
        answers: built.answerKey,
        atualizadoEmIso: new Date().toISOString(),
        atualizadoPorUid: state.user?.uid || "",
        atualizadoPorEmail: state.user?.email || ""
      }, { merge: true })
    ]);
    if (builderMessage) {
      builderMessage.textContent = state.editingQuizId ? "Quiz atualizado com sucesso." : "Quiz publicado com sucesso.";
    }
    await refreshCustomQuizzes();
    resetBuilderForm();
  } catch (error) {
    if (builderMessage) {
      builderMessage.textContent = state.editingQuizId ? "Erro ao atualizar quiz." : "Erro ao publicar quiz.";
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

function getStoredAttempt(quizId) {
  const raw = localStorage.getItem(getStorageKey(quizId));
  return raw ? JSON.parse(raw) : null;
}

function getKnownAttempt(quizId) {
  // Sempre prioriza o estado remoto, se carregado
  if (state.remoteAttemptsLoaded) {
    // Se não existe tentativa remota, remove localStorage para garantir sincronização
    if (!state.remoteAttemptsByQuiz[quizId]) {
      localStorage.removeItem(getStorageKey(quizId));
      return null;
    }
    return state.remoteAttemptsByQuiz[quizId];
  }
  // Se não carregou remoto ainda, usa local
  const localAttempt = getStoredAttempt(quizId);
  return localAttempt || null;
}

function saveStoredAttempt(quizId, payload) {
  localStorage.setItem(getStorageKey(quizId), JSON.stringify(payload));
  if (state.remoteAttemptsLoaded) {
    state.remoteAttemptsByQuiz[quizId] = payload;
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

  return {
    result: {
      quizId: quiz?.id,
      quizTitle: quiz?.title,
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

  const byId = item?.questionId
    ? (quiz?.questions || []).find((question) => String(question?.id || "") === String(item.questionId))
    : null;
  const byIndex = (quiz?.questions || [])[index];
  const question = byId || byIndex;
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

    snap.forEach((docSnap) => {
      const data = docSnap.data();
      if (!data?.quizId) {
        return;
      }

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
        result: {
          quizId: data.quizId,
          quizTitle: data.quizTitulo,
          studentName: data.nome,
          earnedPoints: data.pontos ?? 0,
          maxPoints: data.total ?? 0,
          percent: data.percentual ?? 0,
          date: data.data,
          questionResults: data.correcaoQuestoes || []
        },
        answers: data.respostas || {},
        questionResults: data.correcaoQuestoes || []
      };
    });

    state.remoteAttemptsByQuiz = nextMap;
    state.remoteAttemptsLoaded = true;

    getAllQuizzes().forEach((quiz) => {
      const storageKey = getStorageKey(quiz.id);
      if (nextMap[quiz.id]) {
        localStorage.setItem(storageKey, JSON.stringify(nextMap[quiz.id]));
      } else {
        localStorage.removeItem(storageKey);
      }
    });

    if (options.render !== false) {
      renderHome();
    }
  } catch (error) {
    console.error("Erro ao carregar tentativas remotas:", error);
  }
}

function renderHome() {
  const allQuizzes = getAllQuizzes();
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
    // Se o professor liberou nova tentativa, remove status cancelado/completed
    if (isRetakeReleased(quiz.id)) {
      // Remove do localStorage e do estado
      localStorage.removeItem(getStorageKey(quiz.id));
      if (state.remoteAttemptsByQuiz) delete state.remoteAttemptsByQuiz[quiz.id];
      attempt = null;
    }
    const canOpenQuiz = state.isTeacher || Boolean(attempt) || Boolean(quiz.allowStudentStart) || isRetakeReleased(quiz.id);
    const mainButtonLabel = attempt ? "Ver detalhes" : (canOpenQuiz ? "Abrir quiz" : "Início bloqueado");
    const status = attempt?.status || (isRetakeReleased(quiz.id) ? "available" : "available");
    const statusLabel =
      status === "completed" ? "Concluído" :
        status === "cancelled" ? "Cancelado" :
          "Disponível";

    signatureParts.push([
      quiz.id,
      status,
      canOpenQuiz ? "1" : "0",
      quiz.allowStudentStart ? "1" : "0",
      quiz.allowStudentReview ? "1" : "0",
      attempt?.result?.date || "",
      state.isTeacher ? "1" : "0"
    ].join("|"));

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
      </ul>
      <div class="quiz-card-actions">
        <button class="btn btn-primary btn-sm ${canOpenQuiz ? "" : "btn-start-blocked"}" data-quiz="${quiz.id}" title="${canOpenQuiz ? "" : "Atualizar"}">
          ${mainButtonLabel}
        </button>
        ${attempt?.result ? (quiz.allowStudentReview ? `<button class="btn btn-primary btn-sm" data-quick-review="${quiz.id}">Resumo</button>` : `<button class="btn btn-primary btn-sm btn-summary-blocked" data-summary-blocked="${quiz.id}" title="Atualizar">Resumo</button>`) : ""}
        ${state.isTeacher ? `<button class="btn btn-ghost btn-sm" data-toggle-start="${quiz.id}">${quiz.allowStudentStart ? "Bloquear início" : "Liberar início"}</button>` : ""}
        ${state.isTeacher ? `<button class="btn btn-ghost btn-sm" data-toggle-review="${quiz.id}">${quiz.allowStudentReview ? "Bloquear resumo" : "Liberar resumo"}</button>` : ""}
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
      ${attempt && !quiz.allowStudentReview ? `<p class="quiz-review-note">Resumo bloqueado no momento.</p>` : ""}
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

        if (attempt) {
          // Show result screen for this attempt
          state.selectedQuizId = quiz.id;
          const syncedAttempt = await syncLatestAttemptForQuiz(quiz.id);
          if (!syncedAttempt?.result) {
            return;
          }
          showResult(syncedAttempt.result, "Resultado carregado.");
          updateReviewButtons(quiz, syncedAttempt);
          showOnlyScreen("result");
        } else {
          await openQuiz(quiz.id);
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

    fragment.appendChild(card);
  });

  const nextSignature = `${filterText}|${state.isTeacher ? "teacher" : "student"}|${signatureParts.join("||")}`;
  if (lastHomeRenderSignature === nextSignature) {
    return;
  }

  quizGrid.replaceChildren(fragment);
  lastHomeRenderSignature = nextSignature;
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
      result: {
        quizId: latest.quizId,
        quizTitle: latest.quizTitulo,
        studentName: latest.nome,
        earnedPoints: latest.pontos ?? 0,
        maxPoints: latest.total ?? 0,
        percent: latest.percentual ?? 0,
        date: latest.data,
        questionResults: latest.correcaoQuestoes || []
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

  if (!retakeReleased && attempt?.status === "completed") {
    resultAlreadyCompleted.classList.remove("hidden");
    showResult(attempt.result, "Você já realizou este questionário. Este é seu resultado anterior.");
    updateReviewButtons(quiz, attempt);
    return;
  }

  if (!retakeReleased && attempt?.status === "cancelled") {
    const base = attempt.cancelReason || "Questionário cancelado por violação de regras.";
    if (!quiz.allowStudentReview) {
      cancelMessage.innerHTML = getBlockedReviewNoteHtml(base);
    } else {
      cancelMessage.textContent = base;
    }
    showOnlyScreen("cancel");
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
    `<li>Total de questões: ${quiz.questions.length}</li>`;

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
      `<li>As questões e alternativas são embaralhadas.</li>` +
      `<li>Não recarregue a página e não troque de tela; isso cancela a avaliação.</li>` +
      `<li>Ao cancelar o quiz, você perde o acesso à tentativa.</li>` +
      (isStartAllowed ? "" : "<li><strong>Início bloqueado pelo professor no momento.</strong></li>");
  }

  if (startQuizButton) {
    startQuizButton.disabled = !isStartAllowed;
    startQuizButton.textContent = isStartAllowed ? "Iniciar questionário" : "Início bloqueado";
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
  await Promise.allSettled([
    refreshQuizConfigs({ render: false }),
    refreshCustomQuizzes({ render: false }),
    refreshRetakeReleases({ render: false }),
    refreshRemoteAttempts({ render: false }),
    refreshTeacherAccess({ render: false })
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
  state.isActive = false;
  state.isCancelled = false;
  state.timedOutQuestionId = null;
  state.reviewBackScreen = "result";
  state.violationCount = 0;
  state.isViolationGraceActive = false;
  clearQuestionTimerState();
  clearViolationTimers();
  hidePolicyWarning();
  hideCancelModal();
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

function updateNextButtonVisibility() {
  if (!nextButton) {
    return;
  }

  const selected = questionForm?.querySelector("input[type='radio']:checked");
  const shouldShow = Boolean(selected) && !state.timedOutQuestionId;
  nextButton.classList.toggle("hidden", !shouldShow);
  if (clearAnswerButton) {
    clearAnswerButton.classList.toggle("hidden", !shouldShow);
  }
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

function clearQuestionTimerState() {
  if (window.questionTimerInterval) {
    clearInterval(window.questionTimerInterval);
    window.questionTimerInterval = null;
  }
  if (questionInfoArea) {
    questionInfoArea.innerHTML = "";
  }
  if (questionForm) {
    questionForm.classList.remove("question-timeout-locked");
  }
  nextButton.disabled = false;
}

function advanceQuestion(options = {}) {
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
  nextButton.classList.add("hidden");
  if (clearAnswerButton) {
    clearAnswerButton.classList.add("hidden");
  }

  if (!questionInfoArea) {
    advanceQuestion({ allowBlank: true });
    return;
  }

  const notice = document.createElement("div");
  notice.className = "question-timeout-notice";
  notice.innerHTML = `<p>${hasSelectedAnswer ? "O tempo desta questão acabou. Sua resposta foi registrada." : "O tempo desta questão acabou. Sua resposta não foi registrada."}</p><button type="button" class="btn btn-primary" id="question-timeout-next">Ir para a próxima questão</button>`;
  `< p > ${hasSelectedAnswer ? "O tempo desta questão acabou. Sua resposta foi registrada." : "O tempo desta questão acabou. Sua resposta não foi registrada."}</p > <button type="button" class="btn btn-primary" id="question-timeout-next">Ir para a próxima questão</button>`;
  questionInfoArea.innerHTML = "";
  questionInfoArea.appendChild(notice);

  const timeoutButton = document.getElementById("question-timeout-next");
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

  let timeLeft = question.timer;
  const timerBox = document.createElement("div");
  timerBox.className = "question-timer";
  timerBox.innerHTML = `<span>Tempo restante</span> <strong>${timeLeft}s</strong>`;
  questionInfoArea.appendChild(timerBox);

  window.questionTimerInterval = setInterval(() => {
    timeLeft -= 1;
    const strong = timerBox.querySelector("strong");
    if (strong) {
      strong.textContent = `${Math.max(0, timeLeft)} s`;
    }

    if (timeLeft <= 0) {
      clearInterval(window.questionTimerInterval);
      window.questionTimerInterval = null;
      handleQuestionTimerExpired(question);
    }
  }, 1000);
}

async function handleStartQuiz(event) {
  event.preventDefault();
  const quiz = getSelectedQuiz();
  if (!quiz) {
    return;
  }

  if (!state.isTeacher && !quiz.allowStudentStart) {
    alert("Este quiz está bloqueado para início no momento. Aguarde o professor liberar.");
    return;
  }

  if (!state.studentName) {
    alert("Complete seu perfil antes de iniciar o quiz.");
    return;
  }

  const enteredFullscreen = await requestFullscreenMode();
  if (!enteredFullscreen) {
    alert("�? obrigatório permanecer em tela cheia para realizar o questionário.");
    return;
  }

  clearSessionState();
  state.activeQuestions = buildShuffledQuestionSet(quiz);
  state.isActive = true;

  await consumeRetakeRelease(quiz.id);

  showOnlyScreen("quiz");
  cancelButton.classList.toggle("hidden", state.isTeacher);
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

  progress.textContent = `Questão ${state.currentQuestionIndex + 1} de ${activeQuestions.length} `;
  questionTitle.textContent = question.title;
  questionDescription.textContent = question.description || "";
  questionForm.innerHTML = "";
  renderQuestionTimer(question);

  questionForm.classList.add("masked");
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

  saveStoredAttempt(quiz.id, {
    status: "completed",
    result,
    answers: answersSnapshot,
    questionResults
  });

  if (shouldPersistLocalFallback) {
    saveAttemptToFirestore(quiz, result, answersSnapshot, questionResults);
  }

  const baseMessage = result.earnedPoints === result.maxPoints
    ? "Parabéns, você gabaritou todas as questões!"
    : "Resultado registrado com sucesso.";
  showResult(result, baseMessage);
  updateReviewButtons(quiz, { status: "completed", result, answers: answersSnapshot, questionResults });
}

function showResult(result, detailsText) {
  resultStudentName.textContent = result.studentName;
  if (resultQuizTitle) {
    resultQuizTitle.textContent = `Avaliação: ${result.quizTitle || result.quizId || "-"}`;
  }
  resultScore.textContent = `${result.earnedPoints} / ${result.maxPoints}`;
  resultPercent.textContent = `${result.percent}% de acertos`;
  const quiz = getSelectedQuiz();

  const prevDisabled = document.getElementById("result-review-disabled");
  if (prevDisabled && prevDisabled.parentNode) {
    prevDisabled.parentNode.removeChild(prevDisabled);
  }

  const existingNote = document.getElementById("result-blocked-note");
  if (existingNote) {
    existingNote.remove();
  }

  resultDetails.textContent = detailsText || "";

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
    cancelMessage.textContent = reason;
    showOnlyScreen("cancel");
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
    status: "cancelled",
    bloqueadoPorViolacao: Boolean(cancellation.blockedByViolation),
    motivoCancelamento: cancellation.reason || "",
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
    result: {
      quizId: quiz.id,
      quizTitle: quiz.title,
      studentName: state.studentName || "",
      earnedPoints: result.earnedPoints,
      maxPoints: result.maxPoints,
      percent: result.percent,
      date: result.date,
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
    status: "completed",
    bloqueadoPorViolacao: false,
    motivoCancelamento: "",
    respostas: answers,
    correcaoQuestoes: Array.isArray(questionResults) ? questionResults : []
  };

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

function handlePolicyViolation() {
  if (!state.isActive || state.isTeacher || state.isViolationGraceActive) {
    return;
  }

  if (state.violationCount >= 2) {
    cancelQuiz("Questionário cancelado após múltiplas violações das regras.", { blockedByViolation: true });
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
      cancelQuiz("Questionário cancelado porque você não retornou em até 10 segundos.", { blockedByViolation: true });
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
    cancelQuiz("Questionário cancelado porque você não retornou em até 10 segundos.", { blockedByViolation: true });
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

