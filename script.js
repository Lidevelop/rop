// Initialize jsPDF
const { jsPDF } = window.jspdf;

// Firebase configuration (substitua pelos dados do seu novo projeto)
const firebaseConfig = {
    apiKey: "AIzaSyDpMbzbUesMjUqtoF3UC0LpKU3t6HSkSrE",
    authDomain: "ropgma-bf685.firebaseapp.com",
    projectId: "ropgma-bf685",
    storageBucket: "ropgma-bf685.appspot.com",
    messagingSenderId: "458706183661",
    appId: "1:458706183661:web:0e92990b3f6e31f25133e4",
    measurementId: "G-52PRVFRNFP"
  };

let auth = null;
let db = null;
let storage = null;
let currentRopId = null;
let ropsCache = [];
let isLoadingRops = false;
let forcedRopIdFilter = null;
const ROPS_PAGE_SIZE = 6;
let ropsCurrentPage = 1;

let saveTimeout = null;
let currentUserProfile = null;
let isAdmin = false;
let usersCache = [];
let isLoadingUsers = false;
let warNamesCache = [];
let isLoadingWarNames = false;
let editingUserId = null;
let editingWarNameId = null;
let warNamesPageSize = 20;
let warNamesCurrentPage = 1;
let logsCache = [];
let isLoadingLogs = false;
let logsPageSize = 20;
let logsCurrentPage = 1;
let isGeneratingPdf = false;

function setPdfBusy(isBusy, button = null) {
    const existingOverlay = document.getElementById('pdfLoadingOverlay');
    if (isBusy) {
        if (!existingOverlay) {
            const overlay = document.createElement('div');
            overlay.id = 'pdfLoadingOverlay';
            overlay.innerHTML = `
                <div class="pdf-loading-content">
                    <div class="pdf-loading-spinner"></div>
                    <div class="pdf-loading-text">Gerando PDF...</div>
                </div>
            `;
            document.body.appendChild(overlay);
        }
        document.body.classList.add('pdf-busy');
    } else {
        document.body.classList.remove('pdf-busy');
        if (existingOverlay) existingOverlay.remove();
    }

    if (button) {
        if (isBusy) {
            button.dataset.originalText = button.innerHTML;
            button.disabled = true;
            button.classList.add('is-loading');
            button.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Gerando...';
        } else {
            button.disabled = false;
            button.classList.remove('is-loading');
            if (button.dataset.originalText) {
                button.innerHTML = button.dataset.originalText;
                delete button.dataset.originalText;
            }
        }
    }
}

let delegaciasCache = [];
let isLoadingDelegacias = false;

const SESSION_TIMEOUT_MS = 60 * 60 * 1000;
const SESSION_PROMPT_THRESHOLD_MS = 5 * 60 * 1000;
const SESSION_DEADLINE_STORAGE_KEY = 'pmuSessionDeadlineAt';
let idleTimeoutId = null;
let idleCountdownId = null;
let idleDeadlineAt = null;
let activityListenersBound = false;
let idlePromptShown = false;

let partes = [];
let apreensoes = [];
let agentes = [];
let anexosUrls = [];
let anexosPending = [];

const CLASSE_OPTIONS = [
    'S INSP 2ª CL',
    'S INSP 3ª CL',
    'SUP',
    'GUARDA AUXILIAR',
    'GM'
];

// DOM Elements
const loginView = document.getElementById('loginView');
const appView = document.getElementById('appView');
const loginForm = document.getElementById('loginForm');
const loginEmail = document.getElementById('loginEmail');
const loginPassword = document.getElementById('loginPassword');
const loginError = document.getElementById('loginError');
const logoutBtn = document.getElementById('logoutBtn');
const profileMenuContainer = document.getElementById('profileMenuContainer');
const profileMenuBtn = document.getElementById('profileMenuBtn');
const profileMenu = document.getElementById('profileMenu');
const changePasswordBtn = document.getElementById('changePasswordBtn');
const passwordModal = document.getElementById('passwordModal');
const closePasswordModal = document.getElementById('closePasswordModal');
const cancelPasswordBtn = document.getElementById('cancelPasswordBtn');
const passwordForm = document.getElementById('passwordForm');
const currentPasswordInput = document.getElementById('currentPassword');
const newPasswordInput = document.getElementById('newPassword');
const confirmNewPasswordInput = document.getElementById('confirmNewPassword');
const passwordError = document.getElementById('passwordError');

const addDelegaciaBtn = document.getElementById('addDelegaciaBtn');
const delegaciaModal = document.getElementById('delegaciaModal');
const closeDelegaciaModal = document.getElementById('closeDelegaciaModal');
const cancelDelegaciaBtn = document.getElementById('cancelDelegaciaBtn');
const saveDelegaciaBtn = document.getElementById('saveDelegaciaBtn');
const delegaciaNameInput = document.getElementById('delegaciaName');
const delegaciaError = document.getElementById('delegaciaError');
const delegaciasListUl = document.getElementById('delegaciasList');
const currentRopBtn = document.getElementById('currentRopBtn');
const newRopBtn = document.getElementById('newRopBtn');
const viewRopsBtn = document.getElementById('viewRopsBtn');
const viewUsersBtn = document.getElementById('viewUsersBtn');
const viewWarNamesBtn = document.getElementById('viewWarNamesBtn');
const viewLogsBtn = document.getElementById('viewLogsBtn');
const adminMenuContainer = document.getElementById('adminMenuContainer');
const adminMenuBtn = document.getElementById('adminMenuBtn');
const adminMenu = document.getElementById('adminMenu');
const ropFormView = document.getElementById('ropFormView');
const managerView = document.getElementById('managerView');
const usersView = document.getElementById('usersView');
const warNamesView = document.getElementById('warNamesView');
const logsView = document.getElementById('logsView');
const ropsList = document.getElementById('ropsList');
const ropsPagination = document.getElementById('ropsPagination');
const filterText = document.getElementById('filterText');
const filterDateFrom = document.getElementById('filterDateFrom');
const filterDateTo = document.getElementById('filterDateTo');
const refreshRops = document.getElementById('refreshRops');
const formInputs = document.querySelectorAll('#ropFormView input, #ropFormView textarea, #ropFormView select');
const saveIndicator = document.getElementById('saveIndicator');
const saveStatus = document.getElementById('saveStatus');
const lastSaved = document.getElementById('lastSaved');
const idleTimerDisplay = document.getElementById('idleTimer');
const registroNumeroInput = document.getElementById('registroNumero');
const registroNumeroDisplay = document.getElementById('registroNumeroDisplay');
const idlePromptModal = document.getElementById('idlePromptModal');
const idleModalTimerDisplay = document.getElementById('idleModalTimer');
const idleExtendBtn = document.getElementById('idleExtendBtn');
const idleKeepBtn = document.getElementById('idleKeepBtn');
const addParteBtn = document.getElementById('addParteBtn');
const addApreensaoBtn = document.getElementById('addApreensaoBtn');
const addAgenteBtn = document.getElementById('addAgenteBtn');
const partesList = document.getElementById('partesList');
const apreensoesContainer = document.getElementById('apreensoesContainer');
const apreensoesEmpty = document.getElementById('apreensoesEmpty');
const agentesContainer = document.getElementById('agentesContainer');
const agentesEmpty = document.getElementById('agentesEmpty');
const anexosContainer = document.getElementById('anexosContainer');
const anexosEmpty = document.getElementById('anexosEmpty');
const addAnexosBtn = document.getElementById('addAnexosBtn');
const anexosFileInput = document.getElementById('anexosFileInput');
const anexosUploadsList = document.getElementById('anexosUploadsList');
let isUploadingAnexo = false;
const apreensoesTableBody = document.getElementById('apreensoesTableBody');
const apreensoesEmptyMessage = document.getElementById('apreensoesEmpty');
const agentesTableBody = document.getElementById('agentesTableBody');
const partesEmpty = document.getElementById('partesEmpty');
const noPartesFlagWrapper = document.getElementById('noPartesFlagWrapper');
const noPartesFlag = document.getElementById('noPartesFlag');
const noApreensoesFlagWrapper = document.getElementById('noApreensoesFlagWrapper');
const noApreensoesFlag = document.getElementById('noApreensoesFlag');
const noAnexosFlagWrapper = document.getElementById('noAnexosFlagWrapper');
const noAnexosFlag = document.getElementById('noAnexosFlag');
const saveDbBtn = document.getElementById('saveDbBtn');
const clearBtn = document.getElementById('clearBtn');
const confirmationModal = document.getElementById('confirmationModal');
const confirmClear = document.getElementById('confirmClear');
const cancelClear = document.getElementById('cancelClear');
const newRopWarningModal = document.getElementById('newRopWarningModal');
const confirmNewRopBtn = document.getElementById('confirmNewRop');
const cancelNewRopBtn = document.getElementById('cancelNewRop');
const parteModal = document.getElementById('parteModal');
const parteForm = document.getElementById('parteForm');
const parteRole = document.getElementById('parteRole');
const parteNome = document.getElementById('parteNome');
const parteNascimento = document.getElementById('parteNascimento');
const partePai = document.getElementById('partePai');
const parteMae = document.getElementById('parteMae');
const parteCondicao = document.getElementById('parteCondicao');
const parteSexo = document.getElementById('parteSexo');
const parteNaturalidade = document.getElementById('parteNaturalidade');
const parteEndereco = document.getElementById('parteEndereco');
const parteCpf = document.getElementById('parteCpf');
const parteCidadeUf = document.getElementById('parteCidadeUf');
const parteTelefone = document.getElementById('parteTelefone');
const saveParteBtn = document.getElementById('saveParteBtn');
const cancelParteBtn = document.getElementById('cancelParteBtn');
const closeParteBtn = document.getElementById('closeParteBtn');
let editingParteId = null;
const profileSummaryClass = document.getElementById('profileSummaryClass');
const profileSummaryWarName = document.getElementById('profileSummaryWarName');
const profileSummaryRegistration = document.getElementById('profileSummaryRegistration');
const profileSummaryType = document.getElementById('profileSummaryType');
const usersTableBody = document.getElementById('usersTableBody');
const createUserBtn = document.getElementById('createUserBtn');
const userCreateModal = document.getElementById('userCreateModal');
const userCreateForm = document.getElementById('userCreateForm');
const userCreateEmail = document.getElementById('userCreateEmail');
const userCreatePassword = document.getElementById('userCreatePassword');
const userCreateClass = document.getElementById('userCreateClass');
const userCreateWarName = document.getElementById('userCreateWarName');
const userCreateRegistration = document.getElementById('userCreateRegistration');
const userCreateType = document.getElementById('userCreateType');
const saveUserCreateBtn = document.getElementById('saveUserCreateBtn');
const cancelUserCreateBtn = document.getElementById('cancelUserCreateBtn');
const userCreateError = document.getElementById('userCreateError');
const userEditModal = document.getElementById('userEditModal');
const userEditForm = document.getElementById('userEditForm');
const userEditClass = document.getElementById('userEditClass');
const userEditWarName = document.getElementById('userEditWarName');
const userEditRegistration = document.getElementById('userEditRegistration');
const userEditType = document.getElementById('userEditType');
const saveUserEditBtn = document.getElementById('saveUserEditBtn');
const cancelUserEditBtn = document.getElementById('cancelUserEditBtn');
const closeUserEditBtn = document.getElementById('closeUserEditBtn');
const userEditError = document.getElementById('userEditError');
const warNamesTableBody = document.getElementById('warNamesTableBody');
const warNamesFilterText = document.getElementById('warNamesFilterText');
const warNamesPageSizeSelect = document.getElementById('warNamesPageSize');
const warNamesPagination = document.getElementById('warNamesPagination');
const logsTableBody = document.getElementById('logsTableBody');
const logsFilterText = document.getElementById('logsFilterText');
const logsPageSizeSelect = document.getElementById('logsPageSize');
const logsPagination = document.getElementById('logsPagination');
const createWarNameBtn = document.getElementById('createWarNameBtn');
const warNameModal = document.getElementById('warNameModal');
const warNameForm = document.getElementById('warNameForm');
const warNameClass = document.getElementById('warNameClass');
const warNameValue = document.getElementById('warNameValue');
const warNameRegistration = document.getElementById('warNameRegistration');
const warNameStatus = document.getElementById('warNameStatus');
const saveWarNameBtn = document.getElementById('saveWarNameBtn');
const cancelWarNameBtn = document.getElementById('cancelWarNameBtn');
const warNameError = document.getElementById('warNameError');
const warNamesList = document.getElementById('warNamesList');
const importWarNamesBtn = document.getElementById('importWarNamesBtn');
const resetPasswordModal = document.getElementById('resetPasswordModal');
const resetPasswordName = document.getElementById('resetPasswordName');
const resetPasswordEmail = document.getElementById('resetPasswordEmail');
const resetPasswordError = document.getElementById('resetPasswordError');
const confirmResetPasswordBtn = document.getElementById('confirmResetPasswordBtn');
const cancelResetPasswordBtn = document.getElementById('cancelResetPasswordBtn');
const closeResetPasswordBtn = document.getElementById('closeResetPasswordBtn');
let pendingResetUser = null;

function setSaveButtonLoading(isLoading) {
    if (!saveDbBtn) return;
    if (isLoading) {
        if (!saveDbBtn.dataset.originalHtml) {
            saveDbBtn.dataset.originalHtml = saveDbBtn.innerHTML;
        }
        saveDbBtn.disabled = true;
        saveDbBtn.classList.add('btn-loading');
        saveDbBtn.setAttribute('aria-busy', 'true');
        saveDbBtn.innerHTML = '<span class="btn-spinner" aria-hidden="true"></span> Salvando...';
    } else {
        saveDbBtn.disabled = false;
        saveDbBtn.classList.remove('btn-loading');
        saveDbBtn.removeAttribute('aria-busy');
        if (saveDbBtn.dataset.originalHtml) {
            saveDbBtn.innerHTML = saveDbBtn.dataset.originalHtml;
            delete saveDbBtn.dataset.originalHtml;
        }
    }
}

function buildClasseOptions(selectedValue = '') {
    const options = ['<option value="">Selecione</option>'];
    CLASSE_OPTIONS.forEach(option => {
        const isSelected = option === selectedValue ? ' selected' : '';
        options.push(`<option value="${option}"${isSelected}>${option}</option>`);
    });
    return options.join('');
}

function buildWarNameOptions(selectedValue = '') {
    const activeNames = warNamesCache.filter(item => item.status !== 'inativo');
    if (!activeNames.length) {
        return '<option value="">Nenhum agente cadastrado</option>';
    }
    activeNames.sort((a, b) => String(a.nomeGuerra || '').localeCompare(String(b.nomeGuerra || ''), 'pt-BR'));
    const options = ['<option value="">Selecione</option>'];
    activeNames.forEach(item => {
        const isSelected = item.nomeGuerra === selectedValue ? ' selected' : '';
        options.push(`<option value="${item.nomeGuerra}"${isSelected}>${item.nomeGuerra}</option>`);
    });
    return options.join('');
}

function buildUserCreateWarNameOptions(selectedValue = '') {
    const activeNames = warNamesCache.filter(item => item.status !== 'inativo');
    if (!activeNames.length) {
        return '<option value="">Nenhum agente cadastrado</option>';
    }
    const usedMatriculas = new Set(
        usersCache
            .map(user => (user?.matricula || '').trim())
            .filter(Boolean)
    );
    const usedNames = new Set(
        usersCache
            .map(user => (user?.nomeGuerra || '').trim().toLowerCase())
            .filter(Boolean)
    );
    const available = activeNames.filter(item => {
        const matricula = (item?.matricula || '').trim();
        const nomeGuerra = (item?.nomeGuerra || '').trim().toLowerCase();
        if (matricula && usedMatriculas.has(matricula)) return false;
        if (nomeGuerra && usedNames.has(nomeGuerra)) return false;
        return true;
    });
    if (!available.length) {
        return '<option value="">Nenhum agente disponível</option>';
    }
    available.sort((a, b) => String(a.nomeGuerra || '').localeCompare(String(b.nomeGuerra || ''), 'pt-BR'));
    const options = ['<option value="">Selecione</option>'];
    available.forEach(item => {
        const isSelected = item.nomeGuerra === selectedValue ? ' selected' : '';
        options.push(`<option value="${item.nomeGuerra}"${isSelected}>${item.nomeGuerra}</option>`);
    });
    return options.join('');
}

function updateUserCreateWarNameSelect(selectedValue = '') {
    if (!userCreateWarName) return;
    userCreateWarName.innerHTML = buildUserCreateWarNameOptions(selectedValue);
}

function bootApp() {
    initFirebase();
    initializeEventListeners();
    initializeAuthListeners();

    const today = new Date().toISOString().split('T')[0];
    const now = new Date();
    const hours = now.getHours().toString().padStart(2, '0');
    const minutes = now.getMinutes().toString().padStart(2, '0');

    const dataFatoInput = document.getElementById('dataFato');
    const horaFatoInput = document.getElementById('horaFato');
    if (dataFatoInput) dataFatoInput.value = today;
    if (horaFatoInput) horaFatoInput.value = `${hours}:${minutes}`;

    loadFormData();
    renderPartes();
    renderApreensoes();
    renderAgentes();
    updateEmptyFlags();

    updateLastSavedTime();
    expandAllTextareas();
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bootApp);
} else {
    bootApp();
}

function initFirebase() {
    if (!window.firebase) {
        console.error('Firebase SDK não carregado.');
        return;
    }

    if (!firebase.apps.length) {
        firebase.initializeApp(firebaseConfig);
    }

    auth = firebase.auth();
    db = firebase.firestore();
    storage = firebase.storage ? firebase.storage() : null;

    auth.onAuthStateChanged(async (user) => {
        if (user) {
            // Enforce session timeout even after closing the browser.
            const storedDeadline = Number(localStorage.getItem(SESSION_DEADLINE_STORAGE_KEY) || 0);
            if (storedDeadline && storedDeadline <= Date.now()) {
                localStorage.removeItem(SESSION_DEADLINE_STORAGE_KEY);
                try {
                    await auth.signOut();
                } catch (e) {
                    console.error(e);
                }
                return;
            }

            showAppView();
            await initializeUserProfileFlow(user);
            await loadWarNamesList(true);
            await loadDelegaciasList(true);
            await loadRopsList(true);
            await logEvent('login');
            startIdleLogoutTimer();
            await loadNextRopNumberPreview();
        } else {
            showLoginView();
            resetUserProfileState();
            stopIdleLogoutTimer();
        }
    });
}

function bindIdleActivityListeners() {
    if (activityListenersBound) return;
    const activityEvents = ['mousedown'];
    activityEvents.forEach(eventName => {
        document.addEventListener(eventName, handleUserActivity, { passive: true });
    });
    activityListenersBound = true;
}

function handleUserActivity() {
    return;
}

function startIdleLogoutTimer() {
    if (!auth || !auth.currentUser) return;
    const storedDeadline = Number(localStorage.getItem(SESSION_DEADLINE_STORAGE_KEY) || 0);
    if (storedDeadline && storedDeadline > Date.now()) {
        idleDeadlineAt = storedDeadline;
    } else {
        idleDeadlineAt = Date.now() + SESSION_TIMEOUT_MS;
        localStorage.setItem(SESSION_DEADLINE_STORAGE_KEY, String(idleDeadlineAt));
    }
    idlePromptShown = false;
    if (idleTimeoutId) clearTimeout(idleTimeoutId);
    if (idleCountdownId) clearInterval(idleCountdownId);
    const remaining = Math.max(0, idleDeadlineAt - Date.now());
    idleTimeoutId = setTimeout(forceLogoutDueToInactivity, remaining);
    idleCountdownId = setInterval(updateIdleCountdown, 1000);
    updateIdleCountdown();
}

function stopIdleLogoutTimer() {
    if (idleTimeoutId) {
        clearTimeout(idleTimeoutId);
        idleTimeoutId = null;
    }
    if (idleCountdownId) {
        clearInterval(idleCountdownId);
        idleCountdownId = null;
    }
    idleDeadlineAt = null;
    idlePromptShown = false;
    if (idleTimerDisplay) {
        idleTimerDisplay.textContent = '';
    }
    if (idleModalTimerDisplay) {
        idleModalTimerDisplay.textContent = '';
    }
    hideIdlePrompt();
    localStorage.removeItem(SESSION_DEADLINE_STORAGE_KEY);
}

function forceLogoutDueToInactivity() {
    if (auth && auth.currentUser) {
        localStorage.removeItem(SESSION_DEADLINE_STORAGE_KEY);
        auth.signOut();
    }
}

function updateIdleCountdown() {
    if (!idleTimerDisplay || !idleDeadlineAt) return;
    const remaining = Math.max(0, idleDeadlineAt - Date.now());
    const minutes = Math.floor(remaining / 60000);
    const seconds = Math.floor((remaining % 60000) / 1000);
    const clock = `${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
    idleTimerDisplay.textContent = `Sessão expira em ${clock}`;
    if (idleModalTimerDisplay) idleModalTimerDisplay.textContent = clock;

    if (!idlePromptShown && remaining <= SESSION_PROMPT_THRESHOLD_MS && remaining > 0) {
        showIdlePrompt();
    }

    if (remaining <= 0) {
        forceLogoutDueToInactivity();
    }
}

function showIdlePrompt() {
    if (!idlePromptModal) return;
    idlePromptShown = true;
    idlePromptModal.style.display = 'flex';
}

function hideIdlePrompt() {
    if (!idlePromptModal) return;
    idlePromptModal.style.display = 'none';
}

if (idleExtendBtn) {
    idleExtendBtn.addEventListener('click', () => {
        idleDeadlineAt = Date.now() + SESSION_TIMEOUT_MS;
        localStorage.setItem(SESSION_DEADLINE_STORAGE_KEY, String(idleDeadlineAt));
        idlePromptShown = false;
        hideIdlePrompt();
        startIdleLogoutTimer();
        updateIdleCountdown();
    });
}

if (idleKeepBtn) {
    idleKeepBtn.addEventListener('click', () => {
        hideIdlePrompt();
    });
}

function initializeAuthListeners() {
    if (loginForm) {
        loginForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            loginError.textContent = '';
            try {
                await auth.signInWithEmailAndPassword(loginEmail.value.trim(), loginPassword.value.trim());
            } catch (error) {
                loginError.textContent = 'Falha ao entrar. Verifique e-mail e senha.';
                console.error(error);
            }
        });
    }

    if (logoutBtn) {
        logoutBtn.addEventListener('click', async () => {
            await auth.signOut();
        });
    }
}

function initializeEventListeners() {
        // Delegacias: renderizar lista e excluir
        if (delegaciasListUl) {
            delegaciasListUl.addEventListener('click', function (e) {
                if (e.target.classList.contains('delete-delegacia-btn')) {
                    const id = e.target.getAttribute('data-id');
                    if (id && confirm('Deseja realmente excluir esta delegacia?')) {
                        removeDelegaciaById(id);
                    }
                }
            });
        }
    formInputs.forEach(input => {
        input.addEventListener('input', () => {
            if (input.classList.contains('field-error')) {
                if (input.type === 'checkbox') {
                    if (input.checked) input.classList.remove('field-error');
                } else if (input.value.trim() !== '') {
                    input.classList.remove('field-error');
                }
            }
            triggerSave();
            updateEmptyFlags();
        });
    });

    document.querySelectorAll('#ropFormView textarea').forEach(textarea => {
        autoExpandTextarea(textarea);
        textarea.addEventListener('input', function () {
            autoExpandTextarea(this);
            triggerSave();
        });
    });

    if (addParteBtn) addParteBtn.addEventListener('click', () => openParteModal());
    if (addApreensaoBtn) addApreensaoBtn.addEventListener('click', addApreensao);
    if (addAgenteBtn) addAgenteBtn.addEventListener('click', addAgente);
    if (addAnexosBtn) addAnexosBtn.addEventListener('click', handleAddAnexoClick);

    if (anexosFileInput) {
        anexosFileInput.addEventListener('change', async () => {
            const file = anexosFileInput.files && anexosFileInput.files[0] ? anexosFileInput.files[0] : null;
            anexosFileInput.value = '';
            if (!file) return;
            await uploadAnexoImage(file);
        });
    }

    bindParteInputMasks();

    if (saveDbBtn) {
        saveDbBtn.addEventListener('click', async () => {
            if (!auth || !auth.currentUser) return;
            if (!validateRequiredFields()) {
                saveStatus.textContent = 'Preencha os campos obrigatórios antes de salvar.';
                return;
            }
            setSaveButtonLoading(true);
            try {
                saveStatus.textContent = 'Salvando no Firebase...';
                const formData = getFormData();
                const savedId = await saveRopToFirebase(formData);
                if (savedId) {
                    await uploadPendingAnexosToStorage(savedId);
                    forcedRopIdFilter = savedId;
                    clearForm();
                    setActiveAppView('manager');
                    loadRopsList(true);
                }
            } finally {
                setSaveButtonLoading(false);
            }
        });
    }

    if (clearBtn) {
        clearBtn.addEventListener('click', () => {
            confirmationModal.style.display = 'flex';
        });
    }

    if (confirmClear) confirmClear.addEventListener('click', clearForm);
    if (cancelClear) cancelClear.addEventListener('click', () => {
        confirmationModal.style.display = 'none';
    });

    if (confirmNewRopBtn) {
        confirmNewRopBtn.addEventListener('click', () => {
            if (newRopWarningModal) newRopWarningModal.style.display = 'none';
            confirmNewRop();
            setActiveAppView('form');
        });
    }

    if (cancelNewRopBtn) {
        cancelNewRopBtn.addEventListener('click', () => {
            if (newRopWarningModal) newRopWarningModal.style.display = 'none';
        });
    }


    if (currentRopBtn) currentRopBtn.addEventListener('click', () => setActiveAppView('form'));
    if (newRopBtn) newRopBtn.addEventListener('click', () => {
        if (newRopWarningModal) {
            newRopWarningModal.style.display = 'flex';
        } else {
            confirmNewRop();
            setActiveAppView('form');
        }
    });
    if (viewRopsBtn) viewRopsBtn.addEventListener('click', () => setActiveAppView('manager'));
    if (viewUsersBtn) viewUsersBtn.addEventListener('click', () => {
        if (isAdmin) {
            setActiveAppView('users');
            loadUsersList(true);
        }
    });
    if (viewWarNamesBtn) viewWarNamesBtn.addEventListener('click', () => {
        if (isAdmin) {
            setActiveAppView('warNames');
            loadWarNamesList(true);
        }
    });
    if (viewLogsBtn) viewLogsBtn.addEventListener('click', () => {
        if (isAdmin) {
            setActiveAppView('logs');
            loadLogsList(true);
        }
    });

    [filterText, filterDateFrom, filterDateTo].forEach(input => {
        if (!input) return;
        input.addEventListener('input', () => {
            forcedRopIdFilter = null;
            ropsCurrentPage = 1;
            applyRopsFilter();
        });
        input.addEventListener('change', () => {
            forcedRopIdFilter = null;
            ropsCurrentPage = 1;
            applyRopsFilter();
        });
    });

    if (refreshRops) {
        refreshRops.addEventListener('click', () => {
            filterText.value = '';
            filterDateFrom.value = '';
            filterDateTo.value = '';
            forcedRopIdFilter = null;
            ropsCurrentPage = 1;
            loadRopsList(true);
        });
    }

    if (adminMenuBtn && adminMenu) {
        adminMenuBtn.addEventListener('click', (event) => {
            event.stopPropagation();
            const isHidden = adminMenu.classList.contains('hidden');
            adminMenu.classList.toggle('hidden', !isHidden);
            adminMenuBtn.setAttribute('aria-expanded', String(isHidden));
        });
    }

    if (profileMenuBtn && profileMenu) {
        profileMenuBtn.addEventListener('click', (event) => {
            event.stopPropagation();
            const isHidden = profileMenu.classList.contains('hidden');
            profileMenu.classList.toggle('hidden', !isHidden);
            profileMenuBtn.setAttribute('aria-expanded', String(isHidden));
        });
    }

    if (changePasswordBtn) {
        changePasswordBtn.addEventListener('click', (event) => {
            event.preventDefault();
            openPasswordModal();
        });
    }

    if (closePasswordModal) {
        closePasswordModal.addEventListener('click', () => closePasswordModalFn());
    }
    if (cancelPasswordBtn) {
        cancelPasswordBtn.addEventListener('click', () => closePasswordModalFn());
    }
    if (passwordForm) {
        passwordForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            await handlePasswordChange();
        });
    }

    if (addDelegaciaBtn) {
        addDelegaciaBtn.addEventListener('click', () => openDelegaciaModal());
    }
    if (closeDelegaciaModal) {
        closeDelegaciaModal.addEventListener('click', () => closeDelegaciaModalFn());
    }
    if (cancelDelegaciaBtn) {
        cancelDelegaciaBtn.addEventListener('click', () => closeDelegaciaModalFn());
    }
    if (saveDelegaciaBtn) {
        saveDelegaciaBtn.addEventListener('click', async () => {
            await saveDelegaciaFromModal();
        });
    }

    [noPartesFlag, noApreensoesFlag, noAnexosFlag].forEach(flag => {
        if (!flag) return;
        flag.addEventListener('change', () => {
            const wrapper = flag.closest('.empty-flag');
            if (wrapper) wrapper.classList.remove('field-error');
            triggerSave();
        });
    });

    document.addEventListener('click', () => {
        if (adminMenu && adminMenuBtn) {
            adminMenu.classList.add('hidden');
            adminMenuBtn.setAttribute('aria-expanded', 'false');
        }
        if (profileMenu && profileMenuBtn) {
            profileMenu.classList.add('hidden');
            profileMenuBtn.setAttribute('aria-expanded', 'false');
        }
    });


    if (createUserBtn) {
        createUserBtn.addEventListener('click', () => openUserCreateModal());
    }
    if (saveUserCreateBtn) {
        saveUserCreateBtn.addEventListener('click', (event) => {
            event.preventDefault();
            createNewUser();
        });
    }
    if (cancelUserCreateBtn) {
        cancelUserCreateBtn.addEventListener('click', (event) => {
            event.preventDefault();
            closeUserCreateModal();
        });
    }

    if (saveUserEditBtn) {
        saveUserEditBtn.addEventListener('click', (event) => {
            event.preventDefault();
            saveEditedUserProfile();
        });
    }
    if (cancelUserEditBtn) {
        cancelUserEditBtn.addEventListener('click', (event) => {
            event.preventDefault();
            closeUserEditModal();
        });
    }
    if (closeUserEditBtn) {
        closeUserEditBtn.addEventListener('click', (event) => {
            event.preventDefault();
            closeUserEditModal();
        });
    }

    if (createWarNameBtn) {
        createWarNameBtn.addEventListener('click', () => openWarNameModal());
    }
    if (importWarNamesBtn) {
        importWarNamesBtn.addEventListener('click', () => importWarNamesFromSeed());
    }
    if (warNamesFilterText) {
        warNamesFilterText.addEventListener('input', () => {
            warNamesCurrentPage = 1;
            applyWarNamesFilter();
        });
    }
    if (warNamesPageSizeSelect) {
        warNamesPageSizeSelect.addEventListener('change', () => {
            warNamesPageSize = Number(warNamesPageSizeSelect.value) || 20;
            warNamesCurrentPage = 1;
            applyWarNamesFilter();
        });
    }
    if (logsFilterText) {
        logsFilterText.addEventListener('input', () => {
            logsCurrentPage = 1;
            applyLogsFilter();
        });
    }
    if (logsPageSizeSelect) {
        logsPageSizeSelect.addEventListener('change', () => {
            logsPageSize = Number(logsPageSizeSelect.value) || 20;
            logsCurrentPage = 1;
            applyLogsFilter();
        });
    }
    if (saveWarNameBtn) {
        saveWarNameBtn.addEventListener('click', (event) => {
            event.preventDefault();
            saveWarNameEntry();
        });
    }
    if (cancelWarNameBtn) {
        cancelWarNameBtn.addEventListener('click', (event) => {
            event.preventDefault();
            closeWarNameModal();
        });
    }

    if (saveParteBtn) {
        saveParteBtn.addEventListener('click', (event) => {
            event.preventDefault();
            saveParteFromModal();
        });
    }

    if (cancelParteBtn) {
        cancelParteBtn.addEventListener('click', (event) => {
            event.preventDefault();
            closeParteModal();
        });
    }

    if (closeParteBtn) {
        closeParteBtn.addEventListener('click', (event) => {
            event.preventDefault();
            confirmCloseParteModal();
        });
    }


    if (apreensoesTableBody) {
        apreensoesTableBody.addEventListener('input', handleApreensaoInputChange);
        apreensoesTableBody.addEventListener('change', handleApreensaoInputChange);
    }

    if (agentesTableBody) {
        agentesTableBody.addEventListener('input', handleAgenteInputChange);
        agentesTableBody.addEventListener('change', handleAgenteInputChange);
    }

    if (userCreateWarName) {
        userCreateWarName.addEventListener('change', () => autofillWarNameFields(userCreateWarName, userCreateClass, userCreateRegistration));
    }
    if (userEditWarName) {
        userEditWarName.addEventListener('change', () => autofillWarNameFields(userEditWarName, userEditClass, userEditRegistration));
    }
    if (confirmResetPasswordBtn) {
        confirmResetPasswordBtn.addEventListener('click', (event) => {
            event.preventDefault();
            sendPasswordResetForUser();
        });
    }
    if (cancelResetPasswordBtn) {
        cancelResetPasswordBtn.addEventListener('click', (event) => {
            event.preventDefault();
            closeResetPasswordModal();
        });
    }
    if (closeResetPasswordBtn) {
        closeResetPasswordBtn.addEventListener('click', (event) => {
            event.preventDefault();
            closeResetPasswordModal();
        });
    }
}

function bindParteInputMasks() {
    if (parteCpf) {
        parteCpf.addEventListener('input', () => {
            parteCpf.value = maskCPF(parteCpf.value);
        });
        parteCpf.addEventListener('blur', () => {
            // Normalize mask on blur
            parteCpf.value = maskCPF(parteCpf.value);
        });
    }
    if (parteTelefone) {
        parteTelefone.addEventListener('input', () => {
            parteTelefone.value = maskTelefone(parteTelefone.value);
        });
        parteTelefone.addEventListener('blur', () => {
            parteTelefone.value = maskTelefone(parteTelefone.value);
        });
    }
}

function openPasswordModal() {
    if (!passwordModal) return;
    if (passwordError) passwordError.textContent = '';
    if (currentPasswordInput) currentPasswordInput.value = '';
    if (newPasswordInput) newPasswordInput.value = '';
    if (confirmNewPasswordInput) confirmNewPasswordInput.value = '';
    passwordModal.style.display = 'flex';
    if (profileMenu) profileMenu.classList.add('hidden');
    if (profileMenuBtn) profileMenuBtn.setAttribute('aria-expanded', 'false');
    setTimeout(() => currentPasswordInput?.focus(), 0);
}

function closePasswordModalFn() {
    if (!passwordModal) return;
    passwordModal.style.display = 'none';
}

async function handlePasswordChange() {
    if (!auth || !auth.currentUser) return;
    const currentPassword = currentPasswordInput ? currentPasswordInput.value : '';
    const newPassword = newPasswordInput ? newPasswordInput.value : '';
    const confirmPassword = confirmNewPasswordInput ? confirmNewPasswordInput.value : '';

    if (passwordError) passwordError.textContent = '';
    if (!currentPassword || !newPassword || !confirmPassword) {
        if (passwordError) passwordError.textContent = 'Preencha todos os campos.';
        return;
    }
    if (newPassword.length < 6) {
        if (passwordError) passwordError.textContent = 'A nova senha deve ter no mínimo 6 caracteres.';
        return;
    }
    if (newPassword !== confirmPassword) {
        if (passwordError) passwordError.textContent = 'A confirmação da nova senha não confere.';
        return;
    }

    try {
        const user = auth.currentUser;
        const email = user.email || '';
        const credential = firebase.auth.EmailAuthProvider.credential(email, currentPassword);
        await user.reauthenticateWithCredential(credential);
        await user.updatePassword(newPassword);
        closePasswordModalFn();
        saveStatus.textContent = 'Senha alterada com sucesso.';
    } catch (error) {
        console.error(error);
        if (!passwordError) return;
        if (error?.code === 'auth/wrong-password') {
            passwordError.textContent = 'Senha atual incorreta.';
        } else if (error?.code === 'auth/too-many-requests') {
            passwordError.textContent = 'Muitas tentativas. Tente novamente mais tarde.';
        } else {
            passwordError.textContent = 'Não foi possível alterar a senha.';
        }
    }
}

function openDelegaciaModal() {
    if (!delegaciaModal) return;
    if (delegaciaError) delegaciaError.textContent = '';
    if (delegaciaNameInput) delegaciaNameInput.value = '';
    delegaciaModal.style.display = 'flex';
    renderDelegaciasList();
    setTimeout(() => delegaciaNameInput?.focus(), 0);
}

function closeDelegaciaModalFn() {
    if (!delegaciaModal) return;
    delegaciaModal.style.display = 'none';
}

async function saveDelegaciaFromModal() {
    if (!db) return;
    const name = delegaciaNameInput ? delegaciaNameInput.value.trim() : '';
    if (delegaciaError) delegaciaError.textContent = '';
    if (!name) {
        if (delegaciaError) delegaciaError.textContent = 'Informe o nome da delegacia.';
        delegaciaNameInput?.focus();
        return;
    }

    try {
        await db.collection('delegacias').add({
            nome: name,
            criadoEm: firebase.firestore.FieldValue.serverTimestamp(),
            criadoPor: auth?.currentUser?.uid || null
        });
        closeDelegaciaModalFn();
        await loadDelegaciasList(true);
        const destinatario = document.getElementById('destinatario');
        if (destinatario) destinatario.value = name;
        triggerSave();
    } catch (error) {
        console.error(error);
        if (delegaciaError) delegaciaError.textContent = 'Erro ao salvar delegacia.';
    }
}

async function loadDelegaciasList(force = false) {
    const destinatarioSelect = document.getElementById('destinatario');
    if (!db || !auth || !auth.currentUser || !destinatarioSelect) return;
    if (isLoadingDelegacias) return;

    if (delegaciasCache.length > 0 && !force) {
        hydrateDelegaciasSelect(destinatarioSelect);
        return;
    }

    isLoadingDelegacias = true;
    try {
        const snapshot = await db.collection('delegacias').orderBy('nome').get();
        delegaciasCache = snapshot.docs
            .map(doc => ({ id: doc.id, ...doc.data() }))
            .filter(item => item && item.nome);
        hydrateDelegaciasSelect(destinatarioSelect);
        renderDelegaciasList();
    } catch (error) {
        console.error('Erro ao carregar delegacias:', error);
    } finally {
        isLoadingDelegacias = false;
    }
}

// Renderiza a lista de delegacias no modal
function renderDelegaciasList() {
    if (!delegaciasListUl) return;
    if (!Array.isArray(delegaciasCache) || delegaciasCache.length === 0) {
        delegaciasListUl.innerHTML = '<li style="color:#888;font-size:0.95rem;">Nenhuma delegacia cadastrada.</li>';
        return;
    }
    delegaciasListUl.innerHTML = delegaciasCache.map(item =>
        `<li>${item.nome}<button class="delete-delegacia-btn" data-id="${item.id}" title="Excluir delegacia"><i class="fas fa-trash"></i></button></li>`
    ).join('');
}

// Remove delegacia pelo id e atualiza lista
async function removeDelegaciaById(id) {
    if (!db || !id) return;
    try {
        await db.collection('delegacias').doc(id).delete();
        await loadDelegaciasList(true);
    } catch (error) {
        alert('Erro ao excluir delegacia.');
    }
}

function hydrateDelegaciasSelect(selectEl) {
    const currentValue = selectEl.value;
    selectEl.innerHTML = '';

    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = 'Selecione';
    selectEl.appendChild(placeholder);

    delegaciasCache.forEach(item => {
        const option = document.createElement('option');
        option.value = String(item.nome);
        option.textContent = String(item.nome);
        selectEl.appendChild(option);
    });

    if (currentValue) {
        const exists = Array.from(selectEl.options).some(opt => opt.value === currentValue);
        if (!exists) {
            const extra = document.createElement('option');
            extra.value = currentValue;
            extra.textContent = currentValue;
            selectEl.appendChild(extra);
        }
        selectEl.value = currentValue;
    }
}

function updateEmptyFlags() {
    const hasPartes = Array.isArray(partes) && partes.length > 0;
    if (partesEmpty) partesEmpty.classList.toggle('hidden', hasPartes);
    if (noPartesFlagWrapper) noPartesFlagWrapper.classList.toggle('hidden', hasPartes);
    if (hasPartes && noPartesFlag) noPartesFlag.checked = false;

    const hasApreensoes = Array.isArray(apreensoes) && apreensoes.length > 0;
    if (noApreensoesFlagWrapper) noApreensoesFlagWrapper.classList.toggle('hidden', hasApreensoes);
    if (hasApreensoes && noApreensoesFlag) noApreensoesFlag.checked = false;

    const hasAnexos = Array.isArray(anexosUrls) && anexosUrls.length > 0;
    const hasPending = Array.isArray(anexosPending) && anexosPending.length > 0;
    if (noAnexosFlagWrapper) noAnexosFlagWrapper.classList.toggle('hidden', hasAnexos || hasPending);
    if ((hasAnexos || hasPending) && noAnexosFlag) noAnexosFlag.checked = false;
}

function showLoginView() {
    if (loginView) loginView.classList.remove('hidden');
    if (appView) appView.classList.add('hidden');
}

function showAppView() {
    if (loginView) loginView.classList.add('hidden');
    if (appView) appView.classList.remove('hidden');
}

function resetUserProfileState() {
    currentUserProfile = null;
    isAdmin = false;
    usersCache = [];
    editingUserId = null;
    warNamesCache = [];
    editingWarNameId = null;
    logsCache = [];
    delegaciasCache = [];
    updateProfileSummary(null);

    if (adminMenuContainer) adminMenuContainer.classList.add('hidden');
    if (usersView) usersView.classList.add('hidden');
    if (warNamesView) warNamesView.classList.add('hidden');
    if (logsView) logsView.classList.add('hidden');
    if (userCreateModal) userCreateModal.style.display = 'none';
    if (userEditModal) userEditModal.style.display = 'none';
    if (warNameModal) warNameModal.style.display = 'none';
}

async function initializeUserProfileFlow(user) {
    if (!db || !user) return;
    await ensureUserProfile(user);
    applyAdminVisibility();
}

async function ensureUserProfile(user) {
    const userRef = db.collection('users').doc(user.uid);
    const docSnap = await userRef.get();
    if (!docSnap.exists) {
        await userRef.set({
            email: user.email || '',
            classe: '',
            nomeGuerra: '',
            matricula: '',
            tipo: 'usuario',
            criadoEm: firebase.firestore.FieldValue.serverTimestamp()
        });
    }
    const updatedSnap = await userRef.get();
    currentUserProfile = updatedSnap.data();
    isAdmin = currentUserProfile?.tipo === 'administrador';
    updateProfileSummary(currentUserProfile);
}

function applyAdminVisibility() {
    if (adminMenuContainer) {
        adminMenuContainer.classList.toggle('hidden', !isAdmin);
    }
}

function updateProfileSummary(profile) {
    if (!profileSummaryClass) return;
    profileSummaryClass.textContent = profile?.classe || '-';
    profileSummaryWarName.textContent = profile?.nomeGuerra || '-';
    profileSummaryRegistration.textContent = profile?.matricula || '-';
    profileSummaryType.textContent = profile?.tipo || '-';
}


function validateRequiredFields() {
    const requiredIds = ['registroNumero', 'dataFato', 'horaFato', 'destinatario', 'postoServico', 'turno', 'relato'];
    let isValid = true;
    let firstInvalid = null;
    requiredIds.forEach(id => {
        const input = document.getElementById(id);
        if (input && input.value.trim() === '') {
            input.classList.add('field-error');
            if (!firstInvalid) firstInvalid = input;
            isValid = false;
        } else if (input) {
            input.classList.remove('field-error');
        }
    });

    const hasAgent = Array.isArray(agentes) && agentes.some(agent => agent.nomeGuerra && agent.nomeGuerra.trim() !== '');
    if (!hasAgent) {
        saveStatus.textContent = 'Adicione pelo menos um agente.';
        isValid = false;
        if (!firstInvalid && addAgenteBtn) firstInvalid = addAgenteBtn;
    }

    if (Array.isArray(partes) && partes.length === 0) {
        if (!noPartesFlag || !noPartesFlag.checked) {
            isValid = false;
            saveStatus.textContent = 'Marque "Não houve envolvidos" ou adicione uma parte.';
            if (noPartesFlagWrapper) noPartesFlagWrapper.classList.add('field-error');
            if (!firstInvalid && noPartesFlag) firstInvalid = noPartesFlag;
        } else if (noPartesFlagWrapper) {
            noPartesFlagWrapper.classList.remove('field-error');
        }
    } else if (noPartesFlagWrapper) {
        noPartesFlagWrapper.classList.remove('field-error');
    }

    if (Array.isArray(apreensoes) && apreensoes.length === 0) {
        if (!noApreensoesFlag || !noApreensoesFlag.checked) {
            isValid = false;
            saveStatus.textContent = 'Marque "Não houve apreensões" ou adicione uma apreensão.';
            if (noApreensoesFlagWrapper) noApreensoesFlagWrapper.classList.add('field-error');
            if (!firstInvalid && noApreensoesFlag) firstInvalid = noApreensoesFlag;
        } else if (noApreensoesFlagWrapper) {
            noApreensoesFlagWrapper.classList.remove('field-error');
        }
    } else if (noApreensoesFlagWrapper) {
        noApreensoesFlagWrapper.classList.remove('field-error');
    }

    const hasAnexos = Array.isArray(anexosUrls) && anexosUrls.length > 0;
    const hasPending = Array.isArray(anexosPending) && anexosPending.length > 0;
    if (!hasAnexos && !hasPending) {
        if (!noAnexosFlag || !noAnexosFlag.checked) {
            isValid = false;
            saveStatus.textContent = 'Marque "Não houve anexos" ou adicione anexos.';
            if (noAnexosFlagWrapper) noAnexosFlagWrapper.classList.add('field-error');
            if (!firstInvalid && noAnexosFlag) firstInvalid = noAnexosFlag;
        } else if (noAnexosFlagWrapper) {
            noAnexosFlagWrapper.classList.remove('field-error');
        }
    } else if (noAnexosFlagWrapper) {
        noAnexosFlagWrapper.classList.remove('field-error');
    }

    if (!isValid && firstInvalid) {
        firstInvalid.scrollIntoView({ behavior: 'smooth', block: 'center' });
        firstInvalid.focus();
    }
    return isValid;
}

function setActiveAppView(viewName) {
    const isForm = viewName === 'form';
    const isManager = viewName === 'manager';
    const isUsers = viewName === 'users';
    const isWarNames = viewName === 'warNames';
    const isLogs = viewName === 'logs';

    if (ropFormView) ropFormView.classList.toggle('hidden', !isForm);
    if (managerView) managerView.classList.toggle('hidden', !isManager);
    if (usersView) usersView.classList.toggle('hidden', !isUsers);
    if (warNamesView) warNamesView.classList.toggle('hidden', !isWarNames);
    if (logsView) logsView.classList.toggle('hidden', !isLogs);

    if (currentRopBtn) currentRopBtn.classList.toggle('active', isForm);
    if (newRopBtn) newRopBtn.classList.toggle('active', false);
    if (viewRopsBtn) viewRopsBtn.classList.toggle('active', isManager);
    if (viewUsersBtn) viewUsersBtn.classList.toggle('active', isUsers);
    if (viewWarNamesBtn) viewWarNamesBtn.classList.toggle('active', isWarNames);
    if (viewLogsBtn) viewLogsBtn.classList.toggle('active', isLogs);

    if (isForm) {
        expandAllTextareas();
    }
}

function confirmNewRop() {
    currentRopId = null;
    clearForm();
}

function updateLastSavedTime() {
    const now = new Date();
    lastSaved.textContent = `Último salvamento: ${formatTime(now)}`;
}

function formatTime(date) {
    const hours = date.getHours().toString().padStart(2, '0');
    const minutes = date.getMinutes().toString().padStart(2, '0');
    const seconds = date.getSeconds().toString().padStart(2, '0');
    return `${hours}:${minutes}:${seconds}`;
}

function triggerSave() {
    if (saveTimeout) clearTimeout(saveTimeout);
    saveTimeout = setTimeout(saveFormData, 500);
}

function saveFormData() {
    const formData = getFormData();
    formData.lastSaved = new Date().toISOString();
    localStorage.setItem('pmuRopForm', JSON.stringify(formData));
    updateLastSavedTime();
    saveStatus.textContent = 'Salvo automaticamente';
}

function loadFormData() {
    const savedData = localStorage.getItem('pmuRopForm');
    if (!savedData) return;
    try {
        const formData = JSON.parse(savedData);
        applyFormData(formData);
        if (formData.lastSaved) {
            const lastSavedDate = new Date(formData.lastSaved);
            lastSaved.textContent = `Último salvamento: ${formatTime(lastSavedDate)}`;
        }
    } catch (error) {
        console.error('Falha ao carregar rascunho local:', error);
    }
}

function clearForm() {
    formInputs.forEach(input => {
        if (input.type === 'checkbox') {
            input.checked = false;
        } else if (input.type !== 'date' && input.type !== 'time') {
            input.value = '';
        }
    });

    const today = new Date().toISOString().split('T')[0];
    const dataFatoInput = document.getElementById('dataFato');
    const horaFatoInput = document.getElementById('horaFato');
    if (dataFatoInput) dataFatoInput.value = today;
    if (horaFatoInput) horaFatoInput.value = '';
    const ufInput = document.getElementById('uf');
    const cidadeInput = document.getElementById('cidade');
    const numeroInput = document.getElementById('numero');
    if (ufInput) ufInput.value = 'SE';
    if (cidadeInput) cidadeInput.value = 'Aracaju';
    if (numeroInput) numeroInput.value = 'S/N';

    partes = [];
    apreensoes = [];
    agentes = [];
    anexosUrls = [];
    anexosPending.forEach(item => {
        if (item?.previewUrl) URL.revokeObjectURL(item.previewUrl);
    });
    anexosPending = [];

    renderPartes();
    renderApreensoes();
    renderAgentes();
    toggleAnexos(false);
    renderAnexosUploadsPreview();
    updateEmptyFlags();

    localStorage.removeItem('pmuRopForm');
    currentRopId = null;
    confirmationModal.style.display = 'none';
    updateLastSavedTime();
    saveStatus.textContent = 'Formulário limpo';
    loadNextRopNumberPreview();
    setTimeout(() => {
        saveStatus.textContent = 'Salvo automaticamente';
    }, 2000);
}

function getFormData() {
    return {
        registroNumero: document.getElementById('registroNumero').value.trim(),
        dataFato: document.getElementById('dataFato').value,
        horaFato: document.getElementById('horaFato').value,
        destinatario: document.getElementById('destinatario').value.trim(),
        noPartesFlag: Boolean(noPartesFlag && noPartesFlag.checked),
        noApreensoesFlag: Boolean(noApreensoesFlag && noApreensoesFlag.checked),
        noAnexosFlag: Boolean(noAnexosFlag && noAnexosFlag.checked),
        flags: {
            prisaoFlagrante: false,
            comunicacaoOcorrencia: false,
            mandadoPrisao: false,
            outros: false
        },
        natureza: document.getElementById('natureza').value,
        tipoDelito: document.getElementById('tipoDelito').value.trim(),
        localOcorrencia: document.getElementById('localOcorrencia').value.trim(),
        uf: document.getElementById('uf').value.trim(),
        cidade: document.getElementById('cidade').value.trim(),
        numero: document.getElementById('numero').value.trim(),
        bairro: document.getElementById('bairro').value.trim(),
        partes: partes.map(parte => ({ ...parte })),
        relato: document.getElementById('relato').value.trim(),
        apreensoes: apreensoes.map(item => ({ ...item })),
        agentes: agentes.map(item => ({ ...item })),
        postoServico: document.getElementById('postoServico').value.trim(),
        turno: document.getElementById('turno').value.trim(),
        anexosUrls: Array.isArray(anexosUrls) ? anexosUrls.slice() : [],
        anexos: Array.isArray(anexosUrls) ? anexosUrls.join('\n') : '',
        createdBy: currentUserProfile ? {
            uid: auth?.currentUser?.uid || '',
            nomeGuerra: currentUserProfile.nomeGuerra || '',
            matricula: currentUserProfile.matricula || '',
            classe: currentUserProfile.classe || ''
        } : null
    };
}

function applyFormData(formData) {
    if (registroNumeroInput) {
        registroNumeroInput.value = formData.registroNumero || '';
    }
    if (registroNumeroDisplay) {
        registroNumeroDisplay.textContent = formData.registroNumero || '--';
    }
    if (!currentRopId && (!formData.registroNumero || formData.registroNumero.trim() === '')) {
        loadNextRopNumberPreview();
    }
    document.getElementById('dataFato').value = formData.dataFato || '';
    document.getElementById('horaFato').value = formData.horaFato || '';
    document.getElementById('destinatario').value = formData.destinatario || '';
    document.getElementById('natureza').value = formData.natureza || '';
    document.getElementById('tipoDelito').value = formData.tipoDelito || '';
    document.getElementById('localOcorrencia').value = formData.localOcorrencia || '';
    document.getElementById('uf').value = formData.uf || 'SE';
    document.getElementById('cidade').value = formData.cidade || 'Aracaju';
    document.getElementById('numero').value = formData.numero || 'S/N';
    document.getElementById('bairro').value = formData.bairro || '';
    document.getElementById('relato').value = formData.relato || '';
    document.getElementById('postoServico').value = formData.postoServico || '';
    document.getElementById('turno').value = formData.turno || '';
    anexosUrls = Array.isArray(formData.anexosUrls) ? formData.anexosUrls.slice() : [];
    if (!anexosUrls.length && formData.anexos) {
        anexosUrls = String(formData.anexos)
            .split(/\r?\n/)
            .map(line => line.trim())
            .filter(Boolean);
    }

    if (noPartesFlag) noPartesFlag.checked = Boolean(formData.noPartesFlag);
    if (noApreensoesFlag) noApreensoesFlag.checked = Boolean(formData.noApreensoesFlag);
    if (noAnexosFlag) noAnexosFlag.checked = Boolean(formData.noAnexosFlag);

    partes = Array.isArray(formData.partes) && formData.partes.length > 0
        ? formData.partes.map(item => {
            const normalized = { ...item };
            if (!normalized.role && normalized.roles) {
                const roleKey = Object.keys(normalized.roles).find(key => normalized.roles[key]);
                const roleMap = {
                    autor: 'Autor',
                    vitima: 'Vítima',
                    testemunha: 'Testemunha',
                    outros: 'Outros'
                };
                normalized.role = roleMap[roleKey] || '';
            }
            delete normalized.roles;
            return normalized;
        })
        : [];
    apreensoes = Array.isArray(formData.apreensoes) && formData.apreensoes.length > 0
        ? formData.apreensoes.map(item => ({ ...item }))
        : [];
    agentes = Array.isArray(formData.agentes) && formData.agentes.length > 0
        ? formData.agentes.map(item => ({ ...item }))
        : [];

    toggleAnexos(Boolean(anexosUrls.length || anexosPending.length));
    renderAnexosUploadsPreview();

    renderPartes();
    renderApreensoes();
    renderAgentes();
    updateEmptyFlags();
    expandAllTextareas();
}

async function saveRopToFirebase(formData) {
    if (!db || !auth || !auth.currentUser) return null;
    try {
        const ropsRef = db.collection('rops');
        let docRef = null;

        const participantMatriculas = buildParticipantMatriculas(formData);
        const payload = {
            ...formData,
            ownerId: auth.currentUser.uid,
            participantMatriculas,
            agentesMatriculas: participantMatriculas,
            agentesNomes: Array.isArray(formData.agentes) ? formData.agentes.map(a => (a?.nomeGuerra || '').trim()).filter(Boolean) : [],
            updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
            updatedBy: currentUserProfile ? {
                uid: auth.currentUser.uid,
                nomeGuerra: currentUserProfile.nomeGuerra,
                matricula: currentUserProfile.matricula,
                classe: currentUserProfile.classe
            } : null
        };

        if (currentRopId) {
            docRef = ropsRef.doc(currentRopId);
            await docRef.set(payload, { merge: true });
        } else {
            docRef = ropsRef.doc();
            const counterRef = db.collection('counters').doc('ropNumber');
            let generatedNumber = '';
            await db.runTransaction(async (transaction) => {
                const counterSnap = await transaction.get(counterRef);
                const currentYear = new Date().getFullYear();
                let currentValue = 0;
                if (counterSnap.exists) {
                    const data = counterSnap.data() || {};
                    if (Number(data.year) === currentYear && Number.isFinite(data.current)) {
                        currentValue = Number(data.current);
                    }
                }
                const nextValue = currentValue + 1;
                generatedNumber = formatRopNumber(nextValue, currentYear);
                transaction.set(counterRef, {
                    year: currentYear,
                    current: nextValue,
                    updatedAt: firebase.firestore.FieldValue.serverTimestamp()
                }, { merge: true });
                transaction.set(docRef, {
                    ...payload,
                    registroNumero: generatedNumber,
                    createdAt: firebase.firestore.FieldValue.serverTimestamp(),
                    createdBy: payload.createdBy || null
                });
            });
            currentRopId = docRef.id;
            if (registroNumeroInput) {
                registroNumeroInput.value = generatedNumber;
            }
            if (registroNumeroDisplay) {
                registroNumeroDisplay.textContent = generatedNumber;
            }
            await loadNextRopNumberPreview();
        }

        await logEvent('rop_saved', { ropId: currentRopId, registroNumero: payload.registroNumero || '' });
        saveStatus.textContent = 'ROP salvo com sucesso!';
        return currentRopId;
    } catch (error) {
        console.error('Erro ao salvar ROP:', error);
        saveStatus.textContent = 'Erro ao salvar ROP.';
        return null;
    }
}

function buildParticipantMatriculas(formData) {
    const set = new Set();
    const agentesList = Array.isArray(formData?.agentes) ? formData.agentes : [];
    agentesList.forEach(agent => {
        const m = (agent?.matricula || '').trim();
        if (m) set.add(m);
    });
    return Array.from(set);
}

function formatRopNumber(sequence, year) {
    return `${String(sequence).padStart(4, '0')}/${year}`;
}

async function loadNextRopNumberPreview() {
    if (!db) return;
    const currentYear = new Date().getFullYear();
    let nextValue = 1;
    try {
        const counterSnap = await db.collection('counters').doc('ropNumber').get();
        if (counterSnap.exists) {
            const data = counterSnap.data() || {};
            if (Number(data.year) === currentYear && Number.isFinite(data.current)) {
                nextValue = Number(data.current) + 1;
            }
        }
    } catch (error) {
        console.error('Erro ao carregar próximo número de ROP:', error);
    }
    const formatted = formatRopNumber(nextValue, currentYear);
    const registroNumeroPreview = document.getElementById('registroNumeroPreview');
    if (registroNumeroPreview) {
        registroNumeroPreview.textContent = `Próximo número: ${formatted}`;
    }
    if (!currentRopId && registroNumeroInput) {
        registroNumeroInput.value = formatted;
    }
    if (!currentRopId && registroNumeroDisplay) {
        registroNumeroDisplay.textContent = formatted;
    }
}

async function logEvent(action, metadata = {}) {
    if (!db || !auth || !auth.currentUser) return;
    try {
        await db.collection('logs').add({
            action,
            userId: auth.currentUser.uid,
            userEmail: auth.currentUser.email || '',
            nomeGuerra: currentUserProfile?.nomeGuerra || '',
            matricula: currentUserProfile?.matricula || '',
            metadata,
            createdAt: firebase.firestore.FieldValue.serverTimestamp()
        });
    } catch (error) {
        console.error('Erro ao salvar log:', error);
    }
}

async function loadRopsList(force = false) {
    if (!db || !auth || !auth.currentUser || isLoadingRops) return;
    if (ropsCache.length > 0 && !force) {
        applyRopsFilter();
        return;
    }

    isLoadingRops = true;
    ropsList.innerHTML = '<div class="operation-empty">Carregando ROPs...</div>';
    try {
        let docs = [];

        if (forcedRopIdFilter) {
            const docSnap = await db.collection('rops').doc(forcedRopIdFilter).get();
            docs = docSnap.exists ? [docSnap] : [];
        } else if (isAdmin) {
            const snapshot = await db.collection('rops').orderBy('dataFato', 'desc').get();
            docs = snapshot.docs;
        } else {
            const registration = (currentUserProfile?.matricula || '').trim();
            if (registration) {
                const snapshot = await db.collection('rops')
                    .where('participantMatriculas', 'array-contains', registration)
                    .get();
                docs = snapshot.docs;
            } else {
                docs = [];
            }
        }

        ropsCache = docs.map(doc => ({ id: doc.id, ...doc.data() }));
        ropsCache.sort((a, b) => {
            const aDate = a.dataFato ? new Date(a.dataFato).getTime() : 0;
            const bDate = b.dataFato ? new Date(b.dataFato).getTime() : 0;
            return bDate - aDate;
        });

        if (isAdmin) {
            // Best-effort: backfill participant arrays so queries/rules work reliably.
            backfillRopsParticipantsIfNeeded(ropsCache);
        }
        ropsCurrentPage = 1;
        applyRopsFilter();
    } catch (error) {
        console.error(error);
        ropsList.innerHTML = '<div class="operation-empty">Falha ao carregar ROPs.</div>';
        ropsPagination.innerHTML = '';
    } finally {
        isLoadingRops = false;
    }
}

async function backfillRopsParticipantsIfNeeded(rops) {
    if (!db || !isAdmin || !Array.isArray(rops) || rops.length === 0) return;
    const missing = rops.filter(r => !Array.isArray(r.participantMatriculas) || r.participantMatriculas.length === 0);
    if (!missing.length) return;

    // Limit per run to avoid large write bursts.
    const slice = missing.slice(0, 60);
    const batch = db.batch();
    let touched = 0;
    slice.forEach(rop => {
        const matriculas = new Set();
        if (Array.isArray(rop.agentes)) {
            rop.agentes.forEach(a => {
                const m = (a?.matricula || '').trim();
                if (m) matriculas.add(m);
            });
        }

        const list = Array.from(matriculas);
        if (!list.length) return;
        const ref = db.collection('rops').doc(rop.id);
        batch.set(ref, { participantMatriculas: list, agentesMatriculas: list }, { merge: true });
        touched += 1;
    });

    if (!touched) return;
    try {
        await batch.commit();
    } catch (e) {
        console.error('Falha ao backfill de participantes:', e);
    }
}

function applyRopsFilter() {
    const textValue = filterText.value.trim().toLowerCase();
    const dateFrom = filterDateFrom.value ? new Date(`${filterDateFrom.value}T00:00:00`) : null;
    const dateTo = filterDateTo.value ? new Date(`${filterDateTo.value}T23:59:59`) : null;

    const filtered = ropsCache.filter(rop => {
        if (forcedRopIdFilter && rop.id !== forcedRopIdFilter) return false;
        if (!isUserAllowedInRop(rop)) return false;

        if (textValue) {
            const blob = buildRopSearchBlob(rop);
            if (!blob.includes(textValue)) return false;
        }

        if (dateFrom || dateTo) {
            const ropDate = rop.dataFato ? new Date(rop.dataFato) : null;
            if (!ropDate) return false;
            if (dateFrom && ropDate < dateFrom) return false;
            if (dateTo && ropDate > dateTo) return false;
        }

        return true;
    });

    renderRopsList(filtered);
}

function buildRopSearchBlob(rop) {
    const agentesText = Array.isArray(rop.agentes)
        ? rop.agentes.map(a => `${a?.nomeGuerra || ''} ${a?.matricula || ''} ${a?.classe || ''}`).join(' ')
        : '';
    const partesText = Array.isArray(rop.partes)
        ? rop.partes.map(p => `${p?.role || ''} ${p?.nome || ''} ${p?.cpf || ''} ${p?.telefone || ''}`).join(' ')
        : '';
    const apreensoesText = Array.isArray(rop.apreensoes)
        ? rop.apreensoes.map(a => `${a?.tipo || ''} ${a?.especificacao || ''}`).join(' ')
        : '';
    const anexosUrlsText = Array.isArray(rop?.anexosUrls)
        ? rop.anexosUrls.join(' ')
        : '';

    const base = [
        rop.registroNumero,
        rop.natureza,
        rop.tipoDelito,
        rop.localOcorrencia,
        rop.dataFato,
        rop.horaFato,
        rop.uf,
        rop.cidade,
        rop.numero,
        rop.bairro,
        rop.destinatario,
        rop.postoServico,
        rop.turno,
        rop.relato,
        rop.createdBy?.nomeGuerra,
        rop.createdBy?.matricula,
        rop.updatedBy?.nomeGuerra,
        rop.updatedBy?.matricula,
        formatDateTime(rop.createdAt),
        formatDateTime(rop.updatedAt),
        agentesText,
        partesText,
        apreensoesText,
        rop.anexos,
        anexosUrlsText
    ]
        .filter(Boolean)
        .join(' ')
        .toLowerCase();
    return base;
}

function countTextMatches(blob, term) {
    if (!blob || !term) return 0;
    let count = 0;
    let index = blob.indexOf(term);
    while (index !== -1) {
        count += 1;
        index = blob.indexOf(term, index + term.length);
    }
    return count;
}

function buildAgentesGroupedLine(agentes) {
    const list = Array.isArray(agentes) ? agentes : [];
    if (!list.length) return 'Não informado';

    const groups = new Map();
    list.forEach(agent => {
        const classe = (agent?.classe || 'Sem classe').trim() || 'Sem classe';
        const nome = (agent?.nomeGuerra || '').trim();
        const matricula = (agent?.matricula || '').trim();
        if (!nome && !matricula) return;
        const display = [nome, matricula ? `(${matricula})` : ''].filter(Boolean).join(' ').trim();
        if (!display) return;
        if (!groups.has(classe)) groups.set(classe, []);
        groups.get(classe).push(display);
    });

    if (!groups.size) return 'Não informado';

    return Array.from(groups.entries())
        .sort(([a], [b]) => a.localeCompare(b, 'pt-BR'))
        .map(([classe, names]) => `${classe}: ${names.sort((a, b) => a.localeCompare(b, 'pt-BR')).join(', ')}`)
        .join(' | ');
}

function renderRopsList(rops) {
    if (!rops.length) {
        ropsList.innerHTML = '<div class="operation-empty">Nenhum ROP encontrado.</div>';
        ropsPagination.innerHTML = '';
        return;
    }

    const totalPages = Math.ceil(rops.length / ROPS_PAGE_SIZE);
    if (ropsCurrentPage > totalPages) ropsCurrentPage = totalPages;
    const startIndex = (ropsCurrentPage - 1) * ROPS_PAGE_SIZE;
    const pageItems = rops.slice(startIndex, startIndex + ROPS_PAGE_SIZE);

    ropsList.innerHTML = '';
    const searchText = filterText ? filterText.value.trim().toLowerCase() : '';
    pageItems.forEach(rop => {
        const highlights = [
            { icon: 'fa-gavel', label: 'Delito', value: rop.tipoDelito || 'Não informado' },
            { icon: 'fa-location-dot', label: 'Local', value: rop.localOcorrencia || 'Não informado' },
            { icon: 'fa-calendar-alt', label: 'Data/Hora', value: `${formatDate(rop.dataFato)} ${rop.horaFato || ''}`.trim() }
            // Removido o destaque de Posto de Serviço
        ];

        let matchBadge = '';
        if (searchText) {
            const blob = buildRopSearchBlob(rop);
            const matchCount = countTextMatches(blob, searchText);
            matchBadge = `
                <div class="operation-match" title="Matches da busca">
                    <i class="fas fa-magnifying-glass"></i> ${matchCount}
                </div>
            `;
        }

        const agentesLine = buildAgentesGroupedLine(rop.agentes);

        const card = document.createElement('div');
        card.className = 'operation-card';
        card.innerHTML = `
            <div class="operation-title-row">
                <div class="operation-title">ROP Nº ${rop.registroNumero || 'Sem número'}</div>
                <div class="operation-highlights">
                    ${highlights.map(h => `
                        <span class="highlight-pill"><i class="fas ${h.icon}"></i> ${h.label}: ${escapeHtml(h.value)}</span>
                    `).join('')}
                </div>
                ${matchBadge}
            </div>
            <div class="operation-meta">
                <div class="operation-meta-item"><strong>Natureza:</strong> ${formatNatureza(rop.natureza) || 'Não informado'}</div>
                <div class="operation-meta-item"><strong>Posto de serviço:</strong> ${escapeHtml(rop.postoServico || 'Não informado')}</div>
                <div class="operation-meta-item"><strong>Registrado por:</strong> ${rop.createdBy?.nomeGuerra || 'Não informado'}</div>
                <div class="operation-meta-item"><strong>Criado em:</strong> ${formatDateTime(rop.createdAt)}</div>
                <div class="operation-meta-item"><strong>Atualizado em:</strong> ${formatDateTime(rop.updatedAt)}</div>
                <div class="operation-meta-item"><strong>Atualizado por:</strong> ${rop.updatedBy?.nomeGuerra || 'Não informado'}</div>
                <div class="operation-meta-item operation-meta-agentes"><strong>Agentes na ocorrência:</strong> ${escapeHtml(agentesLine)}</div>
            </div>
            <div class="operation-actions">
                <button class="btn-small" data-action="edit" data-id="${rop.id}">
                    <i class="fas fa-pen"></i> Editar
                </button>
                <button class="btn-small" data-action="pdf" data-id="${rop.id}">
                    <i class="fas fa-file-pdf"></i> Gerar PDF
                </button>
                ${isAdmin ? `
                    <button class="btn-small btn-danger" data-action="delete" data-id="${rop.id}">
                        <i class="fas fa-trash"></i> Excluir
                    </button>
                ` : ''}
            </div>
        `;
        ropsList.appendChild(card);
    });

    ropsList.querySelectorAll('button[data-action="edit"]').forEach(button => {
        button.addEventListener('click', () => {
            const id = button.getAttribute('data-id');
            const rop = ropsCache.find(item => item.id === id);
            const permission = getRopEditPermission(rop);
            if (!permission.allowed) {
                alert(permission.message);
                return;
            }
            if (rop) loadRopById(id);
        });
    });

    ropsList.querySelectorAll('button[data-action="pdf"]').forEach(button => {
        button.addEventListener('click', async () => {
            if (isGeneratingPdf) return;
            const id = button.getAttribute('data-id');
            const rop = ropsCache.find(item => item.id === id);
            if (!rop) return;
            isGeneratingPdf = true;
            setPdfBusy(true, button);
            try {
                await exportToPDF(rop);
            } finally {
                isGeneratingPdf = false;
                setPdfBusy(false, button);
            }
        });
    });

    ropsList.querySelectorAll('button[data-action="delete"]').forEach(button => {
        button.addEventListener('click', () => {
            const id = button.getAttribute('data-id');
            deleteRopById(id);
        });
    });

    renderRopsPagination(totalPages);
}

function isEditBlockedByDate(rop) {
    const createdAt = getRopCreatedAtDate(rop);
    if (!createdAt) return false;
    const now = new Date();
    const diffMs = now.getTime() - createdAt.getTime();
    const twoDaysMs = 2 * 24 * 60 * 60 * 1000;
    return diffMs >= twoDaysMs;
}

function getRopEditPermission(rop) {
    if (!rop) return { allowed: false, message: 'ROP não encontrado.' };
    const createdByUid = rop.createdBy?.uid || '';
    const currentUid = auth?.currentUser?.uid || '';
    const isOwner = createdByUid && currentUid && createdByUid === currentUid;
    if (!isOwner) {
        return { allowed: false, message: 'Só quem criou poderá editar e no máximo em até dois dias.' };
    }
    if (isEditBlockedByDate(rop)) {
        return { allowed: false, message: 'Só quem criou poderá editar e no máximo em até dois dias.' };
    }
    return { allowed: true, message: '' };
}

function getRopCreatedAtDate(rop) {
    if (!rop?.createdAt) return null;
    if (rop.createdAt.toDate) return rop.createdAt.toDate();
    const date = new Date(rop.createdAt);
    return Number.isNaN(date.getTime()) ? null : date;
}

function escapeHtml(value) {
    const str = String(value ?? '');
    return str
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#039;');
}

function renderRopsPagination(totalPages) {
    ropsPagination.innerHTML = '';
    if (totalPages <= 1) return;

    const prevBtn = document.createElement('button');
    prevBtn.className = 'pagination-btn';
    prevBtn.textContent = 'Anterior';
    prevBtn.disabled = ropsCurrentPage === 1;
    prevBtn.addEventListener('click', () => {
        ropsCurrentPage -= 1;
        applyRopsFilter();
    });
    ropsPagination.appendChild(prevBtn);

    for (let i = 1; i <= totalPages; i += 1) {
        const pageBtn = document.createElement('button');
        pageBtn.className = 'pagination-btn';
        if (i === ropsCurrentPage) pageBtn.classList.add('active');
        pageBtn.textContent = String(i);
        pageBtn.addEventListener('click', () => {
            ropsCurrentPage = i;
            applyRopsFilter();
        });
        ropsPagination.appendChild(pageBtn);
    }

    const nextBtn = document.createElement('button');
    nextBtn.className = 'pagination-btn';
    nextBtn.textContent = 'Próximo';
    nextBtn.disabled = ropsCurrentPage === totalPages;
    nextBtn.addEventListener('click', () => {
        ropsCurrentPage += 1;
        applyRopsFilter();
    });
    ropsPagination.appendChild(nextBtn);
}

async function loadRopById(id) {
    if (!db || !auth || !auth.currentUser) return;
    try {
        const docRef = db.collection('rops').doc(id);
        const docSnap = await docRef.get();
        if (docSnap.exists) {
            const data = docSnap.data();
            if (!isUserAllowedInRop(data)) {
                saveStatus.textContent = 'Você não tem permissão para abrir este ROP.';
                alert('Você não tem permissão para abrir este ROP.');
                return;
            }
            currentRopId = docSnap.id;
            applyFormData(data);
            setActiveAppView('form');
            saveStatus.textContent = 'ROP carregado';
        }
    } catch (error) {
        console.error('Erro ao abrir ROP:', error);
    }
}

async function deleteRopById(id) {
    if (!db || !auth || !auth.currentUser || !isAdmin || !id) return;
    const rop = ropsCache.find(item => item.id === id);
    const shouldDelete = window.confirm('Deseja realmente excluir este ROP? Esta ação não pode ser desfeita.');
    if (!shouldDelete) return;
    try {
        await db.collection('rops').doc(id).delete();
        await logEvent('rop_deleted', {
            ropId: id,
            registroNumero: rop?.registroNumero || ''
        });
        ropsCache = ropsCache.filter(item => item.id !== id);
        applyRopsFilter();
    } catch (error) {
        console.error('Erro ao excluir ROP:', error);
        alert('Erro ao excluir ROP.');
    }
}

function isUserAllowedInRop(rop) {
    if (isAdmin) return true;
    const registration = currentUserProfile?.matricula;
    if (!registration) return false;
    if (registration) {
        if (Array.isArray(rop.participantMatriculas) && rop.participantMatriculas.includes(registration)) return true;
        if (Array.isArray(rop.agentesMatriculas) && rop.agentesMatriculas.includes(registration)) return true;
        if (Array.isArray(rop.agentes)) {
            return rop.agentes.some(agent => agent.matricula && agent.matricula === registration);
        }
    }
    return false;
}

function addParte() {
    openParteModal();
}

function renderPartes() {
    if (!partesList) return;
    partesList.innerHTML = '';
    updateEmptyFlags();
    if (!partes.length) return;

    partes.forEach(parte => {
        const item = document.createElement('div');
        item.className = 'parte-item';
        const roleClass = (parte.role || '').toLowerCase().replace(/[^a-z]/g, '');
        item.innerHTML = `
            <div class="parte-info">
                <span class="parte-role parte-role-${roleClass || 'default'}">${parte.role || 'Sem tipo'}</span>
                <strong>${parte.nome || 'Sem nome'}</strong>
                <span>CPF: ${parte.cpf || '-'} | Nasc.: ${parte.dataNascimento || '-'} | Tel: ${parte.telefone || '-'}</span>
                <span>Endereço: ${parte.endereco || '-'} | Cidade/UF: ${parte.cidadeUf || '-'}</span>
            </div>
            <div class="parte-actions">
                <button class="btn-small" data-action="edit-parte" data-id="${parte.id}">Editar</button>
                <button class="btn-small btn-danger" data-action="remove-parte" data-id="${parte.id}">Excluir</button>
            </div>
        `;
        partesList.appendChild(item);
    });

    partesList.querySelectorAll('button[data-action="edit-parte"]').forEach(button => {
        button.addEventListener('click', () => {
            const id = Number(button.getAttribute('data-id'));
            openParteModal(id);
        });
    });

    partesList.querySelectorAll('button[data-action="remove-parte"]').forEach(button => {
        button.addEventListener('click', () => {
            const id = Number(button.getAttribute('data-id'));
            partes = partes.filter(parte => parte.id !== id);
            renderPartes();
            updateEmptyFlags();
            triggerSave();
        });
    });
}

function openParteModal(id = null) {
    if (!parteModal) return;
    editingParteId = id;
    if (id) {
        const parte = partes.find(item => item.id === id);
        fillParteModal(parte);
    } else {
        resetParteModal();
    }
    parteModal.style.display = 'flex';
}

function closeParteModal() {
    if (!parteModal) return;
    parteModal.style.display = 'none';
    resetParteModal();
    editingParteId = null;
}

function confirmCloseParteModal() {
    const shouldClose = window.confirm('Se fechar, os dados não serão salvos. Deseja sair mesmo?');
    if (shouldClose) {
        closeParteModal();
    }
}

function resetParteModal() {
    if (!parteForm) return;
    parteForm.reset();
    if (parteCondicao) parteCondicao.value = 'sem_ferimentos';
    if (parteSexo) parteSexo.value = 'masculino';
}

function fillParteModal(parte) {
    if (!parte) return;
    if (parteRole) parteRole.value = parte.role || '';
    if (parteNome) parteNome.value = parte.nome || '';
    if (parteNascimento) parteNascimento.value = parte.dataNascimento || '';
    if (partePai) partePai.value = parte.pai || '';
    if (parteMae) parteMae.value = parte.mae || '';
    if (parteCondicao) parteCondicao.value = parte.condicaoFisica || 'sem_ferimentos';
    if (parteSexo) parteSexo.value = parte.sexo || 'masculino';
    if (parteNaturalidade) parteNaturalidade.value = parte.naturalidade || '';
    if (parteEndereco) parteEndereco.value = parte.endereco || '';
    if (parteCpf) parteCpf.value = parte.cpf || '';
    if (parteCidadeUf) parteCidadeUf.value = parte.cidadeUf || '';
    if (parteTelefone) parteTelefone.value = parte.telefone || '';
}

function saveParteFromModal() {
    if (!parteRole || !parteNome) return;
    if (!parteRole.value || !parteNome.value.trim()) {
        alert('Preencha os campos obrigatórios: Tipo e Nome.');
        return;
    }

    const cpfRaw = (parteCpf?.value || '').replace(/\D/g, '');
    if (cpfRaw) {
        if (cpfRaw.length !== 11 || !isValidCPF(cpfRaw)) {
            alert('CPF inválido.');
            return;
        }
    }

    if (parteNascimento?.value) {
        const birth = new Date(`${parteNascimento.value}T00:00:00`);
        if (Number.isNaN(birth.getTime()) || birth > new Date()) {
            alert('Data de nascimento inválida.');
            return;
        }
    }

    const phoneRaw = (parteTelefone?.value || '').replace(/\D/g, '');
    if (phoneRaw && phoneRaw.length < 8) {
        alert('Telefone inválido.');
        return;
    }

    const payload = {
        id: editingParteId || Date.now() + Math.random(),
        role: parteRole.value,
        nome: parteNome.value.trim(),
        dataNascimento: parteNascimento?.value || '',
        pai: partePai?.value.trim() || '',
        mae: parteMae?.value.trim() || '',
        condicaoFisica: parteCondicao?.value || 'sem_ferimentos',
        sexo: parteSexo?.value || 'masculino',
        naturalidade: parteNaturalidade?.value.trim() || '',
        endereco: parteEndereco?.value.trim() || '',
        cpf: cpfRaw ? cpfRaw : '',
        cidadeUf: parteCidadeUf?.value.trim() || '',
        telefone: phoneRaw ? phoneRaw : ''
    };

    if (editingParteId) {
        partes = partes.map(item => item.id === editingParteId ? payload : item);
    } else {
        partes.push(payload);
    }

    renderPartes();
    triggerSave();
    closeParteModal();
}

function addApreensao() {
    apreensoes.push({ id: Date.now() + Math.random(), tipo: '', especificacao: '' });
    renderApreensoes();
    updateEmptyFlags();
}

function renderApreensoes() {
    if (!apreensoesTableBody) return;
    apreensoesTableBody.innerHTML = '';
    const hasItems = apreensoes.length > 0;
    if (apreensoesContainer) apreensoesContainer.classList.toggle('hidden', !hasItems);
    if (apreensoesEmptyMessage) apreensoesEmptyMessage.classList.toggle('hidden', hasItems);

    apreensoes.forEach(item => {
        const tipoValue = (item.tipo || '').trim();
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>
                <select data-id="${item.id}" data-field="tipo">
                    ${buildApreensaoTipoOptions(tipoValue)}
                </select>
            </td>
            <td><input type="text" data-id="${item.id}" data-field="especificacao" value="${item.especificacao || ''}"></td>
            <td>
                <button class="btn-small btn-danger" data-action="remove-apreensao" data-id="${item.id}">Remover</button>
            </td>
        `;
        apreensoesTableBody.appendChild(row);
    });

    apreensoesTableBody.querySelectorAll('button[data-action="remove-apreensao"]').forEach(button => {
        button.addEventListener('click', () => {
            const id = Number(button.getAttribute('data-id'));
            apreensoes = apreensoes.filter(item => item.id !== id);
            renderApreensoes();
            updateEmptyFlags();
            triggerSave();
        });
    });
}

function buildApreensaoTipoOptions(selectedValue = '') {
    const options = ['Objeto', 'Arma', 'Veículo', 'Substância', 'Dinheiro', 'Outros'];
    const safeSelected = (selectedValue || '').toLowerCase();
    return ['<option value="">Selecione</option>']
        .concat(options.map(opt => {
            const isSelected = opt.toLowerCase() === safeSelected ? ' selected' : '';
            return `<option value="${opt}"${isSelected}>${opt}</option>`;
        }))
        .join('');
}

function handleApreensaoInputChange(event) {
    const target = event.target;
    const id = Number(target.getAttribute('data-id'));
    const field = target.getAttribute('data-field');
    if (!id || !field) return;
    const item = apreensoes.find(row => row.id === id);
    if (!item) return;
    item[field] = target.value;
    triggerSave();
}

function addAgente() {
    agentes.push({ id: Date.now() + Math.random(), classe: '', nomeGuerra: '', matricula: '' });
    renderAgentes();
}

function renderAgentes() {
    if (!agentesTableBody) return;
    agentesTableBody.innerHTML = '';
    if (agentesContainer && agentesEmpty) {
        const hasItems = agentes.length > 0;
        agentesContainer.classList.toggle('hidden', !hasItems);
        agentesEmpty.classList.toggle('hidden', hasItems);
    }
    agentes.forEach(agent => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td><div class="readonly-value compact" data-field="classe">${agent.classe || '-'}</div></td>
            <td>
                <select data-id="${agent.id}" data-field="nomeGuerra">
                    ${buildWarNameOptions(agent.nomeGuerra || '')}
                </select>
            </td>
            <td><div class="readonly-value compact" data-field="matricula">${agent.matricula || '-'}</div></td>
            <td>
                <button class="btn-small btn-danger" data-action="remove-agente" data-id="${agent.id}">Remover</button>
            </td>
        `;
        agentesTableBody.appendChild(row);
    });

    agentesTableBody.querySelectorAll('button[data-action="remove-agente"]').forEach(button => {
        button.addEventListener('click', () => {
            const id = Number(button.getAttribute('data-id'));
            agentes = agentes.filter(agent => agent.id !== id);
            renderAgentes();
            triggerSave();
        });
    });
}

function toggleAnexos(show) {
    if (anexosContainer) {
        anexosContainer.classList.toggle('hidden', !show);
    }
    if (anexosEmpty) {
        anexosEmpty.classList.toggle('hidden', show);
    }
    updateEmptyFlags();
    if (show) {
        renderAnexosUploadsPreview();
    }
}

function handleAddAnexoClick() {
    toggleAnexos(true);
    if (!anexosFileInput) {
        alert('Input de upload não encontrado.');
        return;
    }
    if (!currentRopId) {
        saveStatus.textContent = 'Imagens serão enviadas após registrar o ROP.';
    }
    anexosFileInput.click();
}

async function uploadAnexoImage(file) {
    if (!file) return;

    if (!file.type || !file.type.startsWith('image/')) {
        alert('Envie apenas arquivos de imagem.');
        return;
    }
    const maxBytes = 5 * 1024 * 1024;
    if (file.size > maxBytes) {
        alert('A imagem deve ter no máximo 5MB.');
        return;
    }

    const previewUrl = URL.createObjectURL(file);
    anexosPending.push({ id: Date.now() + Math.random(), file, previewUrl });
    renderAnexosUploadsPreview();
    updateEmptyFlags();

    if (currentRopId) {
        await uploadPendingAnexosToStorage(currentRopId);
    }
}

async function uploadPendingAnexosToStorage(ropId) {
    if (!ropId || !storage || !auth || !auth.currentUser) return;
    if (!anexosPending.length) return;

    try {
        isUploadingAnexo = true;
        saveStatus.textContent = 'Enviando imagens...';
        const uid = auth.currentUser.uid;
        const uploads = [];

        for (const item of anexosPending) {
            const file = item.file;
            if (!file) continue;
            const safeName = (file.name || 'imagem').replace(/[^a-zA-Z0-9._-]/g, '_');
            const path = `anexos/${ropId}/${uid}/${Date.now()}_${safeName}`;
            const ref = storage.ref().child(path);
            await ref.put(file, { contentType: file.type });
            const url = await ref.getDownloadURL();
            uploads.push(url);
        }

        if (uploads.length) {
            anexosUrls = Array.isArray(anexosUrls) ? anexosUrls.concat(uploads) : uploads.slice();
            await db.collection('rops').doc(ropId).set({
                anexosUrls,
                anexos: anexosUrls.join('\n'),
                updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
                updatedBy: currentUserProfile ? {
                    uid: auth.currentUser.uid,
                    nomeGuerra: currentUserProfile.nomeGuerra,
                    matricula: currentUserProfile.matricula,
                    classe: currentUserProfile.classe
                } : null
            }, { merge: true });
        }

        anexosPending.forEach(item => {
            if (item?.previewUrl) URL.revokeObjectURL(item.previewUrl);
        });
        anexosPending = [];

        renderAnexosUploadsPreview();
        updateEmptyFlags();
        saveStatus.textContent = 'Imagens anexadas com sucesso.';
        setTimeout(() => {
            if (saveStatus) saveStatus.textContent = 'Salvo automaticamente';
        }, 2500);
    } catch (error) {
        console.error(error);
        alert('Falha ao enviar as imagens.');
        saveStatus.textContent = 'Falha ao enviar imagens.';
    } finally {
        isUploadingAnexo = false;
    }
}

function renderAnexosUploadsPreview() {
    if (!anexosUploadsList) return;
    const urls = Array.isArray(anexosUrls) ? anexosUrls : [];
    const pending = Array.isArray(anexosPending) ? anexosPending : [];

    if (!urls.length && !pending.length) {
        anexosUploadsList.innerHTML = '';
        if (anexosEmpty) anexosEmpty.classList.remove('hidden');
        return;
    }

    if (anexosEmpty) anexosEmpty.classList.add('hidden');

    const uploadedCards = urls.map(url => {
        return `
            <div class="upload-item">
                <a href="${url}" target="_blank" rel="noreferrer">
                    <img class="upload-thumb" src="${url}" alt="Anexo" loading="lazy" />
                </a>
                <div class="upload-meta">Imagem registrada</div>
            </div>
        `;
    });

    const pendingCards = pending.map(item => {
        return `
            <div class="upload-item" data-pending-id="${item.id}">
                <img class="upload-thumb" src="${item.previewUrl}" alt="Pré-visualização" loading="lazy" />
                <div class="upload-meta">Pendente de registro</div>
                <div class="upload-actions">
                    <button class="upload-remove" data-action="remove-pending" data-id="${item.id}">Excluir</button>
                </div>
            </div>
        `;
    });

    anexosUploadsList.innerHTML = uploadedCards.concat(pendingCards).join('');

    anexosUploadsList.querySelectorAll('button[data-action="remove-pending"]').forEach(button => {
        button.addEventListener('click', () => {
            const id = Number(button.getAttribute('data-id'));
            const item = anexosPending.find(p => p.id === id);
            if (item?.previewUrl) URL.revokeObjectURL(item.previewUrl);
            anexosPending = anexosPending.filter(p => p.id !== id);
            renderAnexosUploadsPreview();
            updateEmptyFlags();
        });
    });
}

function handleAgenteInputChange(event) {
    const target = event.target;
    const id = Number(target.getAttribute('data-id'));
    const field = target.getAttribute('data-field');
    if (!id || !field) return;
    const agent = agentes.find(row => row.id === id);
    if (!agent) return;

    if (field === 'nomeGuerra') {
        const nextValue = target.value;
        const normalized = nextValue.trim().toLowerCase();
        const duplicated = agentes.some(row => row.id !== id && (row.nomeGuerra || '').trim().toLowerCase() === normalized && normalized !== '');
        if (duplicated) {
            alert('Este agente já foi adicionado à ocorrência.');
            target.value = agent.nomeGuerra || '';
            return;
        }
        agent[field] = nextValue;
        const match = findWarNameByValue(nextValue);
        if (match) {
            agent.classe = match.classe || agent.classe;
            agent.matricula = match.matricula || agent.matricula;
            renderAgentes();
        }
    } else {
        agent[field] = target.value;
    }
    triggerSave();
}

function autoExpandTextarea(textarea) {
    textarea.style.height = 'auto';
    textarea.style.height = `${textarea.scrollHeight}px`;
}

function expandAllTextareas() {
    const run = () => {
        document.querySelectorAll('#ropFormView textarea').forEach(textarea => {
            // Skip hidden elements (scrollHeight can be 0)
            if (textarea.offsetParent === null) return;
            autoExpandTextarea(textarea);
        });
    };

    // Run after layout settles (fragments + hidden->visible transitions)
    requestAnimationFrame(() => requestAnimationFrame(run));
}

function maskCPF(value) {
    const digits = String(value || '').replace(/\D/g, '').slice(0, 11);
    const p1 = digits.slice(0, 3);
    const p2 = digits.slice(3, 6);
    const p3 = digits.slice(6, 9);
    const p4 = digits.slice(9, 11);
    let out = p1;
    if (p2) out += `.${p2}`;
    if (p3) out += `.${p3}`;
    if (p4) out += `-${p4}`;
    return out;
}

function maskTelefone(value) {
    const digits = String(value || '').replace(/\D/g, '').slice(0, 11);
    if (!digits) return '';
    const ddd = digits.slice(0, 2);
    const rest = digits.slice(2);
    if (rest.length <= 4) return `(${ddd}) ${rest}`.trim();
    if (rest.length <= 8) {
        return `(${ddd}) ${rest.slice(0, 4)}-${rest.slice(4)}`.trim();
    }
    return `(${ddd}) ${rest.slice(0, 5)}-${rest.slice(5)}`.trim();
}

function isValidCPF(cpf) {
    const digits = String(cpf || '').replace(/\D/g, '');
    if (digits.length !== 11) return false;
    if (/^(\d)\1{10}$/.test(digits)) return false;

    const calc = (base, factor) => {
        let total = 0;
        for (let i = 0; i < base.length; i += 1) {
            total += Number(base[i]) * (factor - i);
        }
        const mod = total % 11;
        return mod < 2 ? 0 : 11 - mod;
    };

    const d1 = calc(digits.slice(0, 9), 10);
    if (d1 !== Number(digits[9])) return false;
    const d2 = calc(digits.slice(0, 10), 11);
    return d2 === Number(digits[10]);
}

function formatDate(dateValue) {
    if (!dateValue) return 'Não informado';
    if (dateValue.toDate) return dateValue.toDate().toLocaleDateString('pt-BR');
    const date = new Date(dateValue);
    if (Number.isNaN(date.getTime())) return 'Não informado';
    return date.toLocaleDateString('pt-BR');
}

function formatDateTime(dateValue) {
    if (!dateValue) return 'Não informado';
    if (dateValue.toDate) return dateValue.toDate().toLocaleString('pt-BR');
    const date = new Date(dateValue);
    if (Number.isNaN(date.getTime())) return 'Não informado';
    return date.toLocaleString('pt-BR');
}

function formatNatureza(value) {
    const map = {
        prisao_em_flagrante: 'Prisão em Flagrante',
        apreensao_em_flagrante: 'Apreensão em Flagrante',
        prisao_e_apreensao_em_flagrante: 'Prisão e Apreensão em Flagrante',
        mandado_de_prisao_em_aberto: 'Mandado de prisão em aberto',
        comunicado_de_fato_relevante: 'Comunicado de fato relevante',
        termo_circunstanciado_de_ocorrencia: 'Termo Circunstanciado de Ocorrência',
        outros: 'Outros'
    };
    return map[value] || value || '';
}

async function importWarNamesFromSeed() {
    if (!isAdmin) {
        alert('Apenas administradores podem importar a base.');
        return;
    }
    if (!db) return;
    const confirmImport = window.confirm('Deseja importar a base do seed? Os dados existentes serão atualizados por matrícula.');
    if (!confirmImport) return;

    try {
        saveStatus.textContent = 'Importando base de agentes...';
        const response = await fetch('seed-war-names.json', { cache: 'no-store' });
        if (!response.ok) throw new Error('Falha ao carregar seed-war-names.json');
        const data = await response.json();
        const list = Array.isArray(data) && Array.isArray(data[0]) ? data.flat() : data;
        const entries = (Array.isArray(list) ? list : []).filter(item => item && item.nomeGuerra);
        const chunkSize = 450;

        for (let i = 0; i < entries.length; i += chunkSize) {
            const batch = db.batch();
            const slice = entries.slice(i, i + chunkSize);
            slice.forEach(item => {
                const matricula = item.matricula ? String(item.matricula).trim() : '';
                const docRef = matricula
                    ? db.collection('warNames').doc(matricula)
                    : db.collection('warNames').doc();
                batch.set(docRef, {
                    classe: item.classe || '',
                    nomeGuerra: (item.nomeGuerra || '').trim(),
                    matricula: matricula,
                    status: item.status || 'ativo',
                    updatedAt: firebase.firestore.FieldValue.serverTimestamp()
                }, { merge: true });
            });
            await batch.commit();
        }

        saveStatus.textContent = 'Base importada com sucesso.';
        await loadWarNamesList(true);
    } catch (error) {
        console.error(error);
        saveStatus.textContent = 'Falha ao importar a base.';
        alert('Erro ao importar a base. Verifique o arquivo seed-war-names.json.');
    }
}

async function loadUsersList(force = false) {
    if (!db || !auth || !auth.currentUser || !isAdmin || isLoadingUsers) return;
    if (usersCache.length > 0 && !force) {
        renderUsersTable();
        return;
    }
    isLoadingUsers = true;
    try {
        const snapshot = await db.collection('users').orderBy('nomeGuerra').get();
        usersCache = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        renderUsersTable();
        updateUserCreateWarNameSelect(userCreateWarName?.value || '');
    } catch (error) {
        console.error(error);
    } finally {
        isLoadingUsers = false;
    }
}

function renderUsersTable() {
    if (!usersTableBody) return;
    usersTableBody.innerHTML = '';
    usersCache.forEach(user => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${user.classe || '-'}</td>
            <td>${user.nomeGuerra || '-'}</td>
            <td>${user.email || '-'}</td>
            <td>${user.matricula || '-'}</td>
            <td>${user.tipo || '-'}</td>
            <td>
                <button class="btn-small" data-action="edit-user" data-id="${user.id}">Editar</button>
                <button class="btn-small" data-action="reset-user" data-id="${user.id}">Resetar senha</button>
                <button class="btn-small btn-danger" data-action="delete-user" data-id="${user.id}">Excluir</button>
            </td>
        `;
        usersTableBody.appendChild(row);
    });

    usersTableBody.querySelectorAll('button[data-action="edit-user"]').forEach(button => {
        button.addEventListener('click', () => {
            const id = button.getAttribute('data-id');
            const user = usersCache.find(item => item.id === id);
            if (user) openUserEditModal(user);
        });
    });

    usersTableBody.querySelectorAll('button[data-action="delete-user"]').forEach(button => {
        button.addEventListener('click', () => {
            const id = button.getAttribute('data-id');
            const user = usersCache.find(item => item.id === id);
            if (user) deleteUserProfile(user);
        });
    });

    usersTableBody.querySelectorAll('button[data-action="reset-user"]').forEach(button => {
        button.addEventListener('click', () => {
            const id = button.getAttribute('data-id');
            const user = usersCache.find(item => item.id === id);
            if (user) openResetPasswordModal(user);
        });
    });
}

function openUserCreateModal() {
    if (!userCreateModal) return;
    userCreateError.textContent = '';
    userCreateForm.reset();
    updateUserCreateWarNameSelect();
    if (userCreateClass) userCreateClass.value = '';
    if (userCreateRegistration) userCreateRegistration.value = '';
    userCreateModal.style.display = 'flex';
}

function closeUserCreateModal() {
    if (userCreateModal) userCreateModal.style.display = 'none';
}

async function createNewUser() {
    if (!isAdmin) return;
    userCreateError.textContent = '';
    if (!userCreateEmail.value || !userCreatePassword.value || !userCreateClass.value || !userCreateWarName.value || !userCreateRegistration.value) {
        userCreateError.textContent = 'Preencha todos os campos obrigatórios.';
        return;
    }

    try {
        const secondaryApp = firebase.apps.find(app => app.name === 'secondary') || firebase.initializeApp(firebaseConfig, 'secondary');
        const secondaryAuth = secondaryApp.auth();
        const credential = await secondaryAuth.createUserWithEmailAndPassword(userCreateEmail.value.trim(), userCreatePassword.value);
        const newUser = credential.user;
        await db.collection('users').doc(newUser.uid).set({
            email: userCreateEmail.value.trim(),
            classe: userCreateClass.value,
            nomeGuerra: userCreateWarName.value.trim(),
            matricula: userCreateRegistration.value.trim(),
            tipo: userCreateType.value
        }, { merge: true });
        await logEvent('user_created', {
            userId: newUser.uid,
            email: userCreateEmail.value.trim(),
            nomeGuerra: userCreateWarName.value.trim(),
            matricula: userCreateRegistration.value.trim(),
            tipo: userCreateType.value
        });
        await secondaryAuth.signOut();
        closeUserCreateModal();
        loadUsersList(true);
    } catch (error) {
        console.error(error);
        userCreateError.textContent = 'Erro ao criar usuário.';
    }
}

function openUserEditModal(user) {
    if (!userEditModal) return;
    editingUserId = user.id;
    userEditError.textContent = '';
    userEditClass.value = user.classe || '';
    userEditWarName.value = user.nomeGuerra || '';
    userEditRegistration.value = user.matricula || '';
    userEditType.value = user.tipo || 'usuario';
    userEditModal.style.display = 'flex';
}

function openResetPasswordModal(user) {
    if (!resetPasswordModal) return;
    resetPasswordError.textContent = '';
    pendingResetUser = user || null;
    resetPasswordName.value = user?.nomeGuerra || 'Não informado';
    resetPasswordEmail.value = user?.email || '';
    resetPasswordModal.style.display = 'flex';
}

function closeResetPasswordModal() {
    if (resetPasswordModal) resetPasswordModal.style.display = 'none';
    pendingResetUser = null;
}

async function sendPasswordResetForUser() {
    if (!auth || !isAdmin || !pendingResetUser) return;
    const email = (pendingResetUser.email || '').trim();
    if (!email) {
        resetPasswordError.textContent = 'E-mail do usuário não encontrado.';
        return;
    }
    try {
        await auth.sendPasswordResetEmail(email);
        await logEvent('user_password_reset', {
            email,
            userId: pendingResetUser.id || '',
            nomeGuerra: pendingResetUser.nomeGuerra || '',
            matricula: pendingResetUser.matricula || ''
        });
        closeResetPasswordModal();
        alert('E-mail de reset enviado com sucesso.');
    } catch (error) {
        console.error(error);
        resetPasswordError.textContent = 'Não foi possível enviar o e-mail de reset.';
    }
}

function closeUserEditModal() {
    if (userEditModal) userEditModal.style.display = 'none';
    editingUserId = null;
}

async function saveEditedUserProfile() {
    if (!editingUserId || !db || !isAdmin) return;
    userEditError.textContent = '';
    if (!userEditClass.value || !userEditWarName.value || !userEditRegistration.value || !userEditType.value) {
        userEditError.textContent = 'Preencha todos os campos obrigatórios.';
        return;
    }
    try {
        await db.collection('users').doc(editingUserId).set({
            classe: userEditClass.value,
            nomeGuerra: userEditWarName.value.trim(),
            matricula: userEditRegistration.value.trim(),
            tipo: userEditType.value
        }, { merge: true });
        await logEvent('user_updated', {
            userId: editingUserId,
            nomeGuerra: userEditWarName.value.trim(),
            matricula: userEditRegistration.value.trim(),
            tipo: userEditType.value
        });
        closeUserEditModal();
        loadUsersList(true);
    } catch (error) {
        console.error(error);
        userEditError.textContent = 'Erro ao salvar usuário.';
    }
}

async function deleteUserProfile(user) {
    if (!db || !isAdmin || !user?.id) return;
    const shouldDelete = window.confirm('Deseja realmente excluir este usuário?');
    if (!shouldDelete) return;
    try {
        await db.collection('users').doc(user.id).delete();
        await logEvent('user_deleted', {
            userId: user.id,
            email: user.email || '',
            nomeGuerra: user.nomeGuerra || '',
            matricula: user.matricula || ''
        });
        loadUsersList(true);
    } catch (error) {
        console.error(error);
        alert('Erro ao excluir usuário.');
    }
}

async function loadWarNamesList(force = false) {
    if (!db || !auth || !auth.currentUser || isLoadingWarNames) return;
    if (warNamesCache.length > 0 && !force) {
        applyWarNamesFilter();
        updateWarNamesDatalist();
        return;
    }
    isLoadingWarNames = true;
    try {
        const snapshot = await db.collection('warNames').orderBy('nomeGuerra').get();
        warNamesCache = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        applyWarNamesFilter();
        updateWarNamesDatalist();
    } catch (error) {
        console.error(error);
    } finally {
        isLoadingWarNames = false;
    }
}

function applyWarNamesFilter() {
    const textValue = warNamesFilterText ? warNamesFilterText.value.trim().toLowerCase() : '';
    const filtered = warNamesCache.filter(entry => {
        const blob = `${entry.classe || ''} ${entry.nomeGuerra || ''} ${entry.matricula || ''} ${entry.status || ''}`.toLowerCase();
        return textValue ? blob.includes(textValue) : true;
    });
    renderWarNamesTable(filtered);
}

function renderWarNamesTable(entries) {
    if (!warNamesTableBody) return;
    warNamesTableBody.innerHTML = '';
    const totalPages = Math.max(1, Math.ceil(entries.length / warNamesPageSize));
    if (warNamesCurrentPage > totalPages) warNamesCurrentPage = totalPages;
    const startIndex = (warNamesCurrentPage - 1) * warNamesPageSize;
    const pageItems = entries.slice(startIndex, startIndex + warNamesPageSize);

    if (!pageItems.length) {
        warNamesTableBody.innerHTML = '<tr><td colspan="5">Nenhum agente encontrado.</td></tr>';
        if (warNamesPagination) warNamesPagination.innerHTML = '';
        return;
    }

    pageItems.forEach(entry => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${entry.classe || '-'}</td>
            <td>${entry.nomeGuerra || '-'}</td>
            <td>${entry.matricula || '-'}</td>
            <td>${entry.status || 'ativo'}</td>
            <td>
                <button class="btn-small" data-action="edit-warname" data-id="${entry.id}">Editar</button>
                <button class="btn-small btn-danger" data-action="toggle-warname" data-id="${entry.id}">${entry.status === 'inativo' ? 'Ativar' : 'Inativar'}</button>
                <button class="btn-small btn-danger" data-action="delete-warname" data-id="${entry.id}">Excluir</button>
            </td>
        `;
        warNamesTableBody.appendChild(row);
    });

    warNamesTableBody.querySelectorAll('button[data-action="edit-warname"]').forEach(button => {
        button.addEventListener('click', () => {
            const id = button.getAttribute('data-id');
            const entry = warNamesCache.find(item => item.id === id);
            if (entry) openWarNameModal(entry);
        });
    });

    warNamesTableBody.querySelectorAll('button[data-action="toggle-warname"]').forEach(button => {
        button.addEventListener('click', () => {
            const id = button.getAttribute('data-id');
            toggleWarNameStatus(id);
        });
    });

    warNamesTableBody.querySelectorAll('button[data-action="delete-warname"]').forEach(button => {
        button.addEventListener('click', () => {
            const id = button.getAttribute('data-id');
            deleteWarNameEntry(id);
        });
    });

    renderWarNamesPagination(totalPages);
}

function renderWarNamesPagination(totalPages) {
    if (!warNamesPagination) return;
    warNamesPagination.innerHTML = '';
    if (totalPages <= 1) return;

    const prevBtn = document.createElement('button');
    prevBtn.className = 'pagination-btn';
    prevBtn.textContent = 'Anterior';
    prevBtn.disabled = warNamesCurrentPage === 1;
    prevBtn.addEventListener('click', () => {
        warNamesCurrentPage -= 1;
        applyWarNamesFilter();
    });
    warNamesPagination.appendChild(prevBtn);

    for (let i = 1; i <= totalPages; i += 1) {
        const pageBtn = document.createElement('button');
        pageBtn.className = 'pagination-btn';
        if (i === warNamesCurrentPage) pageBtn.classList.add('active');
        pageBtn.textContent = String(i);
        pageBtn.addEventListener('click', () => {
            warNamesCurrentPage = i;
            applyWarNamesFilter();
        });
        warNamesPagination.appendChild(pageBtn);
    }

    const nextBtn = document.createElement('button');
    nextBtn.className = 'pagination-btn';
    nextBtn.textContent = 'Próximo';
    nextBtn.disabled = warNamesCurrentPage === totalPages;
    nextBtn.addEventListener('click', () => {
        warNamesCurrentPage += 1;
        applyWarNamesFilter();
    });
    warNamesPagination.appendChild(nextBtn);
}

function updateWarNamesDatalist() {
    if (!warNamesList) return;
    warNamesList.innerHTML = '';
    warNamesCache.filter(item => item.status !== 'inativo').forEach(item => {
        const option = document.createElement('option');
        option.value = item.nomeGuerra || '';
        warNamesList.appendChild(option);
    });
    updateUserCreateWarNameSelect(userCreateWarName?.value || '');
}

async function loadLogsList(force = false) {
    if (!db || !auth || !auth.currentUser || !isAdmin || isLoadingLogs) return;
    if (logsCache.length > 0 && !force) {
        applyLogsFilter();
        return;
    }
    isLoadingLogs = true;
    try {
        const snapshot = await db.collection('logs').orderBy('createdAt', 'desc').limit(1000).get();
        logsCache = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        applyLogsFilter();
    } catch (error) {
        console.error(error);
    } finally {
        isLoadingLogs = false;
    }
}

function applyLogsFilter() {
    const textValue = logsFilterText ? logsFilterText.value.trim().toLowerCase() : '';
    const filtered = logsCache.filter(entry => {
        const blob = `${entry.action || ''} ${entry.userEmail || ''} ${entry.nomeGuerra || ''} ${entry.matricula || ''} ${entry.metadata?.registroNumero || ''} ${entry.metadata?.ropId || ''}`.toLowerCase();
        return textValue ? blob.includes(textValue) : true;
    });
    renderLogsTable(filtered);
}

function renderLogsTable(entries) {
    if (!logsTableBody) return;
    logsTableBody.innerHTML = '';
    const totalPages = Math.max(1, Math.ceil(entries.length / logsPageSize));
    if (logsCurrentPage > totalPages) logsCurrentPage = totalPages;
    const startIndex = (logsCurrentPage - 1) * logsPageSize;
    const pageItems = entries.slice(startIndex, startIndex + logsPageSize);

    if (!pageItems.length) {
        logsTableBody.innerHTML = '<tr><td colspan="5">Nenhum log encontrado.</td></tr>';
        if (logsPagination) logsPagination.innerHTML = '';
        return;
    }

    pageItems.forEach(entry => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${formatDateTime(entry.createdAt)}</td>
            <td>${entry.action === 'login' ? 'Login' : entry.action === 'rop_saved' ? 'Registro de ROP' : entry.action}</td>
            <td>${entry.nomeGuerra || entry.userEmail || '-'}</td>
            <td>${entry.matricula || '-'}</td>
            <td>${entry.metadata?.registroNumero || entry.metadata?.ropId || '-'}</td>
        `;
        logsTableBody.appendChild(row);
    });

    renderLogsPagination(totalPages);
}

function renderLogsPagination(totalPages) {
    if (!logsPagination) return;
    logsPagination.innerHTML = '';
    if (totalPages <= 1) return;

    const prevBtn = document.createElement('button');
    prevBtn.className = 'pagination-btn';
    prevBtn.textContent = 'Anterior';
    prevBtn.disabled = logsCurrentPage === 1;
    prevBtn.addEventListener('click', () => {
        logsCurrentPage -= 1;
        applyLogsFilter();
    });
    logsPagination.appendChild(prevBtn);

    for (let i = 1; i <= totalPages; i += 1) {
        const pageBtn = document.createElement('button');
        pageBtn.className = 'pagination-btn';
        if (i === logsCurrentPage) pageBtn.classList.add('active');
        pageBtn.textContent = String(i);
        pageBtn.addEventListener('click', () => {
            logsCurrentPage = i;
            applyLogsFilter();
        });
        logsPagination.appendChild(pageBtn);
    }

    const nextBtn = document.createElement('button');
    nextBtn.className = 'pagination-btn';
    nextBtn.textContent = 'Próximo';
    nextBtn.disabled = logsCurrentPage === totalPages;
    nextBtn.addEventListener('click', () => {
        logsCurrentPage += 1;
        applyLogsFilter();
    });
    logsPagination.appendChild(nextBtn);
}

function openWarNameModal(entry = null) {
    if (!warNameModal) return;
    warNameError.textContent = '';
    editingWarNameId = entry?.id || null;
    warNameClass.value = entry?.classe || '';
    warNameValue.value = entry?.nomeGuerra || '';
    warNameRegistration.value = entry?.matricula || '';
    warNameStatus.value = entry?.status || 'ativo';
    warNameModal.style.display = 'flex';
}

function closeWarNameModal() {
    if (warNameModal) warNameModal.style.display = 'none';
    editingWarNameId = null;
}

async function saveWarNameEntry() {
    if (!db || !isAdmin) return;
    warNameError.textContent = '';
    if (!warNameClass.value || !warNameValue.value || !warNameRegistration.value) {
        warNameError.textContent = 'Preencha todos os campos obrigatórios.';
        return;
    }
    try {
        if (editingWarNameId) {
            await db.collection('warNames').doc(editingWarNameId).set({
                classe: warNameClass.value,
                nomeGuerra: warNameValue.value.trim(),
                matricula: warNameRegistration.value.trim(),
                status: warNameStatus.value
            }, { merge: true });
            await logEvent('agent_updated', {
                warNameId: editingWarNameId,
                nomeGuerra: warNameValue.value.trim(),
                matricula: warNameRegistration.value.trim()
            });
        } else {
            await db.collection('warNames').add({
                classe: warNameClass.value,
                nomeGuerra: warNameValue.value.trim(),
                matricula: warNameRegistration.value.trim(),
                status: warNameStatus.value
            });
            await logEvent('agent_created', {
                nomeGuerra: warNameValue.value.trim(),
                matricula: warNameRegistration.value.trim()
            });
        }
        closeWarNameModal();
        loadWarNamesList(true);
    } catch (error) {
        console.error(error);
        warNameError.textContent = 'Erro ao salvar nome de guerra.';
    }
}

async function toggleWarNameStatus(id) {
    if (!db || !isAdmin) return;
    const entry = warNamesCache.find(item => item.id === id);
    if (!entry) return;
    const nextStatus = entry.status === 'inativo' ? 'ativo' : 'inativo';
    try {
        await db.collection('warNames').doc(id).set({ status: nextStatus }, { merge: true });
        await logEvent('agent_updated', {
            warNameId: id,
            nomeGuerra: entry.nomeGuerra || '',
            matricula: entry.matricula || '',
            status: nextStatus
        });
        loadWarNamesList(true);
    } catch (error) {
        console.error(error);
    }
}

async function deleteWarNameEntry(id) {
    if (!db || !isAdmin) return;
    const entry = warNamesCache.find(item => item.id === id);
    if (!entry) return;
    const shouldDelete = window.confirm('Deseja realmente excluir este agente?');
    if (!shouldDelete) return;
    try {
        await db.collection('warNames').doc(id).delete();
        await logEvent('agent_deleted', {
            warNameId: id,
            nomeGuerra: entry.nomeGuerra || '',
            matricula: entry.matricula || ''
        });
        loadWarNamesList(true);
    } catch (error) {
        console.error(error);
        alert('Erro ao excluir agente.');
    }
}

async function sendPasswordResetForAgent(entry) {
    if (!db || !auth || !isAdmin) return;
    const matricula = (entry?.matricula || '').trim();
    if (!matricula) {
        alert('Matrícula não encontrada para resetar senha.');
        return;
    }

    try {
        const snapshot = await db.collection('users').where('matricula', '==', matricula).limit(1).get();
        if (snapshot.empty) {
            alert('Usuário não encontrado para esta matrícula.');
            return;
        }
        const user = snapshot.docs[0].data();
        const email = user?.email || '';
        if (!email) {
            alert('E-mail do usuário não encontrado.');
            return;
        }

        await auth.sendPasswordResetEmail(email);
        await logEvent('user_password_reset', {
            email,
            matricula,
            nomeGuerra: entry?.nomeGuerra || ''
        });
        alert('E-mail de reset enviado com sucesso.');
    } catch (error) {
        console.error(error);
        alert('Não foi possível enviar o e-mail de reset.');
    }
}

function findWarNameByValue(value) {
    return warNamesCache.find(item => item.nomeGuerra?.toLowerCase() === value.toLowerCase() && item.status !== 'inativo');
}

function autofillWarNameFields(nameInput, classInput, registrationInput) {
    const match = findWarNameByValue(nameInput.value);
    if (match) {
        if (classInput) classInput.value = match.classe || '';
        if (registrationInput) registrationInput.value = match.matricula || '';
    } else {
        if (classInput) classInput.value = '';
        if (registrationInput) registrationInput.value = '';
    }
}

function autofillProfileByWarName(value) {
    const match = findWarNameByValue(value);
    if (match) {
        profileClass.value = match.classe || profileClass.value;
        profileRegistration.value = match.matricula || profileRegistration.value;
    }
}

async function exportToPDF(ropData) {
    const doc = new jsPDF('p', 'mm', 'a4');
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 15; // Margens esquerda/direita: 1,5 cm
    const contentWidth = pageWidth - (margin * 2);
    const contentPadding = 6;
    const contentX = margin + contentPadding;
    const contentMaxWidth = contentWidth - (contentPadding * 2);
    let y = margin;

    const sectionHeaderHeight = 8;
    const sectionSpacing = 12;
    const footerHeight = 8;

    // Desenha linhas verticais até um limite (por padrão até o rodapé, mas pode ser customizado)
    let pageBorderLimit = null;
    const drawPageBorders = (limitY) => {
        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.3);
        // Se não houver limite, desenha só o topo (sem linhas verticais)
        if (typeof limitY !== 'number' || limitY <= margin) return;
        doc.line(margin, margin, margin, limitY);
        doc.line(margin + contentWidth, margin, margin + contentWidth, limitY);
    };

    const addPageWithBorders = () => {
        doc.addPage();
        // Só desenha linhas verticais se pageBorderLimit for válido e maior que margin
        if (typeof pageBorderLimit === 'number' && pageBorderLimit > margin) {
            drawPageBorders(pageBorderLimit);
        }
    };

    const drawSectionHeader = (title) => {
        doc.setFillColor(17, 17, 78);
        doc.rect(margin, y, contentWidth, sectionHeaderHeight, 'F');
        doc.setTextColor(255, 255, 255);
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(11);
        doc.text(title, pageWidth / 2, y + 5, { align: 'center' });
        const headerTop = y;
        y += sectionSpacing;
        return { headerTop };
    };

    const drawSectionLines = (startY, endY) => {
        // Se endY não for passado, desenha só o box do header
        if (typeof endY !== 'number' || endY <= startY) {
            endY = startY + sectionHeaderHeight;
        }
        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.3);
        doc.line(margin, startY, margin, endY);
        doc.line(margin + contentWidth, startY, margin + contentWidth, endY);
    };

    const drawJustifiedLine = (line, x, y, maxWidth, isLastLine) => {
        if (isLastLine) {
            doc.text(line, x, y);
            return;
        }
        const words = line.trim().split(/\s+/).filter(Boolean);
        if (words.length <= 1) {
            doc.text(line, x, y);
            return;
        }
        const wordsWidth = words.reduce((sum, word) => sum + doc.getTextWidth(word), 0);
        const spaces = words.length - 1;
        const extraSpace = Math.max(0, maxWidth - wordsWidth);
        const spaceWidth = extraSpace / spaces;
        let currentX = x;
        words.forEach((word, index) => {
            doc.text(word, currentX, y);
            currentX += doc.getTextWidth(word) + (index < spaces ? spaceWidth : 0);
        });
    };

    const renderJustifiedText = (text, startY, title, lineHeight = 5) => {
        const lines = doc.splitTextToSize(text, contentMaxWidth);
        y = startY;
        let localY = startY;
        let { headerTop } = drawSectionHeader(title);
        localY = y;

        doc.setFont('helvetica', 'normal');
        doc.setFontSize(10);
        doc.setTextColor(0, 0, 0);

        for (let i = 0; i < lines.length; i++) {
            if (localY + lineHeight > pageHeight - margin) {
                drawSectionLines(headerTop, localY);
                addPageWithBorders();
                localY = margin;
                y = localY;
                ({ headerTop } = drawSectionHeader(title));
                localY = y;
                // Garante cor preta após quebra de página
                doc.setFont('helvetica', 'normal');
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
            }
            const isLastLine = i === lines.length - 1;
            drawJustifiedLine(lines[i], contentX, localY, contentMaxWidth, isLastLine);
            localY += lineHeight;
        }

        drawSectionLines(headerTop, localY);
        return localY;
    };

    const drawField = (x, yPos, label, value, maxWidth, lineHeight = 4.5) => {
        const safeValue = value || '';
        doc.setFont('helvetica', 'normal');
        doc.text(label, x, yPos);
        const labelWidth = doc.getTextWidth(label) + 2;
        doc.setFont('helvetica', 'bold');
        const valueX = x + labelWidth;
        const valueLines = doc.splitTextToSize(safeValue, Math.max(1, maxWidth - labelWidth));
        valueLines.forEach((line, idx) => {
            doc.text(line, valueX, yPos + (idx * lineHeight));
        });
        return Math.max(1, valueLines.length) * lineHeight;
    };

    drawPageBorders();

    // Título principal (idêntico ao antigo)
    doc.setFillColor(17, 17, 78); // #11114E - azul escuro
    doc.rect(margin, y, contentWidth, 12, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.text('RELATÓRIO DE OCORRÊNCIA POLICIAL', pageWidth / 2, y + 8, { align: 'center' });
    y += 15;

    // Cabeçalho com logo - manter proporção da imagem
    const logoX = contentX;
    const logoY = y;
    
    // Carrega logo mantendo proporção
    const logoImage = await loadImage('logo.jpg');
    if (logoImage) {
        const maxLogoWidth = 18;
        const maxLogoHeight = 18;
        const scale = Math.min(maxLogoWidth / logoImage.width, maxLogoHeight / logoImage.height);
        const logoWidth = logoImage.width * scale;
        const logoHeight = logoImage.height * scale;
        doc.addImage(logoImage, 'JPEG', logoX, logoY, logoWidth, logoHeight);
    }
    
    // Logo e textos alinhados à esquerda, textos centralizados verticalmente à logo
    const logoMaxHeight = 18;
    const logoMaxWidth = 18;
    let logoHeight = logoMaxHeight;
    let logoWidth = logoMaxWidth;
    if (logoImage) {
        const scale = Math.min(logoMaxWidth / logoImage.width, logoMaxHeight / logoImage.height);
        logoWidth = logoImage.width * scale;
        logoHeight = logoImage.height * scale;
        doc.addImage(logoImage, 'JPEG', contentX, logoY, logoWidth, logoHeight);
    }

    const headerLines = [
        'SECRETARIA MUNICIPAL DA SEGURANÇA E CIDADANIA',
        'PREFEITURA DE ARACAJU',
        'POLÍCIA MUNICIPAL DE ARACAJU'
    ];
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.2);
    doc.setTextColor(0, 0, 0);
    // Centralizar verticalmente o bloco de texto em relação à logo
    const headerBlockHeight = headerLines.length * 4.5;
    const headerStartY = logoY + (logoHeight / 2) - (headerBlockHeight / 2) + 2;
    const textX = contentX + logoWidth + 3;
    headerLines.forEach((line, idx) => {
        doc.text(line, textX, headerStartY + idx * 4.5);
    });
    
    // Campos à direita, respeitando a margem e com quebra de linha
    const rightBoxWidth = 60;
    const registroX = pageWidth - margin - rightBoxWidth;
    const rightTextMaxX = pageWidth - margin;
    doc.setFontSize(8.5);
    doc.setFont('helvetica', 'normal');
    let rightY = logoY + 5;
    const drawRightField = (label, value) => {
        const labelWidth = doc.getTextWidth(label);
        const valueMaxWidth = rightTextMaxX - (registroX + labelWidth + 2);
        let valueLines = doc.splitTextToSize(value, valueMaxWidth > 20 ? valueMaxWidth : 20);
        doc.text(label, registroX, rightY);
        doc.setFont('helvetica', 'bold');
        valueLines.forEach((line, idx) => {
            doc.text(line, registroX + labelWidth + 2, rightY + idx * 4.5);
        });
        doc.setFont('helvetica', 'normal');
        rightY += Math.max(7, valueLines.length * 4.5);
    };
    drawRightField('Registro de ocorrência nº:', ropData.registroNumero || '');
    drawRightField('Data do Fato:', formatDate(ropData.dataFato));
    drawRightField('Hora do Fato:', ropData.horaFato || '');

    // Destinatário abaixo à esquerda, com espaçamento da logo
    let destinatarioY = logoY + logoHeight + 7;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.5);
    doc.setTextColor(0, 0, 0);
    doc.text('Destinatário:', contentX, destinatarioY);
    doc.setFont('helvetica', 'bold');
    const destinatarioMaxWidth = contentMaxWidth - 2;
    const destinatarioLines = doc.splitTextToSize(ropData.destinatario || '', destinatarioMaxWidth);
    destinatarioLines.forEach((line, idx) => {
        doc.text(line, contentX + 18, destinatarioY + idx * 4.5);
    });
    destinatarioY += destinatarioLines.length * 4.5;
    
    // Espaço reduzido após destinatário para a próxima seção
    y = destinatarioY + 4;

    // Checkboxes (prisão/flagrante etc) - fundo cinza
    const flagsText = buildFlagsText(ropData.flags || {});
    if (flagsText) {
        doc.setFillColor(180, 180, 180);
        doc.rect(margin, y, contentWidth, 8, 'F');
        doc.setFontSize(9);
        doc.setTextColor(0, 0, 0);
        doc.text(flagsText, pageWidth / 2, y + 5, { align: 'center' });
        y += 10;
    }

    // I - DADOS DA OCORRÊNCIA
    let sectionHeader = drawSectionHeader('I - DADOS DA OCORRÊNCIA');
    let sectionTopY = sectionHeader.headerTop;
    y += 2;

    // Tabela de dados da ocorrência com margens ajustadas
    const dados = [
        { label: 'Natureza:', value: formatNatureza(ropData.natureza) },
        { label: 'Tipo de delito:', value: ropData.tipoDelito || '' },
        { label: 'Local da Ocorrência:', value: ropData.localOcorrencia || '' },
        { label: 'UF:', value: ropData.uf || 'SE' },
        { label: 'Cidade:', value: ropData.cidade || 'Aracaju' },
        { label: 'Número:', value: ropData.numero || 'S/N' },
        { label: 'Bairro:', value: ropData.bairro || '' }
    ];

    // Função para desenhar campo com quebra de linha automática
    const drawFieldAuto = (label, value, x, yPos, maxWidth, labelBold = false, valueBold = true, lineHeight = 4.5) => {
        doc.setFont('helvetica', labelBold ? 'bold' : 'normal');
        doc.setFontSize(9);
        doc.setTextColor(0, 0, 0);
        doc.text(label, x, yPos);
        doc.setFont('helvetica', valueBold ? 'bold' : 'normal');
        const labelWidth = doc.getTextWidth(label) + 2;
        const valueLines = doc.splitTextToSize(value || '', maxWidth - labelWidth);
        valueLines.forEach((line, idx) => {
            doc.text(line, x + labelWidth, yPos + idx * lineHeight);
        });
        return valueLines.length * lineHeight;
    };

    // Primeira linha: Natureza e Tipo de delito
    let maxFieldWidth = 80;
    let usedHeight = drawFieldAuto('Natureza:', dados[0].value, contentX, y, maxFieldWidth);
    let usedHeight2 = drawFieldAuto('Tipo de delito:', dados[1].value, contentX + 100, y, 60);
    y += Math.max(usedHeight, usedHeight2) + 2;

    // Segunda linha: Local e UF
    usedHeight = drawFieldAuto('Local da Ocorrência:', dados[2].value, contentX, y, 90);
    usedHeight2 = drawFieldAuto('UF:', dados[3].value, contentX + 120, y, 20);
    y += Math.max(usedHeight, usedHeight2) + 2;

    // Terceira linha: Cidade, Número, Bairro
    usedHeight = drawFieldAuto('Cidade:', dados[4].value, contentX, y, 40);
    usedHeight2 = drawFieldAuto('Número:', dados[5].value, contentX + 60, y, 30);
    let usedHeight3 = drawFieldAuto('Bairro:', dados[6].value, contentX + 105, y, 40);
    y += Math.max(usedHeight, usedHeight2, usedHeight3) + 6;

    drawSectionLines(sectionTopY, y);

    // II - DADOS DAS PARTES ENVOLVIDAS NA OCORRÊNCIA
    sectionHeader = drawSectionHeader('II - DADOS DAS PARTES ENVOLVIDAS NA OCORRÊNCIA');
    sectionTopY = sectionHeader.headerTop;

    if (ropData.partes && ropData.partes.length > 0) {
        ropData.partes.forEach((parte, index) => {
            if (y > pageHeight - 60) {
                drawSectionLines(sectionTopY, y);
                addPageWithBorders();
                y = margin;
                sectionHeader = drawSectionHeader('II - DADOS DAS PARTES ENVOLVIDAS NA OCORRÊNCIA');
                sectionTopY = sectionHeader.headerTop;
            }
            doc.setFillColor(180, 180, 180);
            doc.rect(margin, y, contentWidth, 6, 'F');
            doc.setFontSize(9);
            doc.setTextColor(0, 0, 0);
            const roleText = `PARTE ENVOLVIDA: ${parte.role || ''}`;
            doc.text(roleText, margin + 5, y + 4);
            y += 10;
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(9);
            const colGap = 6;
            const colWidth = (contentMaxWidth - colGap) / 2;
            const leftX = contentX;
            const rightX = contentX + colWidth + colGap;
            let leftHeight = drawField(leftX, y, 'Nome:', parte.nome, colWidth);
            let rightHeight = drawField(rightX, y, 'Data de Nasc.:', parte.dataNascimento, colWidth);
            y += Math.max(leftHeight, rightHeight) + 2;
            leftHeight = drawField(leftX, y, 'Pai:', parte.pai, colWidth);
            rightHeight = drawField(rightX, y, 'Mãe:', parte.mae, colWidth);
            y += Math.max(leftHeight, rightHeight) + 2;
            leftHeight = drawField(leftX, y, 'Condição Física:', formatCondicaoFisica(parte.condicaoFisica), colWidth);
            rightHeight = drawField(rightX, y, 'Sexo:', formatSexo(parte.sexo), colWidth);
            y += Math.max(leftHeight, rightHeight) + 2;
            leftHeight = drawField(leftX, y, 'Naturalidade:', parte.naturalidade, colWidth);
            rightHeight = drawField(rightX, y, 'Endereço:', parte.endereco, colWidth);
            y += Math.max(leftHeight, rightHeight) + 2;
            leftHeight = drawField(leftX, y, 'CPF:', formatCPF(parte.cpf), colWidth);
            rightHeight = drawField(rightX, y, 'Cidade/UF:', parte.cidadeUf, colWidth);
            y += Math.max(leftHeight, rightHeight) + 2;
            leftHeight = drawField(leftX, y, 'Telefone:', formatTelefone(parte.telefone), contentMaxWidth);
            y += leftHeight + 6;
        });
    } else {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(10);
        doc.text('Sem envolvidos', contentX, y);
        y += 15;
    }

    drawSectionLines(sectionTopY, y);

    // III - RELATO DA OCORRÊNCIA
    if (y > pageHeight - 50) {
        addPageWithBorders();
        y = margin + 5;
    }

    // Relato com altura dinâmica e justificado (última linha à esquerda)
    const relatoText = ropData.relato || 'Não informado';
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    doc.setTextColor(0, 0, 0);
    const relatoStartY = y;
    y = renderJustifiedText(relatoText, y, 'III - RELATO DA OCORRÊNCIA', 5) + 5;
    drawSectionLines(relatoStartY - sectionSpacing, y - 5);

    // IV - APREENSÕES
    if (y > pageHeight - 50) {
        addPageWithBorders();
        y = margin + 5;
    }

    sectionHeader = drawSectionHeader('IV - APREENSÕES (OBJETOS, ARMAS, VEÍCULOS, SUBSTÂNCIA ENTORPECENTE, OUTROS...)');
    sectionTopY = sectionHeader.headerTop;

    if (ropData.apreensoes && ropData.apreensoes.length > 0) {
        // Tabela de apreensões com altura dinâmica
        const tableWidth = contentWidth;
        const col1Width = tableWidth * 0.3;
        const col2Width = tableWidth * 0.7;
        
        // Cabeçalho da tabela
        doc.setFillColor(180, 180, 180);
        doc.rect(margin, y, col1Width, 8, 'F');
        doc.rect(margin + col1Width, y, col2Width, 8, 'F');
        
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(10);
        doc.setTextColor(0, 0, 0);
        doc.text('Tipo', margin + (col1Width / 2), y + 5, { align: 'center' });
        doc.text('Especificação', margin + col1Width + (col2Width / 2), y + 5, { align: 'center' });
        
        y += 8;
        
        // Linhas da tabela com quebra de linha
        ropData.apreensoes.forEach((item, index) => {
            if (y > pageHeight - 50) {
                drawSectionLines(sectionTopY, y);
                addPageWithBorders();
                y = margin;
                sectionHeader = drawSectionHeader('IV - APREENSÕES (OBJETOS, ARMAS, VEÍCULOS, SUBSTÂNCIA ENTORPECENTE, OUTROS...)');
                sectionTopY = sectionHeader.headerTop;

                // Redesenha cabeçalho da tabela na nova página
                doc.setFillColor(180, 180, 180);
                doc.rect(margin, y, col1Width, 8, 'F');
                doc.rect(margin + col1Width, y, col2Width, 8, 'F');
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.text('Tipo', margin + (col1Width / 2), y + 5, { align: 'center' });
                doc.text('Especificação', margin + col1Width + (col2Width / 2), y + 5, { align: 'center' });
                y += 8;
            }
            
            const tipoText = item.tipo || '';
            const especificacaoText = item.especificacao || '';
            
            // Calcula altura necessária baseada no texto mais longo
            const tipoLines = doc.splitTextToSize(tipoText, col1Width - 5);
            const especificacaoLines = doc.splitTextToSize(especificacaoText, col2Width - 10);
            
            const linesCount = Math.max(tipoLines.length, especificacaoLines.length);
            const rowHeight = 8 + (Math.max(0, linesCount - 1) * 5);
            
            // Desenha bordas
            doc.setDrawColor(0, 0, 0);
            doc.setLineWidth(0.2);
            doc.rect(margin, y, col1Width, rowHeight);
            doc.rect(margin + col1Width, y, col2Width, rowHeight);
            
            // Texto do tipo (centralizado)
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(9);
            let tipoY = y + 5;
            tipoLines.forEach((line, i) => {
                doc.text(line, margin + (col1Width / 2), tipoY, { align: 'center' });
                tipoY += 5;
            });
            
            // Texto da especificação (alinhado à esquerda)
            let especY = y + 5;
            especificacaoLines.forEach((line, i) => {
                doc.text(line, margin + col1Width + 5, especY);
                especY += 5;
            });
            
            y += rowHeight;
        });
    } else {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(10);
        doc.text('Sem objetos apreendidos', contentX, y);
        y += 15;
    }

    y += 5;

    drawSectionLines(sectionTopY, y);

    // V - IMAGENS DO OCORRIDO
    if (y > pageHeight - 50) {
        addPageWithBorders();
        y = margin + 5;
    }

    sectionHeader = drawSectionHeader('V - IMAGENS DO OCORRIDO');
    sectionTopY = sectionHeader.headerTop;

    const anexosList = collectAnexosUrls(ropData);

    if (anexosList.length > 0) {
        const gap = 6;
        const colWidth = (contentMaxWidth - gap) / 2;
        const rowHeight = 55;
        let col = 0;
        let cursorX = contentX;
        let cursorY = y;

        for (let i = 0; i < anexosList.length; i += 1) {
            if (cursorY + rowHeight > pageHeight - margin) {
                drawSectionLines(sectionTopY, cursorY);
                addPageWithBorders();
                cursorY = margin;
                sectionHeader = drawSectionHeader('V - IMAGENS DO OCORRIDO');
                sectionTopY = sectionHeader.headerTop;
                col = 0;
                cursorX = contentX;
            }

            const url = anexosList[i];
            const imgData = await loadImageForPdf(url);
            if (imgData?.dataUrl) {
                const scale = Math.min(colWidth / imgData.width, rowHeight / imgData.height, 1);
                const drawW = imgData.width * scale;
                const drawH = imgData.height * scale;
                const offsetX = cursorX + (colWidth - drawW) / 2;
                const offsetY = cursorY + (rowHeight - drawH) / 2;
                doc.addImage(imgData.dataUrl, imgData.format, offsetX, offsetY, drawW, drawH);
            } else {
                doc.setFont('helvetica', 'normal');
                doc.setFontSize(9);
                doc.text('Imagem indisponível', cursorX + 5, cursorY + 20);
            }

            doc.setDrawColor(0, 0, 0);
            doc.setLineWidth(0.2);
            doc.rect(cursorX, cursorY, colWidth, rowHeight);

            col += 1;
            if (col >= 2) {
                col = 0;
                cursorX = contentX;
                cursorY += rowHeight + gap;
            } else {
                cursorX += colWidth + gap;
            }
        }

        y = cursorY + (col === 0 ? 2 : rowHeight + 2);
    } else {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(10);
        doc.text('Sem imagens anexadas', contentX, y);
        y += 15;
    }

    drawSectionLines(sectionTopY, y);

    // VI - AGENTES ENVOLVIDOS NA OCORRÊNCIA
    if (y > pageHeight - 50) {
        addPageWithBorders();
        y = margin + 5;
    }

    sectionHeader = drawSectionHeader('VI - AGENTES ENVOLVIDOS NA OCORRÊNCIA');
    sectionTopY = sectionHeader.headerTop;

    if (ropData.agentes && ropData.agentes.length > 0) {
        // Tabela de agentes
        const tableWidth = contentWidth;
        const col1Width = tableWidth * 0.25;
        const col2Width = tableWidth * 0.45;
        const col3Width = tableWidth * 0.3;
        
        // Cabeçalho da tabela
        doc.setFillColor(180, 180, 180);
        doc.rect(margin, y, col1Width, 8, 'F');
        doc.rect(margin + col1Width, y, col2Width, 8, 'F');
        doc.rect(margin + col1Width + col2Width, y, col3Width, 8, 'F');
        
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(10);
        doc.setTextColor(0, 0, 0);
        doc.text('Nível', margin + (col1Width / 2), y + 5, { align: 'center' });
        doc.text('Nome de Guerra', margin + col1Width + (col2Width / 2), y + 5, { align: 'center' });
        doc.text('Matrícula', margin + col1Width + col2Width + (col3Width / 2), y + 5, { align: 'center' });
        
        y += 8;
        
        // Linhas da tabela
        ropData.agentes.forEach((agent, index) => {
            if (y > pageHeight - 50) {
                drawSectionLines(sectionTopY, y);
                addPageWithBorders();
                y = margin;
                sectionHeader = drawSectionHeader('VI - AGENTES ENVOLVIDOS NA OCORRÊNCIA');
                sectionTopY = sectionHeader.headerTop;

                // Redesenha cabeçalho da tabela na nova página
                doc.setFillColor(180, 180, 180);
                doc.rect(margin, y, col1Width, 8, 'F');
                doc.rect(margin + col1Width, y, col2Width, 8, 'F');
                doc.rect(margin + col1Width + col2Width, y, col3Width, 8, 'F');
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.text('Nível', margin + (col1Width / 2), y + 5, { align: 'center' });
                doc.text('Nome de Guerra', margin + col1Width + (col2Width / 2), y + 5, { align: 'center' });
                doc.text('Matrícula', margin + col1Width + col2Width + (col3Width / 2), y + 5, { align: 'center' });
                y += 8;
            }
            
            // Desenha bordas
            doc.setDrawColor(0, 0, 0);
            doc.setLineWidth(0.2);
            doc.rect(margin, y, col1Width, 8);
            doc.rect(margin + col1Width, y, col2Width, 8);
            doc.rect(margin + col1Width + col2Width, y, col3Width, 8);
            
            // Texto com quebra de linha se necessário
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(9);
            
            // Nível (centralizado)
            const nivelText = agent.classe || '';
            const nivelLines = doc.splitTextToSize(nivelText, col1Width - 5);
            let nivelY = y + 5;
            nivelLines.forEach((line, i) => {
                doc.text(line, margin + (col1Width / 2), nivelY, { align: 'center' });
                if (i < nivelLines.length - 1) nivelY += 4;
            });
            
            // Nome de Guerra
            const nomeGuerraText = agent.nomeGuerra || '';
            const nomeGuerraLines = doc.splitTextToSize(nomeGuerraText, col2Width - 10);
            let nomeGuerraY = y + 5;
            nomeGuerraLines.forEach((line, i) => {
                doc.text(line, margin + col1Width + 5, nomeGuerraY);
                if (i < nomeGuerraLines.length - 1) nomeGuerraY += 4;
            });
            
            // Matrícula (centralizada)
            const matriculaText = agent.matricula || '';
            const matriculaLines = doc.splitTextToSize(matriculaText, col3Width - 5);
            let matriculaY = y + 5;
            matriculaLines.forEach((line, i) => {
                doc.text(line, margin + col1Width + col2Width + (col3Width / 2), matriculaY, { align: 'center' });
                if (i < matriculaLines.length - 1) matriculaY += 4;
            });
            
            // Aumenta altura da linha se houver múltiplas linhas
            const maxLines = Math.max(nivelLines.length, nomeGuerraLines.length, matriculaLines.length);
            y += 8 + (Math.max(0, maxLines - 1) * 4);
        });
    } else {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(10);
        doc.text('Não houve registro de agentes envolvidos.', contentX, y);
        y += 15;
    }

    drawSectionLines(sectionTopY, y);

    y += 10; // Espaçamento antes do posto de serviço

    // Posto de serviço e plantão - CORRIGIDO: não sobrepor
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    doc.text('POSTO DE SERVIÇO:', contentX, y);
    
    const postoServicoText = ropData.postoServico || '';
    const postoServicoWidth = doc.getTextWidth(postoServicoText);
    if (postoServicoWidth > 60) {
        const postoServicoLines = doc.splitTextToSize(postoServicoText, 60);
        doc.setFont('helvetica', 'bold');
        doc.text(postoServicoLines[0], contentX + 40, y);
        if (postoServicoLines.length > 1) {
            y += 5;
            doc.text(postoServicoLines[1], contentX + 40, y);
        }
    } else {
        doc.setFont('helvetica', 'bold');
        doc.text(postoServicoText, contentX + 40, y);
    }
    
    y += 6; // Espaçamento entre linha do posto e turno
    
    doc.setFont('helvetica', 'normal');
    doc.text('PLANTÃO:', contentX, y);
    doc.setFont('helvetica', 'bold');
    doc.text(ropData.turno || '', contentX + 25, y);
    
    y += 25; // Espaçamento maior antes da assinatura

    // Assinatura da equipe - CENTRALIZADA
    if (y > pageHeight - 50) {
        addPageWithBorders();
        y = margin + 5;
    }
    
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    const assinaturaStartX = pageWidth / 2 - 50; // Centralizada
    doc.line(assinaturaStartX, y, assinaturaStartX + 100, y);
    doc.text('Assinatura do responsável pela equipe', pageWidth / 2, y + 5, { align: 'center' });
    
    y += 35;

    // VII - RECIBO DA AUTORIDADE
    if (y > pageHeight - 80) {
        addPageWithBorders();
        y = margin + 5;
    }

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    doc.setTextColor(0, 0, 0);

    const reciboText = `Recebi este ROP com as devidas qualificações das pessoas envolvidas, com as condições físicas descritas acima e portando os objetos especificados no campo IV que foram apreendidos e apresentados nesta unidade policial.`;

    y = renderJustifiedText(reciboText, y, 'VII - RECIBO DA AUTORIDADE A QUE SE DESTINA OU SEU REPRESENTANTE LEGAL', 5) + 15;
    
    // Assinatura da autoridade - CENTRALIZADA
    const authAssinaturaStartX = pageWidth / 2 - 50;
    doc.line(authAssinaturaStartX, y, authAssinaturaStartX + 100, y);
    doc.text('Assinatura da Autoridade Policial', pageWidth / 2, y + 5, { align: 'center' });
    
    y += 35;

    // Rodapé - CORRIGIDO: tamanho menor para caber
    // Rodapé logo após a última informação
    y += 5;
    // Se o rodapé ultrapassar a margem inferior, mova para nova página
    if (y + footerHeight > pageHeight - margin) {
        addPageWithBorders();
        y = margin;
    }
    // Linhas verticais só até o início do rodapé (não sobrepor o box azul)
    drawPageBorders(y);
    // Após desenhar as linhas, zera o limite para não desenhar em novas páginas
    pageBorderLimit = null;

    // Rodapé informativo (registro/emissão)
    const registradoPor = ropData.createdBy?.nomeGuerra || 'Não informado';
    const registradoEm = formatDateTime(ropData.createdAt);
    const atualizadoPor = ropData.updatedBy?.nomeGuerra || 'Não informado';
    const atualizadoEm = formatDateTime(ropData.updatedAt);
    const emitidoPor = currentUserProfile?.nomeGuerra || 'Não informado';
    const emitidoEm = formatDateTime(new Date());
    const footerInfoLine1 = `Registrado por: ${registradoPor} | Criado em: ${registradoEm} | Atualizado por: ${atualizadoPor} | Atualizado em: ${atualizadoEm}`;
    const footerInfoLine2 = `Emitido por: ${emitidoPor} | Data de emissão: ${emitidoEm}`;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(70, 70, 70);
    doc.text('Gestão de Relatórios de Ocorrência (ROP)', pageWidth / 2, y - 8, { align: 'center' });
    doc.setFontSize(6);
    doc.text(footerInfoLine1, pageWidth / 2, y - 5, { align: 'center' });
    doc.text(footerInfoLine2, pageWidth / 2, y - 2, { align: 'center' });

    // Rodapé
    doc.setFillColor(17, 17, 78);
    doc.rect(margin, y, contentWidth, footerHeight, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    const footerText = 'Polícia Municipal de Aracaju | Endereço: Av. Ivo do Prado, 904, Bairro São José | CEP: 49.015-070 | Tel.: (79) 98166-7790';
    doc.text(footerText, pageWidth / 2, y + 5, { align: 'center' });

    // Impede que addPageWithBorders redesenhe linhas após o rodapé
    pageBorderLimit = y;

    // Numeração das páginas
    const totalPages = doc.internal.getNumberOfPages();
    for (let i = 1; i <= totalPages; i += 1) {
        doc.setPage(i);
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(7);
        doc.setTextColor(90, 90, 90);
        doc.text(`Página ${i} de ${totalPages}`, pageWidth - margin, pageHeight - 4, { align: 'right' });
    }

    await logEvent('rop_pdf_generated', {
        ropId: ropData.id || '',
        registroNumero: ropData.registroNumero || ''
    });

    // Salva o PDF
    const cleanRegistro = (ropData.registroNumero || 'sem_registro').replace(/[^a-zA-Z0-9]/g, '_');
    doc.save(`ROP_${cleanRegistro}.pdf`);
}

// Funções auxiliares
function buildFlagsText(flags = {}) {
    const parts = [];
    if (flags.prisaoFlagrante) parts.push('Prisão/Apreensão em Flagrante');
    if (flags.comunicacaoOcorrencia) parts.push('Comunicação de Ocorrência');
    if (flags.mandadoPrisao) parts.push('Mandado de prisão em aberto');
    if (flags.outros) parts.push('Outros');
    return parts.length > 0 ? parts.join(' | ') : '';
}

function formatCondicaoFisica(condicao) {
    const map = {
        'sem_ferimentos': 'Sem Ferimentos',
        'especificado_abaixo': 'Especificado abaixo'
    };
    return map[condicao] || condicao || 'Sem Ferimentos';
}

function formatSexo(sexo) {
    const map = {
        'masculino': 'Masculino',
        'feminino': 'Feminino',
        'outros': 'Outros'
    };
    return map[sexo] || sexo || 'Masculino';
}

function formatCPF(cpf) {
    if (!cpf || cpf.length !== 11) return cpf || '-';
    return `${cpf.substring(0, 3)}.${cpf.substring(3, 6)}.${cpf.substring(6, 9)}-${cpf.substring(9, 11)}`;
}

function formatTelefone(telefone) {
    if (!telefone) return '-';
    const digits = telefone.replace(/\D/g, '');
    if (digits.length === 10) {
        return `(${digits.substring(0, 2)}) ${digits.substring(2, 6)}-${digits.substring(6, 10)}`;
    } else if (digits.length === 11) {
        return `(${digits.substring(0, 2)}) ${digits.substring(2, 7)}-${digits.substring(7, 11)}`;
    }
    return telefone;
}

function formatDate(dateValue) {
    if (!dateValue) return 'Não informado';
    if (dateValue.toDate) return dateValue.toDate().toLocaleDateString('pt-BR');
    const date = new Date(dateValue);
    if (Number.isNaN(date.getTime())) return 'Não informado';
    return date.toLocaleDateString('pt-BR');
}

function formatNatureza(value) {
    const map = {
        prisao_em_flagrante: 'Prisão em Flagrante',
        apreensao_em_flagrante: 'Apreensão em Flagrante',
        prisao_e_apreensao_em_flagrante: 'Prisão e Apreensão em Flagrante',
        mandado_de_prisao_em_aberto: 'Mandado de prisão em aberto',
        comunicado_de_fato_relevante: 'Comunicado de fato relevante',
        termo_circunstanciado_de_ocorrencia: 'Termo Circunstanciado de Ocorrência',
        outros: 'Outros'
    };
    return map[value] || value || '';
}

function collectAnexosUrls(ropData) {
    const list = Array.isArray(ropData?.anexosUrls) ? ropData.anexosUrls.slice() : [];
    if (!list.length && ropData?.anexos) {
        const parsed = String(ropData.anexos)
            .split(/\r?\n/)
            .map(line => line.trim())
            .filter(Boolean)
            .filter(line => /^https?:\/\//i.test(line));
        return parsed;
    }
    return list.filter(line => /^https?:\/\//i.test(line));
}

function getImageFormatFromUrl(url) {
    const lower = String(url || '').toLowerCase();
    if (lower.includes('.png')) return 'PNG';
    if (lower.includes('.webp')) return 'JPEG';
    if (lower.includes('.jpeg') || lower.includes('.jpg')) return 'JPEG';
    return 'JPEG';
}

function blobToDataUrl(blob) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result || '');
        reader.onerror = () => resolve('');
        reader.readAsDataURL(blob);
    });
}

function wait(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

async function resolveStorageUrl(url) {
    const cleaned = String(url || '').trim();
    if (!cleaned) return '';
    if (cleaned.startsWith('gs://') && storage?.refFromURL) {
        try {
            return await storage.refFromURL(cleaned).getDownloadURL();
        } catch (error) {
            return '';
        }
    }
    return cleaned;
}

async function fetchImageBlob(url, attempts = 2) {
    let lastError = null;
    for (let i = 0; i < attempts; i += 1) {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 15000);
        try {
            const response = await fetch(url, {
                mode: 'cors',
                cache: 'no-store',
                credentials: 'omit',
                signal: controller.signal
            });
            clearTimeout(timeoutId);
            if (!response.ok) throw new Error('Falha ao baixar imagem');
            return await response.blob();
        } catch (error) {
            clearTimeout(timeoutId);
            lastError = error;
            await wait(500);
        }
    }
    throw lastError || new Error('Falha ao baixar imagem');
}

async function loadImageForPdf(url) {
    try {
        const resolvedUrl = await resolveStorageUrl(url);
        if (!resolvedUrl) return null;
        const blob = await fetchImageBlob(resolvedUrl, 2);
        const type = (blob.type || '').toLowerCase();
        const dataUrl = await blobToDataUrl(blob);
        const img = await loadImage(dataUrl);
        if (!img) return null;

        if (type.includes('png')) {
            return { dataUrl, format: 'PNG', width: img.width, height: img.height };
        }
        if (type.includes('jpeg') || type.includes('jpg')) {
            return { dataUrl, format: 'JPEG', width: img.width, height: img.height };
        }

        const canvas = document.createElement('canvas');
        canvas.width = img.width;
        canvas.height = img.height;
        const ctx = canvas.getContext('2d');
        if (!ctx) return null;
        ctx.drawImage(img, 0, 0);
        const jpegUrl = canvas.toDataURL('image/jpeg', 0.92);
        return { dataUrl: jpegUrl, format: 'JPEG', width: img.width, height: img.height };
    } catch (error) {
        try {
            const resolvedUrl = await resolveStorageUrl(url);
            if (!resolvedUrl) return null;
            const img = await loadImage(resolvedUrl, true);
            if (!img) return null;
            const canvas = document.createElement('canvas');
            canvas.width = img.width;
            canvas.height = img.height;
            const ctx = canvas.getContext('2d');
            if (!ctx) return null;
            ctx.drawImage(img, 0, 0);
            const dataUrl = canvas.toDataURL('image/jpeg', 0.92);
            return { dataUrl, format: 'JPEG', width: img.width, height: img.height };
        } catch (err) {
            return null;
        }
    }
}

function loadImage(src, useCors = false) {
    return new Promise((resolve) => {
        const img = new Image();
        if (useCors) {
            img.crossOrigin = 'anonymous';
        }
        img.onload = () => resolve(img);
        img.onerror = () => resolve(null);
        img.src = src;
    });
}
