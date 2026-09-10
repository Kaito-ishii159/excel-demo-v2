"use strict";

const AUTH = Object.freeze({
  id: "care-admin-5821",
  saltBase64: "hQmjgD1detlyqAI8s3utfQ==",
  derivedKeyBase64: "4iiQz3DcJkH57dPERaJ3bDFE61Cr7tkFfGT0BIXcSG8=",
  iterations: 210000,
  draftSaltBase64: "orAu57dPhXexdKCYOGbqaA==",
  draftIterations: 210000,
});

const STORAGE = Object.freeze({
  draft: "postpartumCareReportDraft.v1",
  authenticated: "postpartumCareAuthenticated.v1",
  authenticatedUntil: "postpartumCareAuthenticatedUntil.v2",
  draftKey: "postpartumCareDraftKey.v2",
  failedAttempts: "postpartumCareFailedAttempts.v1",
  lockedUntil: "postpartumCareLockedUntil.v1",
});

const AUTH_SESSION_DURATION_MS = 24 * 60 * 60 * 1000;
const AUTH_DATABASE = Object.freeze({
  name: "postpartumCareAuth.v1",
  store: "sessions",
  key: "current",
  version: 1,
});

const WORKBOOK_PATHS = Object.freeze({
  sheet: "xl/worksheets/sheet1.xml",
  styles: "xl/styles.xml",
  sharedStrings: "xl/sharedStrings.xml",
  vml: "xl/drawings/vmlDrawing1.vml",
});

const ASSESSMENTS = [
  { key: "bodyRecovery", group: "母の身体状況", label: "身体の回復", good: 5, observe: 6 },
  { key: "breastCondition", group: "母の身体状況", label: "乳房の状態", good: 9, observe: 10 },
  { key: "woundCondition", group: "母の身体状況", label: "創部の状態", good: 11, observe: 12 },
  { key: "meal", group: "母の生活状況", label: "食事", good: 13, observe: 14 },
  { key: "elimination", group: "母の生活状況", label: "排泄", good: 15, observe: 16 },
  { key: "sleepRest", group: "母の生活状況", label: "睡眠・休息", good: 17, observe: 18 },
  { key: "mood", group: "母の心理状態", label: "気分・感情の状態", good: 19, observe: 20 },
  { key: "birthReflection", group: "母の心理状態", label: "お産の振り返り", good: 21, observe: 22 },
  { key: "stressResponse", group: "母の心理状態", label: "ストレスへの対応", good: 23, observe: 24 },
  { key: "infantGeneral", group: "児の健康", label: "全身状態", good: 25, observe: 26 },
  { key: "feeding", group: "児の健康", label: "哺乳状態", good: 27, observe: 28 },
  { key: "growth", group: "児の健康", label: "発育状態", good: 29, observe: 30 },
  { key: "development", group: "児の健康", label: "発達状態", good: 31, observe: 32 },
  { key: "childCareSkills", group: "育児状況", label: "育児手技", good: 33, observe: 34 },
  { key: "breastfeedingStatus", group: "育児状況", label: "授乳状況", good: 35, observe: 36 },
  { key: "partnerRelationship", group: "育児環境", label: "パートナーとの関係", good: 37, observe: 38 },
  { key: "familyRelationship", group: "育児環境", label: "家族関係", good: 39, observe: 40 },
  { key: "supportSystem", group: "育児環境", label: "サポート体制", good: 41, observe: 42 },
];

const GUIDANCE_GROUPS = [
  {
    title: "産婦のケア",
    items: [
      ["motherHealth", "母親の健康管理", "B32"],
      ["breastCare", "乳房ケア", "B33"],
      ["psychologicalCare", "心理面のケア", "B34"],
      ["birthReview", "バースレビュー", "B35"],
      ["familyPlanning", "家族計画", "B36"],
      ["gdmInfo", "GDM情報提供", "B37"],
    ],
  },
  {
    title: "児のケア",
    items: [
      ["growthDevelopment", "発育・発達の確認", "F32"],
      ["eliminationObservation", "排泄の観察", "F33"],
      ["skinObservation", "皮膚の観察", "F34"],
      ["environmentAdjustment", "環境調整", "F35"],
    ],
  },
  {
    title: "母子のケア",
    items: [
      ["breastfeedingGuidance", "授乳指導", "J32"],
      ["bathingGuidance", "沐浴指導", "J33"],
      ["childrearingGuidance", "育児相談・指導", "J34"],
      ["serviceInfo", "母子保健サービスの情報提供", "J35"],
      ["familySupport", "家族への支援・情報提供", "J36"],
      ["otherGuidance", "その他必要とする保健指導等", "J37"],
    ],
  },
];

const CONTROL_SHAPES = Object.freeze({
  1: 28673, 2: 28674, 3: 28678, 4: 28679,
  5: 28682, 6: 28683, 7: 28719, 8: 28720,
  9: 28745, 10: 28746, 11: 28747, 12: 28748,
  13: 28749, 14: 28750, 15: 28751, 16: 28752,
  17: 28753, 18: 28754, 19: 28755, 20: 28756,
  21: 28757, 22: 28758, 23: 28759, 24: 28760,
  25: 28761, 26: 28762, 27: 28763, 28: 28764,
  29: 28765, 30: 28766, 31: 28767, 32: 28768,
  33: 28769, 34: 28770, 35: 28771, 36: 28772,
  37: 28773, 38: 28774, 39: 28775, 40: 28776,
  41: 28777, 42: 28778, 43: 28779, 44: 28780,
  45: 28781, 46: 28782,
});

const NS = Object.freeze({
  spreadsheet: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  excel: "urn:schemas-microsoft-com:office:excel",
  xml: "http://www.w3.org/XML/1998/namespace",
});

const DRAFT_FORMAT_VERSION = 2;
const DRAFT_AAD = new TextEncoder().encode("postpartum-care-report-draft-v2");

const loginView = document.getElementById("loginView");
const appView = document.getElementById("appView");
const loginForm = document.getElementById("loginForm");
const loginId = document.getElementById("loginId");
const loginPassword = document.getElementById("loginPassword");
const loginButton = document.getElementById("loginButton");
const loginMessage = document.getElementById("loginMessage");
const togglePassword = document.getElementById("togglePassword");
const reportForm = document.getElementById("reportForm");
const appMessage = document.getElementById("appMessage");
const downloadButton = document.getElementById("downloadButton");
const clearButton = document.getElementById("clearButton");
const logoutButton = document.getElementById("logoutButton");
const multipleBirthField = document.getElementById("multipleBirth");
const multipleAdditionalCountField = document.getElementById("multipleAdditionalCount");
const excelImportButton = document.getElementById("excelImportButton");
const excelImportInput = document.getElementById("excelImportInput");
const excelImportMessage = document.getElementById("excelImportMessage");

let saveTimer = null;
let authExpiryTimer = null;
let hasUnsavedInput = false;
let draftCryptoKey = null;
let draftSaveRequest = 0;

buildAssessmentFields();
buildGuidanceFields();
setDefaultValues();
void initializeAuthentication();

loginForm.addEventListener("submit", handleLogin);
togglePassword.addEventListener("click", () => {
  const showing = loginPassword.type === "text";
  loginPassword.type = showing ? "password" : "text";
  togglePassword.textContent = showing ? "表示" : "非表示";
});

reportForm.addEventListener("input", scheduleDraftSave);
reportForm.addEventListener("change", scheduleDraftSave);
reportForm.addEventListener("submit", handleDownload);
clearButton.addEventListener("click", clearAllInputs);
logoutButton.addEventListener("click", () => void logout());
excelImportButton.addEventListener("click", () => {
  excelImportInput.value = "";
  excelImportInput.click();
});
excelImportInput.addEventListener("change", handleExcelImport);

multipleBirthField.addEventListener("change", syncMultipleBirthFields);
multipleAdditionalCountField.addEventListener("input", () => {
  if (Number(multipleAdditionalCountField.value) > 0) {
    multipleBirthField.value = "yes";
  }
  syncMultipleBirthFields();
});

window.addEventListener("beforeunload", (event) => {
  if (!hasUnsavedInput) return;
  event.preventDefault();
  event.returnValue = "";
});

function buildAssessmentFields() {
  const container = document.getElementById("assessmentFields");
  container.innerHTML = ASSESSMENTS.map((item) => `
    <div class="assessment-row">
      <div class="assessment-label">
        <span class="assessment-group">${escapeHtml(item.group)}</span>
        ${escapeHtml(item.label)}
      </div>
      <select name="assessment_${item.key}" aria-label="${escapeHtml(item.label)}" required>
        <option value="">選択してください</option>
        <option value="good" selected>良</option>
        <option value="observe">要観察</option>
      </select>
    </div>
  `).join("");
}

function buildGuidanceFields() {
  const container = document.getElementById("guidanceFields");
  container.innerHTML = GUIDANCE_GROUPS.map((group) => `
    <div class="guidance-column">
      <h3>${escapeHtml(group.title)}</h3>
      ${group.items.map(([key, label]) => `
        <label class="check-row">
          <input type="checkbox" name="guidance_${key}" value="1" checked>
          <span>${escapeHtml(label)}</span>
        </label>
      `).join("")}
    </div>
  `).join("");
}

function setDefaultValues() {
  const now = new Date();
  document.getElementById("reportMonth").value = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  syncMultipleBirthFields();
}

async function initializeAuthentication() {
  try {
    const persistentStatus = await restorePersistentAuthentication();
    if (persistentStatus === "restored") {
      await showApp();
      return;
    }
    if (persistentStatus === "expired") {
      showMessage(loginMessage, "ログインから24時間が経過しました。再度ログインしてください。", "error");
    }
  } catch (error) {
    console.warn("Persistent authentication could not be restored.", error);
    localStorage.removeItem(STORAGE.authenticatedUntil);
  }

  if (sessionStorage.getItem(STORAGE.authenticated) === "true") {
    const keyRestored = await restoreDraftKeyFromSession();
    if (keyRestored) {
      await showApp();
      return;
    }
    sessionStorage.removeItem(STORAGE.authenticated);
    sessionStorage.removeItem(STORAGE.draftKey);
  }

  loginView.hidden = false;
  appView.hidden = true;
  loginId.focus();
}

async function handleLogin(event) {
  event.preventDefault();
  clearMessage(loginMessage);

  const lockedUntil = Number(sessionStorage.getItem(STORAGE.lockedUntil) || 0);
  if (Date.now() < lockedUntil) {
    const seconds = Math.ceil((lockedUntil - Date.now()) / 1000);
    showMessage(loginMessage, `入力回数が多いため、${seconds}秒後に再度お試しください。`, "error");
    return;
  }

  loginButton.disabled = true;
  loginButton.textContent = "確認中…";

  try {
    const idMatches = loginId.value.trim() === AUTH.id;
    const passwordMatches = await verifyPassword(loginPassword.value);

    if (idMatches && passwordMatches) {
      const keyBytes = await initializeDraftEncryption(loginPassword.value);
      const authenticatedUntil = Date.now() + AUTH_SESSION_DURATION_MS;

      try {
        await persistAuthentication(authenticatedUntil);
        sessionStorage.removeItem(STORAGE.draftKey);
      } catch (error) {
        console.warn("24-hour authentication persistence is unavailable.", error);
        sessionStorage.setItem(STORAGE.draftKey, bytesToBase64(keyBytes));
      }

      sessionStorage.setItem(STORAGE.authenticated, "true");
      sessionStorage.removeItem(STORAGE.failedAttempts);
      sessionStorage.removeItem(STORAGE.lockedUntil);
      loginPassword.value = "";
      await showApp();
      return;
    }

    const failures = Number(sessionStorage.getItem(STORAGE.failedAttempts) || 0) + 1;
    sessionStorage.setItem(STORAGE.failedAttempts, String(failures));
    if (failures >= 5) {
      sessionStorage.setItem(STORAGE.lockedUntil, String(Date.now() + 30000));
      sessionStorage.setItem(STORAGE.failedAttempts, "0");
      showMessage(loginMessage, "入力回数が多いため、30秒間ロックしました。", "error");
    } else {
      showMessage(loginMessage, "IDまたはパスワードが正しくありません。", "error");
    }
  } catch (error) {
    console.error(error);
    showMessage(loginMessage, "認証処理に失敗しました。ブラウザを更新してください。", "error");
  } finally {
    loginButton.disabled = false;
    loginButton.textContent = "ログイン";
  }
}

async function verifyPassword(password) {
  if (!window.crypto?.subtle || !password) return false;
  const keyMaterial = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(password),
    "PBKDF2",
    false,
    ["deriveBits"]
  );
  const derived = await crypto.subtle.deriveBits(
    {
      name: "PBKDF2",
      salt: base64ToBytes(AUTH.saltBase64),
      iterations: AUTH.iterations,
      hash: "SHA-256",
    },
    keyMaterial,
    256
  );
  return constantTimeEqual(new Uint8Array(derived), base64ToBytes(AUTH.derivedKeyBase64));
}

async function initializeDraftEncryption(password) {
  const keyBytes = await deriveDraftKeyBytes(password);
  draftCryptoKey = await importDraftCryptoKey(keyBytes);
  return keyBytes;
}

async function restoreDraftKeyFromSession() {
  const encodedKey = sessionStorage.getItem(STORAGE.draftKey);
  if (!encodedKey) return false;

  try {
    draftCryptoKey = await importDraftCryptoKey(base64ToBytes(encodedKey));
    return true;
  } catch (error) {
    console.error(error);
    draftCryptoKey = null;
    return false;
  }
}

async function deriveDraftKeyBytes(password) {
  const keyMaterial = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(password),
    "PBKDF2",
    false,
    ["deriveBits"]
  );
  const derived = await crypto.subtle.deriveBits(
    {
      name: "PBKDF2",
      salt: base64ToBytes(AUTH.draftSaltBase64),
      iterations: AUTH.draftIterations,
      hash: "SHA-256",
    },
    keyMaterial,
    256
  );
  return new Uint8Array(derived);
}

async function importDraftCryptoKey(keyBytes) {
  return crypto.subtle.importKey(
    "raw",
    keyBytes,
    { name: "AES-GCM" },
    false,
    ["encrypt", "decrypt"]
  );
}

function openAuthDatabase() {
  return new Promise((resolve, reject) => {
    if (!window.indexedDB) {
      reject(new Error("IndexedDB is not available"));
      return;
    }

    const request = indexedDB.open(AUTH_DATABASE.name, AUTH_DATABASE.version);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(AUTH_DATABASE.store)) {
        request.result.createObjectStore(AUTH_DATABASE.store);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Could not open authentication storage"));
    request.onblocked = () => reject(new Error("Authentication storage is blocked"));
  });
}

async function getPersistentAuthentication() {
  const database = await openAuthDatabase();
  try {
    return await new Promise((resolve, reject) => {
      const transaction = database.transaction(AUTH_DATABASE.store, "readonly");
      const request = transaction.objectStore(AUTH_DATABASE.store).get(AUTH_DATABASE.key);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error || new Error("Could not read authentication storage"));
    });
  } finally {
    database.close();
  }
}

async function putPersistentAuthentication(record) {
  const database = await openAuthDatabase();
  try {
    await new Promise((resolve, reject) => {
      const transaction = database.transaction(AUTH_DATABASE.store, "readwrite");
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error || new Error("Could not write authentication storage"));
      transaction.onabort = () => reject(transaction.error || new Error("Authentication storage write was aborted"));
      transaction.objectStore(AUTH_DATABASE.store).put(record, AUTH_DATABASE.key);
    });
  } finally {
    database.close();
  }
}

async function deletePersistentAuthentication() {
  const database = await openAuthDatabase();
  try {
    await new Promise((resolve, reject) => {
      const transaction = database.transaction(AUTH_DATABASE.store, "readwrite");
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error || new Error("Could not clear authentication storage"));
      transaction.onabort = () => reject(transaction.error || new Error("Authentication storage clear was aborted"));
      transaction.objectStore(AUTH_DATABASE.store).delete(AUTH_DATABASE.key);
    });
  } finally {
    database.close();
  }
}

async function canUseDraftKey(key) {
  if (!key) return false;
  try {
    const iv = crypto.getRandomValues(new Uint8Array(12));
    await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, new Uint8Array(0));
    return true;
  } catch (error) {
    return false;
  }
}

async function persistAuthentication(authenticatedUntil) {
  if (!draftCryptoKey) throw new Error("Draft encryption key is not available");

  await putPersistentAuthentication({ authenticatedUntil, draftKey: draftCryptoKey });
  const stored = await getPersistentAuthentication();
  if (
    !stored
    || Number(stored.authenticatedUntil) !== authenticatedUntil
    || !(await canUseDraftKey(stored.draftKey))
  ) {
    await deletePersistentAuthentication();
    throw new Error("Persistent authentication verification failed");
  }

  draftCryptoKey = stored.draftKey;
  localStorage.setItem(STORAGE.authenticatedUntil, String(authenticatedUntil));
  scheduleAuthenticationExpiry(authenticatedUntil);
}

async function restorePersistentAuthentication() {
  const authenticatedUntil = Number(localStorage.getItem(STORAGE.authenticatedUntil));
  if (!Number.isFinite(authenticatedUntil) || authenticatedUntil <= 0) return "none";

  const remaining = authenticatedUntil - Date.now();
  if (remaining <= 0) {
    await clearAuthenticationStorage();
    return "expired";
  }
  if (remaining > AUTH_SESSION_DURATION_MS) {
    await clearAuthenticationStorage();
    return "invalid";
  }

  const stored = await getPersistentAuthentication();
  if (
    !stored
    || Number(stored.authenticatedUntil) !== authenticatedUntil
    || !(await canUseDraftKey(stored.draftKey))
  ) {
    await clearAuthenticationStorage();
    return "invalid";
  }

  draftCryptoKey = stored.draftKey;
  sessionStorage.setItem(STORAGE.authenticated, "true");
  sessionStorage.removeItem(STORAGE.draftKey);
  scheduleAuthenticationExpiry(authenticatedUntil);
  return "restored";
}

function scheduleAuthenticationExpiry(authenticatedUntil) {
  clearTimeout(authExpiryTimer);
  const remaining = Math.max(0, authenticatedUntil - Date.now());
  authExpiryTimer = setTimeout(() => {
    void logout({ expired: true });
  }, remaining);
}

async function clearAuthenticationStorage() {
  localStorage.removeItem(STORAGE.authenticatedUntil);
  sessionStorage.removeItem(STORAGE.authenticated);
  sessionStorage.removeItem(STORAGE.draftKey);
  try {
    await deletePersistentAuthentication();
  } catch (error) {
    console.warn("Persistent authentication could not be cleared.", error);
  }
}

async function encryptDraft(draft) {
  if (!draftCryptoKey) throw new Error("Draft encryption key is not available");

  const iv = crypto.getRandomValues(new Uint8Array(12));
  const plaintext = new TextEncoder().encode(JSON.stringify(draft));
  const ciphertext = await crypto.subtle.encrypt(
    { name: "AES-GCM", iv, additionalData: DRAFT_AAD },
    draftCryptoKey,
    plaintext
  );

  return {
    version: DRAFT_FORMAT_VERSION,
    algorithm: "AES-GCM",
    ivBase64: bytesToBase64(iv),
    ciphertextBase64: bytesToBase64(new Uint8Array(ciphertext)),
  };
}

async function decryptDraft(payload) {
  if (!draftCryptoKey) throw new Error("Draft encryption key is not available");
  if (payload.version !== DRAFT_FORMAT_VERSION || payload.algorithm !== "AES-GCM") {
    throw new Error("Unsupported draft format");
  }

  const plaintext = await crypto.subtle.decrypt(
    {
      name: "AES-GCM",
      iv: base64ToBytes(payload.ivBase64),
      additionalData: DRAFT_AAD,
    },
    draftCryptoKey,
    base64ToBytes(payload.ciphertextBase64)
  );

  return JSON.parse(new TextDecoder().decode(plaintext));
}

async function writeEncryptedDraft(draft) {
  const encrypted = await encryptDraft(draft);
  localStorage.setItem(STORAGE.draft, JSON.stringify(encrypted));
}

async function showApp() {
  document.body.classList.add("authenticated");
  loginView.hidden = true;
  loginView.setAttribute("aria-hidden", "true");
  loginView.style.setProperty("display", "none", "important");
  appView.hidden = false;
  appView.removeAttribute("aria-hidden");
  appView.style.removeProperty("display");
  await restoreDraft();
  window.scrollTo({ top: 0, behavior: "auto" });
}

async function logout(options = {}) {
  const { expired = false } = options;
  clearTimeout(authExpiryTimer);
  authExpiryTimer = null;
  const clearStoragePromise = clearAuthenticationStorage();
  draftCryptoKey = null;
  document.body.classList.remove("authenticated");
  appView.hidden = true;
  appView.setAttribute("aria-hidden", "true");
  appView.style.setProperty("display", "none", "important");
  loginView.hidden = false;
  loginView.removeAttribute("aria-hidden");
  loginView.style.removeProperty("display");
  loginId.value = "";
  loginPassword.value = "";
  if (expired) {
    showMessage(loginMessage, "ログインから24時間が経過しました。再度ログインしてください。", "error");
  } else {
    clearMessage(loginMessage);
  }
  window.scrollTo({ top: 0, behavior: "auto" });
  loginId.focus();
  await clearStoragePromise;
}

function scheduleDraftSave() {
  hasUnsavedInput = true;
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    void saveDraft();
  }, 450);
}

async function saveDraft(options = {}) {
  const { showStatus = true } = options;
  const requestId = ++draftSaveRequest;
  const draft = {
    savedAt: new Date().toISOString(),
    data: collectFormData(),
  };

  try {
    const encrypted = await encryptDraft(draft);
    if (requestId !== draftSaveRequest) return false;

    localStorage.setItem(STORAGE.draft, JSON.stringify(encrypted));
    hasUnsavedInput = false;
    if (showStatus) {
      showMessage(appMessage, "入力内容を暗号化して、この端末に一時保存しました。", "success");
    }
    return true;
  } catch (error) {
    console.error(error);
    if (requestId === draftSaveRequest && showStatus) {
      showMessage(appMessage, "一時保存に失敗しました。再ログインしてお試しください。", "error");
    }
    return false;
  }
}

async function restoreDraft() {
  const raw = localStorage.getItem(STORAGE.draft);
  if (!raw) return;

  try {
    const stored = JSON.parse(raw);
    const isLegacyPlaintext = stored && stored.data && !stored.ciphertextBase64;
    const draft = isLegacyPlaintext ? stored : await decryptDraft(stored);

    restoreFormValues(draft.data || {});

    if (isLegacyPlaintext) {
      await writeEncryptedDraft(draft);
    }

    const savedAt = draft.savedAt ? new Date(draft.savedAt).toLocaleString("ja-JP") : "";
    showMessage(
      appMessage,
      `暗号化された一時保存データを復元しました。${savedAt ? `（${savedAt}）` : ""}`,
      "success"
    );
    hasUnsavedInput = false;
  } catch (error) {
    console.error(error);
    showMessage(
      appMessage,
      "一時保存データを復元できませんでした。必要に応じて「入力内容を全消去」で削除してください。",
      "error"
    );
  }
}

function restoreFormValues(data) {
  Object.entries(data).forEach(([name, value]) => {
    const elements = reportForm.elements.namedItem(name);
    if (!elements) return;

    if (elements instanceof RadioNodeList) {
      elements.value = String(value ?? "");
    } else if (elements.type === "checkbox") {
      elements.checked = Boolean(value);
    } else {
      elements.value = value ?? "";
    }
  });
  syncMultipleBirthFields();
}

function collectFormData() {
  const data = {};
  const formData = new FormData(reportForm);

  for (const element of reportForm.elements) {
    if (!element.name || element.disabled || element.type === "submit" || element.type === "button") continue;
    if (element.type === "checkbox") {
      data[element.name] = element.checked;
    } else if (element.type !== "radio") {
      data[element.name] = formData.get(element.name) ?? "";
    }
  }
  return data;
}

async function handleExcelImport(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  clearMessage(excelImportMessage);
  excelImportButton.disabled = true;
  excelImportButton.textContent = "Excelを読み込み中…";

  try {
    if (!/\.xlsx$/i.test(file.name || "")) {
      throw new Error("Unsupported file type");
    }

    const workbook = await readWorkbook(file);
    const data = await readFormDataFromWorkbook(workbook);
    restoreFormValues(data);
    hasUnsavedInput = true;

    const saved = await saveDraft({ showStatus: false });
    if (saved) {
      showMessage(excelImportMessage, "過去のExcelをフォームへ復元し、暗号化して一時保存しました。", "success");
    } else {
      showMessage(
        excelImportMessage,
        "フォームへの復元は完了しましたが、一時保存に失敗しました。内容を確認してからExcelを作成してください。",
        "error"
      );
    }
  } catch (error) {
    showMessage(
      excelImportMessage,
      "このExcelは読み込めませんでした。このアプリで作成した報告書（.xlsx）を選択してください。",
      "error"
    );
  } finally {
    excelImportInput.value = "";
    excelImportButton.disabled = false;
    excelImportButton.textContent = "過去のExcelを読み込む";
  }
}

async function readWorkbook(file) {
  if (typeof JSZip === "undefined") throw new Error("JSZip is not loaded");

  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const sheetFile = zip.file(WORKBOOK_PATHS.sheet);
  if (!sheetFile) throw new Error("Worksheet is missing");

  const sheetDoc = parseXml(await sheetFile.async("string"), WORKBOOK_PATHS.sheet);
  validateWorkbookStructure(sheetDoc);

  const sharedStringsFile = zip.file(WORKBOOK_PATHS.sharedStrings);
  const sharedStrings = sharedStringsFile
    ? readSharedStrings(parseXml(await sharedStringsFile.async("string"), WORKBOOK_PATHS.sharedStrings))
    : [];

  const vmlFile = zip.file(WORKBOOK_PATHS.vml);
  const vmlDoc = vmlFile ? parseXml(await vmlFile.async("string"), WORKBOOK_PATHS.vml) : null;
  return { zip, sheetDoc, sharedStrings, vmlDoc };
}

function validateWorkbookStructure(sheetDoc) {
  const worksheet = sheetDoc.documentElement;
  if (!worksheet || worksheet.localName !== "worksheet") throw new Error("Invalid worksheet");

  const requiredCells = ["B4", "J4", "K6", "B16", "B42"];
  const hasRequiredCells = requiredCells.every((reference) => findCell(sheetDoc, reference));
  const hasComplaintMerge = Array.from(sheetDoc.getElementsByTagNameNS(NS.spreadsheet, "mergeCell"))
    .some((merge) => merge.getAttribute("ref") === "B16:P16");
  if (!hasRequiredCells || !hasComplaintMerge) throw new Error("Unexpected worksheet structure");
}

function readSharedStrings(sharedStringsDoc) {
  return Array.from(sharedStringsDoc.getElementsByTagNameNS(NS.spreadsheet, "si"))
    .map((item) => readStringContainer(item));
}

function readStringContainer(container) {
  return Array.from(container.children).map((child) => {
    if (child.localName === "t") return child.textContent || "";
    if (child.localName !== "r") return "";
    return Array.from(child.getElementsByTagNameNS(NS.spreadsheet, "t"))
      .map((text) => text.textContent || "")
      .join("");
  }).join("");
}

function findCell(sheetDoc, cellReference) {
  return Array.from(sheetDoc.getElementsByTagNameNS(NS.spreadsheet, "c"))
    .find((cell) => cell.getAttribute("r") === cellReference) || null;
}

function readCellValue(sheetDoc, sharedStrings, cellReference) {
  const cell = findCell(sheetDoc, cellReference);
  if (!cell) return "";

  if (cell.getAttribute("t") === "inlineStr") {
    const inlineString = Array.from(cell.children).find((child) => child.localName === "is");
    return inlineString ? readStringContainer(inlineString) : "";
  }

  const value = Array.from(cell.children).find((child) => child.localName === "v")?.textContent || "";
  if (cell.getAttribute("t") === "s") {
    const index = Number(value);
    return Number.isInteger(index) && index >= 0 ? sharedStrings[index] || "" : "";
  }
  if (cell.getAttribute("t") === "b") return value === "1" ? "TRUE" : "FALSE";
  return value;
}

async function readFormDataFromWorkbook(workbook) {
  const { zip, sheetDoc, sharedStrings, vmlDoc } = workbook;
  const cell = (reference) => cleanImportedText(readCellValue(sheetDoc, sharedStrings, reference));
  const controlStates = await readControlStates(zip, vmlDoc);
  const child = parseChildIdentity(cell("K6"));
  const ages = parseBabyAges(cell("E9"));
  const supportCell = cell("D39");
  const continuedSupport = readExclusivePair(controlStates, 43, 44, "no", "yes")
    || readExclusivePair(controlStates, 7, 8, "no", "yes")
    || (/要/.test(supportCell) ? "yes" : /否/.test(supportCell) ? "no" : "");

  const data = {
    registrationNumber: ["B4", "C4", "D4", "E4", "F4", "G4"]
      .map(cell).join("").replace(/\s/g, "").slice(0, 6),
    reportMonth: parseReportMonth(cell("J4"), cell("O4")),
    motherKana: cell("C5"),
    childKana: cell("K5"),
    motherName: cell("C6"),
    motherAge: normalizeImportedNumber(cell("H6")),
    childName: child.name,
    childOrder: child.order,
    address: cell("B7"),
    pregnancyWeeks: normalizeImportedNumber(cell("F8")),
    birthDate: parseReiwaDate(cell("J8")),
    babyAgeMonths: ages.baby,
    correctedAgeMonths: ages.corrected,
    pregnancyNoteStatus: readExclusivePair(controlStates, 3, 4, "none", "present"),
    birthWeight: normalizeImportedNumber(cell("E12")),
    pregnancyCourse: cell("E11"),
    multipleBirth: readExclusivePair(controlStates, 1, 2, "no", "yes"),
    multipleAdditionalCount: normalizeImportedNumber(cell("H13")),
    visitStart: parseVisitDateTime(sheetDoc, sharedStrings, 14),
    visitEnd: parseVisitDateTime(sheetDoc, sharedStrings, 15),
    mainComplaint: cell("B16"),
    assessmentNotes: cell("H24"),
    resultNotes: cell("B38"),
    continuedSupport,
    healthDeptContact: readExclusivePair(controlStates, 45, 46, "yes", "no"),
    supportDetails: parseSupportDetails(supportCell),
    facilityName: cell("B42"),
    staffName: cell("L42"),
  };

  for (const item of ASSESSMENTS) {
    data[`assessment_${item.key}`] = readExclusivePair(
      controlStates,
      item.good,
      item.observe,
      "good",
      "observe"
    );
  }

  for (const group of GUIDANCE_GROUPS) {
    for (const [key, , reference] of group.items) {
      data[`guidance_${key}`] = /[○〇]/.test(cell(reference));
    }
  }

  return data;
}

async function readControlStates(zip, vmlDoc) {
  const states = new Map();
  await Promise.all(Array.from({ length: 46 }, async (_, index) => {
    const controlNumber = index + 1;
    states.set(controlNumber, await readCheckboxState(zip, vmlDoc, controlNumber));
  }));
  return states;
}

async function readCheckboxState(zip, vmlDoc, controlNumber) {
  const propPath = `xl/ctrlProps/ctrlProp${controlNumber}.xml`;
  const propFile = zip.file(propPath);
  if (propFile) {
    const propDoc = parseXml(await propFile.async("string"), propPath);
    const root = propDoc.documentElement;
    if (!root.hasAttribute("checked")) return false;
    const value = root.getAttribute("checked").toLowerCase();
    return !["0", "false", "unchecked"].includes(value);
  }

  if (!vmlDoc) return null;
  const shapeId = CONTROL_SHAPES[controlNumber];
  const shape = Array.from(vmlDoc.getElementsByTagNameNS("*", "shape"))
    .find((item) => item.getAttribute("id") === `_x0000_s${shapeId}`);
  if (!shape) return null;

  const checked = Array.from(shape.getElementsByTagNameNS("*", "Checked"))[0];
  if (!checked) return false;
  return !["0", "false"].includes(String(checked.textContent || "1").trim().toLowerCase());
}

function readExclusivePair(states, firstControl, secondControl, firstValue, secondValue) {
  const firstChecked = states.get(firstControl) === true;
  const secondChecked = states.get(secondControl) === true;
  if (firstChecked && !secondChecked) return firstValue;
  if (secondChecked && !firstChecked) return secondValue;
  return "";
}

function parseReportMonth(yearValue, monthValue) {
  const year = westernYearFromExcel(yearValue);
  const month = Number(normalizeImportedNumber(monthValue));
  if (!year || month < 1 || month > 12) return "";
  return `${year}-${String(month).padStart(2, "0")}`;
}

function parseReiwaDate(value) {
  const year = westernYearFromExcel(value);
  const normalized = normalizeDigits(value);
  const parts = normalized.match(/(?:年|[RＲ]\s*(?:元|\d+)\s*)\D*(\d{1,2})\D+(\d{1,2})\D*日?/i);
  if (!year || !parts) return "";

  const month = Number(parts[1]);
  const day = Number(parts[2]);
  return isValidDateParts(year, month, day)
    ? `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`
    : "";
}

function parseVisitDateTime(sheetDoc, sharedStrings, row) {
  const value = (column) => cleanImportedText(readCellValue(sheetDoc, sharedStrings, `${column}${row}`));
  const year = westernYearFromExcel(value("B"));
  const month = Number(normalizeImportedNumber(value("E")));
  const day = Number(normalizeImportedNumber(value("G")));
  const hour = Number(normalizeImportedNumber(value("I")));
  const minute = Number(normalizeImportedNumber(value("K")));

  if (!isValidDateParts(year, month, day) || hour < 0 || hour > 23 || minute < 0 || minute > 59) return "";
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}T${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;
}

function parseChildIdentity(value) {
  const match = String(value || "").match(/^([\s\S]*?)\s*【\s*第\s*([0-9０-９]*)\s*子\s*】/);
  if (!match) return { name: cleanImportedText(value), order: "" };
  return {
    name: cleanImportedText(match[1]),
    order: normalizeImportedNumber(match[2]),
  };
}

function parseBabyAges(value) {
  const normalized = normalizeDigits(value);
  const match = normalized.match(/生後\s*(\d*)\s*か月[\s\S]*?修正月齢\s*[：:]\s*(\d*)\s*か月/);
  return match
    ? { baby: match[1] || "", corrected: match[2] || "" }
    : { baby: "", corrected: "" };
}

function parseSupportDetails(value) {
  const match = String(value || "").match(/支援内容\s*[：:]\s*([\s\S]*?)\s*[〕）)]?\s*$/);
  return match ? match[1].trim() : "";
}

function westernYearFromExcel(value) {
  const normalized = normalizeDigits(value);
  const western = normalized.match(/\b(20\d{2})\b/);
  if (western) return Number(western[1]);

  const reiwa = normalized.match(/(?:令和|[RＲ])\s*(元|\d{1,2})/i);
  if (!reiwa) return null;
  const eraYear = reiwa[1] === "元" ? 1 : Number(reiwa[1]);
  return eraYear > 0 ? eraYear + 2018 : null;
}

function normalizeImportedNumber(value) {
  const match = normalizeDigits(value).replaceAll(",", "").match(/-?\d+(?:\.\d+)?/);
  return match ? match[0] : "";
}

function cleanImportedText(value) {
  return String(value || "").replace(/\r\n?/g, "\n").trim();
}

function normalizeDigits(value) {
  return String(value || "").replace(/[０-９]/g, (digit) => String(digit.charCodeAt(0) - 0xFEE0));
}

function isValidDateParts(year, month, day) {
  if (!year || month < 1 || month > 12 || day < 1 || day > 31) return false;
  const date = new Date(year, month - 1, day);
  return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day;
}

async function handleDownload(event) {
  event.preventDefault();
  clearMessage(appMessage);

  // 未入力項目があっても、入力済みの内容だけでExcelを作成できるようにする。
  const data = collectFormData();

  downloadButton.disabled = true;
  downloadButton.textContent = "Excelを作成中…";

  try {
    if (typeof JSZip === "undefined") throw new Error("JSZip is not loaded");

    const response = await fetch(`./template.xlsx?t=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) throw new Error(`Template fetch failed: ${response.status}`);

    const zip = await JSZip.loadAsync(await response.arrayBuffer());
    const sheetXml = await zip.file(WORKBOOK_PATHS.sheet).async("string");
    const vmlXml = await zip.file(WORKBOOK_PATHS.vml).async("string");

    const sheetDoc = parseXml(sheetXml, WORKBOOK_PATHS.sheet);
    const vmlDoc = parseXml(vmlXml, WORKBOOK_PATHS.vml);

    applyCellValues(sheetDoc, data);
    await applyMainComplaintBlackStyle(zip, sheetDoc);
    await applyCheckboxValues(zip, vmlDoc, data);

    zip.file(WORKBOOK_PATHS.sheet, serializeXml(sheetDoc));
    zip.file(WORKBOOK_PATHS.vml, serializeXml(vmlDoc));

    const blob = await zip.generateAsync({
      type: "blob",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const filename = createFilename(data);
    triggerDownload(blob, filename);

    if (document.getElementById("clearDraftAfterDownload").checked) {
      localStorage.removeItem(STORAGE.draft);
      hasUnsavedInput = false;
      showMessage(appMessage, "Excelを作成しました。一時保存データは削除済みです。", "success");
    } else {
      await saveDraft({ showStatus: false });
      showMessage(appMessage, "Excelを作成しました。暗号化された一時保存データは端末内に残しています。", "success");
    }
  } catch (error) {
    console.error(error);
    showMessage(appMessage, "Excelの作成に失敗しました。ページを更新して再度お試しください。", "error");
  } finally {
    downloadButton.disabled = false;
    downloadButton.textContent = "Excelを作成・ダウンロード";
  }
}

function applyCellValues(sheetDoc, data) {
  const report = monthParts(data.reportMonth);
  const reportReiwa = toReiwaYear(report.year);
  const birth = dateParts(data.birthDate);
  const start = dateTimeParts(data.visitStart);
  const end = dateTimeParts(data.visitEnd);

  const registration = String(data.registrationNumber || "").replace(/\D/g, "").slice(0, 6).padEnd(6, " ");
  ["B4", "C4", "D4", "E4", "F4", "G4"].forEach((cell, index) => setCellString(sheetDoc, cell, registration[index].trim()));
  setCellString(sheetDoc, "J4", report.year ? `R${reportReiwa}` : "");
  setCellString(sheetDoc, "O4", report.month || "");

  setCellString(sheetDoc, "C5", data.motherKana);
  setCellString(sheetDoc, "K5", data.childKana);
  setCellString(sheetDoc, "C6", data.motherName);
  setCellString(sheetDoc, "H6", data.motherAge, "18");

  const childOrder = data.childOrder ? String(data.childOrder) : "　";
  const childName = String(data.childName || "").trim();
  setCellString(sheetDoc, "K6", `${childName}${childName ? "　" : ""}【第${childOrder}子】`);
  setCellString(sheetDoc, "B7", data.address);

  setCellString(sheetDoc, "F8", data.pregnancyWeeks);
  setCellString(sheetDoc, "J8", birth.year ? `Ｒ${toReiwaYear(birth.year)}　年　${birth.month}　月　${birth.day}　日` : "");
  const corrected = data.correctedAgeMonths ? String(data.correctedAgeMonths) : "";
  setCellString(sheetDoc, "E9", `生後　${data.babyAgeMonths || ""}　か月　【修正月齢：　${corrected}　か月】`);
  setCellString(sheetDoc, "E11", data.pregnancyCourse);
  setCellString(sheetDoc, "E12", data.birthWeight ? `${data.birthWeight} g` : "");
  setCellNumber(sheetDoc, "H13", data.multipleAdditionalCount, "87");

  setVisitRow(sheetDoc, 14, start);
  setVisitRow(sheetDoc, 15, end);
  setCellString(sheetDoc, "B16", data.mainComplaint);
  setCellString(sheetDoc, "H24", data.assessmentNotes);

  for (const group of GUIDANCE_GROUPS) {
    for (const [key, , cell] of group.items) {
      setCellString(sheetDoc, cell, data[`guidance_${key}`] ? "○" : "");
    }
  }

  setCellString(sheetDoc, "B38", data.resultNotes);
  const supportDetails = String(data.supportDetails || "").trim();
  setCellString(sheetDoc, "D39", `要〔支援内容：${supportDetails}〕`);
  setCellString(sheetDoc, "B42", data.facilityName);
  setCellString(sheetDoc, "L42", data.staffName);
}

async function applyMainComplaintBlackStyle(zip, sheetDoc) {
  const stylesFile = zip.file(WORKBOOK_PATHS.styles);
  const complaintCell = findCell(sheetDoc, "B16");
  if (!stylesFile || !complaintCell) throw new Error("Workbook styles are missing");

  const stylesDoc = parseXml(await stylesFile.async("string"), WORKBOOK_PATHS.styles);
  const fonts = stylesDoc.getElementsByTagNameNS(NS.spreadsheet, "fonts")[0];
  const cellXfs = stylesDoc.getElementsByTagNameNS(NS.spreadsheet, "cellXfs")[0];
  if (!fonts || !cellXfs) throw new Error("Workbook styles are missing");
  const sourceStyleIndex = Number(complaintCell.getAttribute("s"));
  const sourceStyle = Array.from(cellXfs.children)[sourceStyleIndex];
  if (!sourceStyle) throw new Error("Complaint style is missing");

  const sourceFontIndex = Number(sourceStyle.getAttribute("fontId"));
  const sourceFont = Array.from(fonts.children)[sourceFontIndex];
  if (!sourceFont) throw new Error("Complaint font is missing");

  const blackFont = sourceFont.cloneNode(true);
  Array.from(blackFont.children)
    .filter((child) => child.localName === "color")
    .forEach((color) => color.remove());
  const blackColor = stylesDoc.createElementNS(NS.spreadsheet, "color");
  blackColor.setAttribute("rgb", "FF000000");
  const fontName = Array.from(blackFont.children).find((child) => child.localName === "name");
  blackFont.insertBefore(blackColor, fontName || null);
  fonts.appendChild(blackFont);
  const blackFontIndex = Array.from(fonts.children).length - 1;
  fonts.setAttribute("count", String(blackFontIndex + 1));

  const blackStyle = sourceStyle.cloneNode(true);
  blackStyle.setAttribute("fontId", String(blackFontIndex));
  blackStyle.setAttribute("applyFont", "1");
  cellXfs.appendChild(blackStyle);
  const blackStyleIndex = Array.from(cellXfs.children).length - 1;
  cellXfs.setAttribute("count", String(blackStyleIndex + 1));

  complaintCell.setAttribute("s", String(blackStyleIndex));
  zip.file(WORKBOOK_PATHS.styles, serializeXml(stylesDoc));
}

function setVisitRow(sheetDoc, row, parts) {
  setCellString(sheetDoc, `B${row}`, parts.year ? `Ｒ${toReiwaYear(parts.year)}` : "");
  setCellString(sheetDoc, `E${row}`, parts.month || "");
  setCellString(sheetDoc, `G${row}`, parts.day || "");
  setCellString(sheetDoc, `I${row}`, parts.hour || "");
  setCellString(sheetDoc, `K${row}`, parts.minute || "");
}

async function applyCheckboxValues(zip, vmlDoc, data) {
  const states = new Map();

  setPair(states, 3, 4, data.pregnancyNoteStatus, "none", "present");
  setPair(states, 1, 2, data.multipleBirth, "no", "yes");

  for (const item of ASSESSMENTS) {
    setPair(states, item.good, item.observe, data[`assessment_${item.key}`], "good", "observe");
  }

  setPair(states, 7, 8, data.continuedSupport, "no", "yes");
  setPair(states, 43, 44, data.continuedSupport, "no", "yes");
  setPair(states, 45, 46, data.healthDeptContact, "yes", "no");

  for (let controlNumber = 1; controlNumber <= 46; controlNumber += 1) {
    const checked = states.get(controlNumber) === true;
    const propPath = `xl/ctrlProps/ctrlProp${controlNumber}.xml`;
    const propFile = zip.file(propPath);
    if (propFile) {
      const propDoc = parseXml(await propFile.async("string"), propPath);
      const root = propDoc.documentElement;
      if (checked) root.setAttribute("checked", "Checked");
      else root.removeAttribute("checked");
      zip.file(propPath, serializeXml(propDoc));
    }

    const shapeId = CONTROL_SHAPES[controlNumber];
    if (shapeId) setVmlCheckbox(vmlDoc, shapeId, checked);
  }
}

function setPair(states, firstControl, secondControl, value, firstValue, secondValue) {
  states.set(firstControl, value === firstValue);
  states.set(secondControl, value === secondValue);
}

function setVmlCheckbox(vmlDoc, shapeId, checked) {
  const shapes = Array.from(vmlDoc.getElementsByTagNameNS("*", "shape"));
  const shape = shapes.find((item) => item.getAttribute("id") === `_x0000_s${shapeId}`);
  if (!shape) return;

  const clientData = Array.from(shape.getElementsByTagNameNS("*", "ClientData"))[0];
  if (!clientData) return;

  const checkedElements = Array.from(clientData.getElementsByTagNameNS("*", "Checked"));
  checkedElements.forEach((element) => element.remove());

  if (checked) {
    const checkedElement = vmlDoc.createElementNS(NS.excel, "x:Checked");
    checkedElement.textContent = "1";
    const noThreeD = Array.from(clientData.children).find((element) => element.localName === "NoThreeD");
    clientData.insertBefore(checkedElement, noThreeD || null);
  }
}

function syncMultipleBirthFields() {
  const isMultiple = multipleBirthField.value === "yes";
  multipleAdditionalCountField.required = isMultiple;
  if (multipleBirthField.value === "no") {
    multipleAdditionalCountField.value = "";
  }
}

function setCellNumber(doc, cellReference, value, fallbackStyle = null) {
  const cell = getOrCreateCell(doc, cellReference, fallbackStyle);
  if (fallbackStyle != null) cell.setAttribute("s", String(fallbackStyle));
  Array.from(cell.children).forEach((child) => {
    if (["v", "f", "is"].includes(child.localName)) child.remove();
  });
  cell.removeAttribute("t");
  const numeric = String(value ?? "").trim();
  if (!numeric) return;
  const numberValue = Number(numeric);
  if (!Number.isFinite(numberValue)) return;
  const v = doc.createElementNS(NS.spreadsheet, "v");
  v.textContent = String(numberValue);
  cell.appendChild(v);
}

function setCellString(doc, cellReference, value, fallbackStyle = null) {
  const cell = getOrCreateCell(doc, cellReference, fallbackStyle);
  Array.from(cell.children).forEach((child) => {
    if (["v", "f", "is"].includes(child.localName)) child.remove();
  });

  cell.setAttribute("t", "inlineStr");
  const inlineString = doc.createElementNS(NS.spreadsheet, "is");
  const text = doc.createElementNS(NS.spreadsheet, "t");
  text.setAttributeNS(NS.xml, "xml:space", "preserve");
  text.textContent = value == null ? "" : String(value);
  inlineString.appendChild(text);
  cell.appendChild(inlineString);
}

function getOrCreateCell(doc, cellReference, fallbackStyle) {
  const cells = Array.from(doc.getElementsByTagNameNS(NS.spreadsheet, "c"));
  const existing = cells.find((cell) => cell.getAttribute("r") === cellReference);
  if (existing) return existing;

  const rowNumber = Number(cellReference.match(/\d+/)?.[0]);
  const sheetData = doc.getElementsByTagNameNS(NS.spreadsheet, "sheetData")[0];
  let row = Array.from(sheetData.children).find((item) => item.localName === "row" && Number(item.getAttribute("r")) === rowNumber);

  if (!row) {
    row = doc.createElementNS(NS.spreadsheet, "row");
    row.setAttribute("r", String(rowNumber));
    const followingRow = Array.from(sheetData.children).find((item) => Number(item.getAttribute("r")) > rowNumber);
    sheetData.insertBefore(row, followingRow || null);
  }

  const cell = doc.createElementNS(NS.spreadsheet, "c");
  cell.setAttribute("r", cellReference);
  if (fallbackStyle != null) cell.setAttribute("s", String(fallbackStyle));

  const targetColumn = columnNumber(cellReference);
  const followingCell = Array.from(row.children).find((item) => item.localName === "c" && columnNumber(item.getAttribute("r")) > targetColumn);
  row.insertBefore(cell, followingCell || null);
  return cell;
}

function parseXml(xml, filename) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const error = doc.getElementsByTagName("parsererror")[0];
  if (error) throw new Error(`XML parse error: ${filename}`);
  return doc;
}

function serializeXml(doc) {
  const serialized = new XMLSerializer().serializeToString(doc);
  return serialized.startsWith("<?xml") ? serialized : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${serialized}`;
}

function clearAllInputs() {
  if (!confirm("入力内容とこの端末の一時保存データを削除しますか？")) return;
  reportForm.reset();
  localStorage.removeItem(STORAGE.draft);
  setDefaultValues();
  document.getElementById("clearDraftAfterDownload").checked = true;
  hasUnsavedInput = false;
  showMessage(appMessage, "入力内容と一時保存データを削除しました。", "success");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function createFilename(data) {
  const date = String(data.visitStart || "").slice(0, 10).replaceAll("-", "") || new Date().toISOString().slice(0, 10).replaceAll("-", "");
  const name = sanitizeFilename(data.motherName || "利用者");
  return `産後ケア実施報告書_${date}_${name}.xlsx`;
}

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

function monthParts(value) {
  const match = String(value || "").match(/^(\d{4})-(\d{2})$/);
  return match ? { year: Number(match[1]), month: Number(match[2]) } : { year: null, month: null };
}

function dateParts(value) {
  const match = String(value || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return match ? { year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) } : { year: null, month: null, day: null };
}

function dateTimeParts(value) {
  const match = String(value || "").match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  return match ? {
    year: Number(match[1]), month: Number(match[2]), day: Number(match[3]),
    hour: Number(match[4]), minute: String(match[5]).padStart(2, "0"),
  } : { year: null, month: null, day: null, hour: null, minute: null };
}

function toReiwaYear(westernYear) {
  return Number(westernYear) - 2018;
}

function columnNumber(cellReference) {
  const letters = String(cellReference).match(/^[A-Z]+/)?.[0] || "A";
  return letters.split("").reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
}

function bytesToBase64(bytes) {
  let binary = "";
  for (let index = 0; index < bytes.length; index += 1) {
    binary += String.fromCharCode(bytes[index]);
  }
  return btoa(binary);
}

function base64ToBytes(value) {
  return Uint8Array.from(atob(value), (char) => char.charCodeAt(0));
}

function constantTimeEqual(first, second) {
  if (first.length !== second.length) return false;
  let difference = 0;
  for (let index = 0; index < first.length; index += 1) difference |= first[index] ^ second[index];
  return difference === 0;
}

function sanitizeFilename(value) {
  return String(value).replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_").trim().slice(0, 50) || "利用者";
}

function escapeHtml(value) {
  return String(value).replace(/[&<>'"]/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;",
  })[char]);
}

function showMessage(element, text, type) {
  element.textContent = text;
  element.className = `message ${type || ""}`.trim();
}

function clearMessage(element) {
  element.textContent = "";
  element.className = "message";
}
