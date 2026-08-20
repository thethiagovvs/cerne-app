/* ============================================================
   Integração com Dropbox — backup automático com rotação
   OAuth 2.0 + PKCE (sem client secret, seguro para app 100% front-end)
   ============================================================ */

// Preencha com o "App key" do seu app no Dropbox App Console (dropbox.com/developers/apps).
// Não é segredo — pode ficar no código do front-end tranquilamente.
export const DROPBOX_APP_KEY = 'qmc8z2n56lgnxuj';

const TOKEN_KEY = 'cerne-dropbox-tokens-v1';
const VERIFIER_KEY = 'cerne-dropbox-pkce-verifier';
const BACKUP_PREFIX = 'backup-';
const KEEP_BACKUPS = 5;

/* ---------- PKCE helpers ---------- */

function base64url(buffer) {
  return btoa(String.fromCharCode(...new Uint8Array(buffer)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
function randomVerifier() {
  const arr = new Uint8Array(64);
  crypto.getRandomValues(arr);
  return base64url(arr.buffer);
}
async function challengeFromVerifier(verifier) {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier));
  return base64url(digest);
}
function getRedirectUri() {
  return window.location.origin + window.location.pathname;
}

/* ---------- Tokens (localStorage) ---------- */

function getTokens() {
  try { return JSON.parse(window.localStorage.getItem(TOKEN_KEY)) || null; } catch { return null; }
}
function setTokens(tokens) {
  window.localStorage.setItem(TOKEN_KEY, JSON.stringify(tokens));
}
export function disconnectDropbox() {
  window.localStorage.removeItem(TOKEN_KEY);
}
export function isDropboxConnected() {
  return !!getTokens()?.refresh_token;
}
export function isDropboxConfigured() {
  return !!DROPBOX_APP_KEY && DROPBOX_APP_KEY !== 'COLE_AQUI_SEU_APP_KEY';
}

/* ---------- Fluxo de conexão ---------- */

export async function startDropboxConnect() {
  const verifier = randomVerifier();
  window.sessionStorage.setItem(VERIFIER_KEY, verifier);
  const challenge = await challengeFromVerifier(verifier);
  const params = new URLSearchParams({
    client_id: DROPBOX_APP_KEY,
    response_type: 'code',
    code_challenge: challenge,
    code_challenge_method: 'S256',
    token_access_type: 'offline',
    redirect_uri: getRedirectUri(),
  });
  window.location.href = `https://www.dropbox.com/oauth2/authorize?${params.toString()}`;
}

// Chamar uma vez ao carregar o app. Retorna { status: 'connected' | 'none' | 'error', message? }.
// 'none' = não voltamos de um login do Dropbox (fluxo normal). 'error' = voltamos, mas algo falhou.
export async function handleDropboxRedirect() {
  const url = new URL(window.location.href);
  const code = url.searchParams.get('code');
  const oauthError = url.searchParams.get('error_description') || url.searchParams.get('error');
  if (!code && !oauthError) return { status: 'none' };

  const verifier = window.sessionStorage.getItem(VERIFIER_KEY);
  window.sessionStorage.removeItem(VERIFIER_KEY);
  url.searchParams.delete('code');
  url.searchParams.delete('state');
  url.searchParams.delete('error');
  url.searchParams.delete('error_description');
  window.history.replaceState({}, '', url.pathname + (url.search || ''));

  if (oauthError) return { status: 'error', message: oauthError };
  if (!verifier) return { status: 'error', message: 'Sessão de login expirada. Tente conectar novamente.' };

  const body = new URLSearchParams({
    code, grant_type: 'authorization_code', client_id: DROPBOX_APP_KEY,
    code_verifier: verifier, redirect_uri: getRedirectUri(),
  });
  let res;
  try {
    res = await fetch('https://api.dropboxapi.com/oauth2/token', {
      method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body,
    });
  } catch (e) {
    return { status: 'error', message: 'Falha de rede ao conectar com o Dropbox.' };
  }
  if (!res.ok) {
    const errBody = await res.json().catch(() => null);
    return { status: 'error', message: errBody?.error_description || `Erro ao conectar com o Dropbox (${res.status}).` };
  }
  const data = await res.json();
  setTokens({ access_token: data.access_token, refresh_token: data.refresh_token, expires_at: Date.now() + data.expires_in * 1000 });
  return { status: 'connected' };
}

/* ---------- Chamadas autenticadas ---------- */

async function refreshAccessToken() {
  const tokens = getTokens();
  if (!tokens?.refresh_token) throw new Error('Dropbox não conectado.');
  const body = new URLSearchParams({ grant_type: 'refresh_token', refresh_token: tokens.refresh_token, client_id: DROPBOX_APP_KEY });
  const res = await fetch('https://api.dropboxapi.com/oauth2/token', {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body,
  });
  if (!res.ok) throw new Error('Não foi possível renovar o acesso ao Dropbox. Reconecte sua conta.');
  const data = await res.json();
  const updated = { ...tokens, access_token: data.access_token, expires_at: Date.now() + data.expires_in * 1000 };
  setTokens(updated);
  return updated.access_token;
}
async function getValidAccessToken() {
  const tokens = getTokens();
  if (!tokens) throw new Error('Dropbox não conectado.');
  if (tokens.expires_at - Date.now() > 60_000) return tokens.access_token;
  return refreshAccessToken();
}
async function dbxFetch(url, options) {
  const token = await getValidAccessToken();
  const res = await fetch(url, { ...options, headers: { ...(options.headers || {}), Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Erro na comunicação com o Dropbox (${res.status}).`);
  return res;
}

/* ---------- Backups (com escopo "App folder" — o app só enxerga sua própria pasta) ---------- */

export async function listBackups() {
  const res = await dbxFetch('https://api.dropboxapi.com/2/files/list_folder', {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ path: '' }),
  });
  const data = await res.json();
  return (data.entries || [])
    .filter((e) => e['.tag'] === 'file' && e.name.startsWith(BACKUP_PREFIX))
    .sort((a, b) => a.name.localeCompare(b.name));
}

export async function uploadBackup(jsonString) {
  const filename = `${BACKUP_PREFIX}${new Date().toISOString().replace(/[:.]/g, '-')}.json`;
  // O upload em si é a parte que importa: se ela funcionar, o backup existe — mesmo que a
  // rotação (apagar os mais antigos) falhe logo em seguida por algum motivo transitório.
  await dbxFetch('https://content.dropboxapi.com/2/files/upload', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/octet-stream',
      'Dropbox-API-Arg': JSON.stringify({ path: `/${filename}`, mode: 'add', autorename: false, mute: true }),
    },
    body: jsonString,
  });
  try {
    const backups = await listBackups();
    const toDelete = backups.slice(0, Math.max(0, backups.length - KEEP_BACKUPS));
    for (const file of toDelete) {
      await dbxFetch('https://api.dropboxapi.com/2/files/delete_v2', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ path: file.path_lower }),
      }).catch(() => {});
    }
    return { filename, total: Math.min(backups.length, KEEP_BACKUPS) };
  } catch (e) {
    // Upload confirmado; só a limpeza dos backups antigos falhou — não é motivo pra reportar erro.
    return { filename, total: null };
  }
}

export async function downloadLatestBackup() {
  const backups = await listBackups();
  if (backups.length === 0) return null;
  const latest = backups[backups.length - 1];
  const res = await dbxFetch('https://content.dropboxapi.com/2/files/download', {
    method: 'POST', headers: { 'Dropbox-API-Arg': JSON.stringify({ path: latest.path_lower }) },
  });
  const text = await res.text();
  return { filename: latest.name, data: JSON.parse(text) };
}

/* ---------- Agendamento (evita subir um arquivo a cada tecla digitada) ---------- */

let backupTimer = null;
export function scheduleBackup(getStateFn, onDone, delay = 4000) {
  if (!isDropboxConnected()) return;
  clearTimeout(backupTimer);
  backupTimer = setTimeout(() => {
    uploadBackup(JSON.stringify(getStateFn()))
      .then((info) => onDone && onDone(null, info))
      .catch((err) => onDone && onDone(err, null));
  }, delay);
}
