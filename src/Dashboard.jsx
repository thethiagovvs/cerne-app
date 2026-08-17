import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, BarChart, Bar, AreaChart, Area,
} from 'recharts';
import {
  LayoutDashboard, ArrowLeftRight, TrendingUp, TrendingDown, Landmark, CreditCard,
  PieChart as PieChartIcon, Target, RefreshCw, BarChart3, Settings, Search, Bell,
  ChevronDown, ChevronLeft, ChevronRight, ChevronUp, Plus, Filter, Download, Upload,
  MoreHorizontal, Pencil, Trash2, Eye, X, Check, AlertCircle, Sparkles, Calendar as CalendarIcon,
  Wallet, PiggyBank, Banknote, ArrowUpRight, ArrowDownRight, Menu, Info, Sprout,
  ShoppingBag, ShoppingCart, Home, Car, Utensils, Heart, Flower2, Film, GraduationCap,
  Plane, Laptop, Award, Briefcase, Clock, ArrowRight, FileText, Loader2, Cloud,
} from 'lucide-react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import {
  isDropboxConfigured, isDropboxConnected, startDropboxConnect, handleDropboxRedirect,
  disconnectDropbox, uploadBackup, downloadLatestBackup, scheduleBackup,
} from './dropboxSync.js';

/* ============================================================
   CERNE — Dashboard financeiro pessoal
   Design tokens, dados de exemplo e componentes compartilhados
   ============================================================ */

const GLOBAL_STYLES = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Plus+Jakarta+Sans:ital,wght@0,400;0,500;0,600;0,700;1,500&display=swap');

:root {
  --primary: #5D7052;
  --primary-dark: #4C5C43;
  --primary-soft: #E9EDE4;
  --secondary: #8A9B7D;
  --bg: #F4F1EA;
  --card: #FFFFFF;
  --border: #E7E4DE;
  --text: #232323;
  --text-soft: #6D6D6D;
  --income: #5A8F5A;
  --income-soft: #E7F0E5;
  --expense: #B66B6B;
  --expense-soft: #F6EAEA;
  --invest: #6B7FAF;
  --invest-soft: #E8EBF3;
  --goals: #B9A88A;
  --goals-soft: #F2EEE3;
  --alert: #D8A84C;
  --alert-soft: #FBF1E0;
}

.cerne-root, .cerne-root * { font-family: 'Plus Jakarta Sans', ui-sans-serif, system-ui, sans-serif; box-sizing: border-box; }
.font-display { font-family: 'Space Grotesk', ui-sans-serif, system-ui, sans-serif; }

.cerne-root ::-webkit-scrollbar { width: 8px; height: 8px; }
.cerne-root ::-webkit-scrollbar-track { background: transparent; }
.cerne-root ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 8px; }
.cerne-root ::-webkit-scrollbar-thumb:hover { background: var(--secondary); }

.shadow-soft { box-shadow: 0 1px 2px rgba(35,35,35,0.04), 0 4px 14px rgba(35,35,35,0.05); }
.shadow-soft-lg { box-shadow: 0 2px 6px rgba(35,35,35,0.05), 0 14px 28px rgba(35,35,35,0.08); }

.focus-ring:focus { outline: none; box-shadow: 0 0 0 3px rgba(93,112,82,0.22); }
.focus-ring:focus-visible { outline: none; box-shadow: 0 0 0 3px rgba(93,112,82,0.22); }

@keyframes fadeSlideUp { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
.animate-fade-up { animation: fadeSlideUp 0.4s cubic-bezier(0.16,1,0.3,1) both; }

@keyframes toastIn { from { opacity: 0; transform: translateY(10px) scale(0.98); } to { opacity: 1; transform: translateY(0) scale(1); } }
.animate-toast { animation: toastIn 0.25s ease-out both; }

@keyframes ringPulse { 0% { transform: scale(0.85); opacity: 0.6; } 100% { transform: scale(1.6); opacity: 0; } }
.ring-pulse { animation: ringPulse 1.8s ease-out infinite; }

@keyframes shimmer { 0% { background-position: -200px 0; } 100% { background-position: 200px 0; } }
.skeleton { background: linear-gradient(90deg, #EDEAE2 25%, #F7F5EF 37%, #EDEAE2 63%); background-size: 400px 100%; animation: shimmer 1.4s ease-in-out infinite; }

@media print {
  .no-print { display: none !important; }
  .print-area { padding: 0 !important; overflow: visible !important; }
}
`;

const CATEGORIES = {
  'Moradia': { color: '#8A9B7D', soft: '#EEF1EA', icon: Home },
  'Mercado': { color: '#C98A5E', soft: '#F5EBE2', icon: ShoppingCart },
  'Alimentação': { color: '#D4A574', soft: '#F7EFE4', icon: Utensils },
  'Transporte': { color: '#7B93A8', soft: '#EBEFF3', icon: Car },
  'Saúde': { color: '#C97A7A', soft: '#F6EAEA', icon: Heart },
  'Cuidado e Beleza': { color: '#B98DAF', soft: '#F3EBF1', icon: Flower2 },
  'Lazer': { color: '#6FA8A0', soft: '#E9F2F0', icon: Film },
  'Compras': { color: '#B9A24C', soft: '#F5F1E2', icon: ShoppingBag },
  'Outros': { color: '#A8A398', soft: '#F1EFEA', icon: MoreHorizontal },
};
const CATEGORY_NAMES = Object.keys(CATEGORIES);

const ACCOUNTS_ICONS = { 'Conta Corrente': Wallet, 'Poupança': PiggyBank };

const PAYMENT_METHODS = ['Pix', 'Cartão de crédito', 'Cartão de débito', 'Dinheiro', 'Transferência', 'Boleto'];
const STATUS_OPTIONS = ['Pago', 'Pendente', 'Agendado'];

/* ---------- Helpers ---------- */

function formatBRL(value) {
  return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value || 0);
}
function formatDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr + 'T00:00:00');
  return d.toLocaleDateString('pt-BR');
}
function formatDateLong(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  return d.toLocaleDateString('pt-BR', { day: 'numeric', month: 'long', year: 'numeric' });
}
function getGreeting() {
  const hour = new Date().getHours();
  if (hour < 12) return 'Bom dia';
  if (hour < 18) return 'Boa tarde';
  return 'Boa noite';
}
function uid() {
  try { return crypto.randomUUID(); } catch (e) { return 'id-' + Date.now() + '-' + Math.random().toString(36).slice(2, 8); }
}
function getMonthGrid(year, month) {
  const firstDay = new Date(year, month, 1);
  const startOffset = firstDay.getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const cells = [];
  for (let i = 0; i < startOffset; i++) cells.push({ day: null, key: 'pre-' + i });
  for (let d = 1; d <= daysInMonth; d++) cells.push({ day: d, key: 'day-' + d });
  return cells;
}
function isSameMonth(dateStr, year, month) {
  const d = new Date(dateStr + 'T00:00:00');
  return d.getFullYear() === year && d.getMonth() === month;
}
const MONTH_LABELS = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];

/* ---------- Agregações reais (a partir dos lançamentos, contas etc.) ---------- */

// Efeito de um lançamento sobre o saldo da conta: só lançamentos "Pago" já saíram/entraram de fato.
function transactionBalanceEffect(t) {
  if (!t || t.status !== 'Pago') return 0;
  return t.type === 'receita' ? t.amount : -t.amount;
}
function applyBalanceDelta(accountsList, accountId, delta) {
  if (!accountId || !delta) return accountsList;
  return accountsList.map((a) => (a.id === accountId ? { ...a, balance: a.balance + delta } : a));
}
// Desfaz o efeito do lançamento antigo (se houver) e aplica o efeito do novo (se houver) —
// cobre criar (oldTx null), editar (os dois) e excluir (newTx null) com a mesma lógica.
function reapplyAccountEffect(accountsList, oldTx, newTx) {
  let updated = accountsList;
  if (oldTx) updated = applyBalanceDelta(updated, oldTx.account, -transactionBalanceEffect(oldTx));
  if (newTx) updated = applyBalanceDelta(updated, newTx.account, transactionBalanceEffect(newTx));
  return updated;
}

// Histórico mensal real (receitas/despesas pagas por mês + patrimônio reconstruído a partir do
// saldo atual das contas, "desfazendo" o líquido de cada mês para trás).
function computeMonthlyHistory(transactions, accounts, investmentsTotal, monthsBack = 12) {
  const now = new Date();
  const months = [];
  for (let i = monthsBack - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    months.push({ year: d.getFullYear(), month: d.getMonth(), label: `${MONTH_LABELS[d.getMonth()]}/${String(d.getFullYear()).slice(2)}` });
  }
  const byKey = {};
  transactions.forEach((t) => {
    if (t.status !== 'Pago') return;
    const d = new Date(t.date + 'T00:00:00');
    const key = `${d.getFullYear()}-${d.getMonth()}`;
    if (!byKey[key]) byKey[key] = { receitas: 0, despesas: 0 };
    if (t.type === 'receita') byKey[key].receitas += t.amount; else byKey[key].despesas += t.amount;
  });
  const currentBalance = accounts.reduce((s, a) => s + a.balance, 0) + (investmentsTotal || 0);
  const result = new Array(months.length);
  let carry = currentBalance;
  for (let i = months.length - 1; i >= 0; i--) {
    const m = months[i];
    const data = byKey[`${m.year}-${m.month}`] || { receitas: 0, despesas: 0 };
    result[i] = { month: m.label, receitas: data.receitas, despesas: data.despesas, patrimonio: carry };
    carry -= (data.receitas - data.despesas);
  }
  return result;
}

// Gastos por categoria (despesas pagas) em um mês específico.
function computeCategoryTotals(transactions, year, month) {
  const totals = {};
  transactions.filter((t) => t.type === 'despesa' && t.status === 'Pago' && isSameMonth(t.date, year, month)).forEach((t) => {
    totals[t.category] = (totals[t.category] || 0) + t.amount;
  });
  return totals;
}

/* ---------- Dias úteis / feriados nacionais (para vencimento adaptativo) ---------- */

function ymd(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}
function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}
// Algoritmo de Meeus/Jones/Butcher para o Domingo de Páscoa (base dos feriados móveis).
function computeEaster(year) {
  const a = year % 19, b = Math.floor(year / 100), c = year % 100;
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month - 1, day);
}
function getBrazilianHolidays(year) {
  const easter = computeEaster(year);
  const dates = [
    new Date(year, 0, 1), addDays(easter, -48), addDays(easter, -47), addDays(easter, -2), addDays(easter, 60),
    new Date(year, 3, 21), new Date(year, 4, 1), new Date(year, 8, 7), new Date(year, 9, 12),
    new Date(year, 10, 2), new Date(year, 10, 15), new Date(year, 10, 20), new Date(year, 11, 25),
  ];
  return new Set(dates.map(ymd));
}
function isBusinessDay(date) {
  const day = date.getDay();
  if (day === 0 || day === 6) return false;
  return !getBrazilianHolidays(date.getFullYear()).has(ymd(date));
}
function nextBusinessDay(date) {
  let d = new Date(date);
  while (!isBusinessDay(d)) d = addDays(d, 1);
  return d;
}
// Data de vencimento "adaptativa": se o dia configurado cair em fim de semana/feriado,
// o vencimento é antecipado para o próximo dia útil (regra bancária padrão no Brasil).
function getAdjustedDueDate(year, month, dueDay) {
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const raw = new Date(year, month, Math.min(dueDay, daysInMonth));
  return nextBusinessDay(raw);
}
// Próxima data de vencimento a partir de uma data de referência (hoje, por padrão).
function getNextCardDueDate(card, referenceDate = new Date()) {
  const ref = new Date(referenceDate.getFullYear(), referenceDate.getMonth(), referenceDate.getDate());
  let due = getAdjustedDueDate(ref.getFullYear(), ref.getMonth(), card.dueDay);
  if (due < ref) due = getAdjustedDueDate(ref.getFullYear(), ref.getMonth() + 1, card.dueDay);
  return due;
}

/* ---------- Cálculo de salário líquido CLT (INSS + IRRF, tabelas 2026) ---------- */

const INSS_BRACKETS_2026 = [
  { upTo: 1621.00, rate: 0.075 },
  { upTo: 2902.84, rate: 0.09 },
  { upTo: 4354.27, rate: 0.12 },
  { upTo: 8475.55, rate: 0.14 },
];
const IRRF_BRACKETS_2026 = [
  { upTo: 2428.80, rate: 0, deduction: 0 },
  { upTo: 2826.65, rate: 0.075, deduction: 182.16 },
  { upTo: 3751.05, rate: 0.15, deduction: 394.16 },
  { upTo: 4664.68, rate: 0.225, deduction: 675.49 },
  { upTo: Infinity, rate: 0.275, deduction: 908.73 },
];
const IRRF_DEPENDENT_DEDUCTION_2026 = 189.59;
const IRRF_SIMPLIFIED_DEDUCTION_2026 = 607.20;

function calcINSS(gross) {
  const capped = Math.min(gross, 8475.55);
  let inss = 0, lower = 0;
  for (const bracket of INSS_BRACKETS_2026) {
    if (capped <= lower) break;
    inss += (Math.min(capped, bracket.upTo) - lower) * bracket.rate;
    lower = bracket.upTo;
  }
  return Math.round(inss * 100) / 100;
}
function calcIRRF(gross, inss, dependents) {
  const deduction = Math.max(inss + dependents * IRRF_DEPENDENT_DEDUCTION_2026, IRRF_SIMPLIFIED_DEDUCTION_2026);
  const base = Math.max(0, gross - deduction);
  const bracket = IRRF_BRACKETS_2026.find((b) => base <= b.upTo);
  const irrfBruto = Math.max(0, base * bracket.rate - bracket.deduction);
  // Redutor da Lei 15.270/2025, em vigor desde jan/2026: isenção total até R$5.000 e
  // redução decrescente entre R$5.000,01 e R$7.350,00 de rendimento tributável.
  if (gross <= 5000) return 0;
  if (gross <= 7350) {
    const reducao = Math.max(0, 978.62 - 0.133145 * gross);
    return Math.round(Math.max(0, irrfBruto - reducao) * 100) / 100;
  }
  return Math.round(irrfBruto * 100) / 100;
}
function calcCLTNetSalary(gross, dependents = 0) {
  const g = Number(gross) || 0;
  const inss = calcINSS(g);
  const irrf = calcIRRF(g, inss, Number(dependents) || 0);
  return { inss, irrf, net: Math.max(0, Math.round((g - inss - irrf) * 100) / 100) };
}

/* ---------- Metas: estimativa de conclusão a partir do histórico de aportes ---------- */

function estimateGoalCompletion(goal) {
  const remaining = goal.target - goal.current;
  if (remaining <= 0) return { status: 'done' };
  const history = goal.history || [];
  if (history.length === 0) return { status: 'no-data' };
  const byMonth = {};
  history.forEach((h) => { byMonth[h.date.slice(0, 7)] = (byMonth[h.date.slice(0, 7)] || 0) + h.amount; });
  const months = Object.keys(byMonth);
  const avgMonthly = Object.values(byMonth).reduce((s, v) => s + v, 0) / months.length;
  if (avgMonthly <= 0) return { status: 'no-data' };
  const monthsNeeded = Math.ceil(remaining / avgMonthly);
  const eta = new Date();
  eta.setMonth(eta.getMonth() + monthsNeeded);
  return { status: 'ok', avgMonthly, monthsNeeded, etaLabel: eta.toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' }) };
}

/* ---------- Importação de faturas (CSV) ---------- */

function stripDiacritics(str) {
  return String(str || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

// Converte valores no formato do extrato Nubank ("101,15", "- 101,15", "1.234,56") para número.
function parseBRLAmount(raw) {
  if (raw === null || raw === undefined) return 0;
  let s = String(raw).trim();
  const negative = /^-/.test(s.replace(/\s+/g, ''));
  s = s.replace(/[^\d,.-]/g, '').replace(/^-/, '');
  if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
  const value = parseFloat(s) || 0;
  return negative ? -value : value;
}

// Reconhece "Nome da compra - Parcela 3/5" e "Pix no Crédito - Fulano - 1/2".
const INSTALLMENT_REGEX = /\s*-\s*Parcela\s+(\d+)\/(\d+)\s*$|\s*-\s*(\d+)\/(\d+)\s*$/i;

function parseInstallment(title) {
  const match = String(title || '').match(INSTALLMENT_REGEX);
  if (!match) return null;
  const current = match[1] || match[3];
  const total = match[2] || match[4];
  if (!current || !total) return null;
  return { clean: title.replace(INSTALLMENT_REGEX, '').trim(), current: parseInt(current, 10), total: parseInt(total, 10) };
}

// No extrato do Nubank, "Pagamento recebido" é o adiantamento/parcela da fatura que você mesmo
// paga (não é um gasto novo) — precisa ser identificado e mantido fora da importação por padrão,
// senão duplicaria o valor que já sai da sua conta corrente.
function isPaymentEntry(title) {
  return /pagamento\s+recebido/i.test(String(title || ''));
}

const CATEGORY_KEYWORDS = [
  { category: 'Alimentação', keywords: ['ifood', 'ifd*', 'restaurante', 'lanchonete', 'padaria', 'burger', 'pizza'] },
  { category: 'Mercado', keywords: ['mercado', 'supermercado', 'atacad', 'hortifruti'] },
  { category: 'Transporte', keywords: ['99app', '99*', 'uber', 'combust', 'posto ', 'estacionamento'] },
  { category: 'Saúde', keywords: ['totalpass', 'farmacia', 'farmácia', 'drogaria', 'academia'] },
  { category: 'Lazer', keywords: ['steam', 'netflix', 'spotify', 'ingresso', 'cinema', 'nupay'] },
  { category: 'Compras', keywords: ['mercadolivre', 'mercado livre', 'shopee', 'amazon', 'shein', 'aliexpress'] },
  { category: 'Cuidado e Beleza', keywords: ['salao', 'salão', 'barbearia', 'cosmetic'] },
  { category: 'Moradia', keywords: ['aluguel', 'condominio', 'condomínio', 'energia', 'sanepar', 'copel', 'claro', 'vivo', 'tim', 'oi fibra'] },
];

function guessCategory(title) {
  const t = stripDiacritics(String(title || '')).toLowerCase();
  for (const group of CATEGORY_KEYWORDS) {
    if (group.keywords.some((k) => t.includes(stripDiacritics(k).toLowerCase()))) return group.category;
  }
  return 'Outros';
}

function normalizeForFingerprint(description) {
  return stripDiacritics(String(description || ''))
    .toLowerCase()
    .replace(/parcela\s*\d+\/\d+/g, '')
    .replace(/\d+\/\d+\s*$/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

// Identidade de um lançamento para fins de deduplicação: mesma data + descrição normalizada + valor.
function computeFingerprint(date, description, amount) {
  return `${date}|${normalizeForFingerprint(description)}|${Math.abs(amount).toFixed(2)}`;
}

function buildFingerprintMultiset(transactions) {
  const map = new Map();
  transactions.forEach((t) => {
    const fp = computeFingerprint(t.date, t.description, t.amount);
    map.set(fp, (map.get(fp) || 0) + 1);
  });
  return map;
}

// Faz o parse do CSV de fatura do Nubank (colunas: date,title,amount) separando compras reais
// de pagamentos/adiantamentos de fatura, e sinalizando duplicatas contra o que já existe no app.
function parseNubankFaturaCSV(rows, existingTransactions) {
  const available = buildFingerprintMultiset(existingTransactions);
  const purchases = [];
  const payments = [];
  let paymentsTotal = 0;

  rows.forEach((row, idx) => {
    const keys = Object.keys(row);
    const find = (patterns) => keys.find((k) => patterns.some((p) => k.toLowerCase().includes(p)));
    const dateKey = find(['data', 'date']);
    const titleKey = find(['title', 'descri', 'histor']);
    const amountKey = find(['amount', 'valor', 'value']);
    if (!dateKey || !titleKey || !amountKey) return;

    const rawTitle = String(row[titleKey] || '').trim();
    if (!rawTitle) return;
    const rawDate = String(row[dateKey] || '').trim();
    const amount = parseBRLAmount(row[amountKey]);

    if (isPaymentEntry(rawTitle)) {
      const value = Math.abs(amount);
      payments.push({ rowId: `pay-${idx}`, date: rawDate, title: rawTitle, amount: value });
      paymentsTotal += value;
      return;
    }

    const installment = parseInstallment(rawTitle);
    const description = installment ? installment.clean : rawTitle;
    const fingerprint = computeFingerprint(rawDate, description, amount);
    const remaining = available.get(fingerprint) || 0;
    const isDuplicate = remaining > 0;
    if (isDuplicate) available.set(fingerprint, remaining - 1);

    purchases.push({
      rowId: `row-${idx}`,
      date: rawDate,
      description,
      category: guessCategory(description),
      amount: Math.abs(amount),
      installment: installment ? { current: installment.current, total: installment.total } : null,
      fingerprint,
      isDuplicate,
      include: !isDuplicate,
    });
  });

  return { purchases, payments, paymentsTotal };
}

/* ---------- Dados de exemplo ---------- */

const initialAccounts = [
  { id: 'acc-1', bank: 'Nubank', type: 'Conta Corrente', balance: 3200, alertThreshold: 500 },
  { id: 'acc-2', bank: 'Itaú', type: 'Conta Corrente', balance: 4150, alertThreshold: 0 },
  { id: 'acc-3', bank: 'Inter', type: 'Poupança', balance: 1100, alertThreshold: 0 },
];

const initialCards = [
  { id: 'card-1', bank: 'Nubank', brand: 'Mastercard Black', limit: 8000, used: 6240, invoice: 2180, closingDay: 22, dueDay: 29 },
  { id: 'card-2', bank: 'Itaú', brand: 'Visa Click', limit: 5000, used: 1870, invoice: 1450, closingDay: 5, dueDay: 12 },
  { id: 'card-3', bank: 'Inter', brand: 'Gold Mastercard', limit: 4000, used: 980, invoice: 620, closingDay: 15, dueDay: 22 },
];

const initialBenefits = [
  { id: 'benefit-1', provider: 'Flash', foodBalance: 380, mobilityBalance: 150 },
];

const initialGoals = [
  { id: 'goal-1', name: 'Reserva de emergência', target: 20000, current: 18200, deadline: '2026-10-31', icon: 'Wallet', history: [
    { date: '2026-05-10', amount: 900 }, { date: '2026-06-10', amount: 900 }, { date: '2026-07-10', amount: 900 },
  ] },
  { id: 'goal-2', name: 'Viagem de fim de ano', target: 6000, current: 2400, deadline: '2026-12-15', icon: 'Plane', history: [
    { date: '2026-06-05', amount: 1200 }, { date: '2026-07-05', amount: 1200 },
  ] },
  { id: 'goal-3', name: 'Notebook novo', target: 5500, current: 1650, deadline: '2027-03-01', icon: 'Laptop', history: [
    { date: '2026-06-20', amount: 750 }, { date: '2026-07-20', amount: 900 },
  ] },
  { id: 'goal-4', name: 'Certificação em TI', target: 1800, current: 900, deadline: '2026-11-30', icon: 'Award', history: [
    { date: '2026-07-15', amount: 900 },
  ] },
];
const GOAL_ICONS = { Wallet, Plane, Laptop, Award };

const initialCaixinhas = [
  { id: 'cx-1', name: 'Imprevistos do dia a dia', balance: 340, accountId: 'acc-1' },
  { id: 'cx-2', name: 'Presentes e datas especiais', balance: 480, accountId: 'acc-2' },
  { id: 'cx-3', name: 'Manutenção do carro', balance: 620, accountId: 'acc-2' },
];

const initialRecurring = [
  { id: 'rec-1', name: 'Netflix', category: 'Lazer', value: 44.90, renewalDay: 14 },
  { id: 'rec-2', name: 'Spotify', category: 'Lazer', value: 21.90, renewalDay: 8 },
  { id: 'rec-3', name: 'Amazon Prime', category: 'Lazer', value: 14.90, renewalDay: 20 },
  { id: 'rec-4', name: 'iCloud+', category: 'Lazer', value: 12.90, renewalDay: 3 },
  { id: 'rec-5', name: 'YouTube Premium', category: 'Lazer', value: 24.90, renewalDay: 18 },
  { id: 'rec-6', name: 'Academia', category: 'Saúde', value: 99.90, renewalDay: 5 },
  { id: 'rec-7', name: 'Seguro do celular', category: 'Compras', value: 29.90, renewalDay: 10 },
];

const INVESTMENT_CATEGORIES = {
  'Renda Fixa': '#8A9B7D',
  'Ações': '#6B7FAF',
  'Fundos Imobiliários': '#B9A88A',
  'Criptomoedas': '#D8A84C',
  'Reserva de Liquidez': '#7B93A8',
  'Outros': '#A8A398',
};
const INVESTMENT_CATEGORY_NAMES = Object.keys(INVESTMENT_CATEGORIES);

const initialInvestments = [
  { id: 'inv-1', name: 'Tesouro Selic', category: 'Renda Fixa', invested: 10200, currentValue: 11160 },
  { id: 'inv-2', name: 'Ações (carteira)', category: 'Ações', invested: 5400, currentValue: 6200 },
  { id: 'inv-3', name: 'Fundos Imobiliários', category: 'Fundos Imobiliários', invested: 4700, currentValue: 4960 },
  { id: 'inv-4', name: 'Reserva de liquidez (CDB)', category: 'Reserva de Liquidez', invested: 2350, currentValue: 2480 },
];

function buildInitialTransactions() {
  const rows = [
    ['2026-08-01', 'Salário', 'Outros', 'receita', 'acc-2', 'Transferência', 6200, 'Pago'],
    ['2026-08-02', 'Supermercado Extra', 'Mercado', 'despesa', 'acc-1', 'Cartão de débito', 312, 'Pago'],
    ['2026-08-03', 'Aluguel', 'Moradia', 'despesa', 'acc-2', 'Transferência', 1150, 'Pago'],
    ['2026-08-03', 'Condomínio', 'Moradia', 'despesa', 'acc-2', 'Boleto', 300, 'Pago'],
    ['2026-08-04', 'iFood', 'Alimentação', 'despesa', 'acc-1', 'Cartão de crédito', 68, 'Pago'],
    ['2026-08-05', 'Academia', 'Saúde', 'despesa', 'acc-1', 'Cartão de crédito', 99.90, 'Pago'],
    ['2026-08-05', 'Posto Ipiranga', 'Transporte', 'despesa', 'acc-1', 'Cartão de débito', 180, 'Pago'],
    ['2026-08-06', 'Farmácia São João', 'Saúde', 'despesa', 'acc-1', 'Pix', 74, 'Pago'],
    ['2026-08-06', 'Salão de beleza', 'Cuidado e Beleza', 'despesa', 'acc-1', 'Pix', 90, 'Pago'],
    ['2026-08-07', 'Uber', 'Transporte', 'despesa', 'acc-1', 'Cartão de crédito', 32, 'Pago'],
    ['2026-08-08', 'Spotify', 'Lazer', 'despesa', 'acc-1', 'Cartão de crédito', 21.90, 'Pago'],
    ['2026-08-08', 'Restaurante', 'Alimentação', 'despesa', 'acc-1', 'Cartão de crédito', 96, 'Pago'],
    ['2026-08-09', 'Farmácia (cosméticos)', 'Cuidado e Beleza', 'despesa', 'acc-1', 'Cartão de débito', 58, 'Pago'],
    ['2026-08-10', 'Seguro do celular', 'Compras', 'despesa', 'acc-2', 'Cartão de crédito', 29.90, 'Pago'],
    ['2026-08-10', 'Padaria', 'Alimentação', 'despesa', 'acc-1', 'Dinheiro', 24, 'Pago'],
    ['2026-08-11', 'Freelance design', 'Outros', 'receita', 'acc-1', 'Pix', 850, 'Pago'],
    ['2026-08-12', 'Supermercado Carrefour', 'Mercado', 'despesa', 'acc-1', 'Cartão de débito', 205, 'Pago'],
    ['2026-08-12', 'Internet + TV', 'Moradia', 'despesa', 'acc-2', 'Boleto', 140, 'Pago'],
    ['2026-08-13', 'Cinema', 'Lazer', 'despesa', 'acc-1', 'Cartão de crédito', 62, 'Pago'],
    ['2026-08-14', 'Netflix', 'Lazer', 'despesa', 'acc-1', 'Cartão de crédito', 44.90, 'Pago'],
    ['2026-08-14', 'Estacionamento', 'Transporte', 'despesa', 'acc-1', 'Dinheiro', 18, 'Pago'],
    ['2026-08-15', 'Barbearia', 'Cuidado e Beleza', 'despesa', 'acc-1', 'Pix', 55, 'Pago'],
    ['2026-08-16', 'Loja de roupas', 'Compras', 'despesa', 'acc-1', 'Cartão de crédito', 189, 'Pago'],
    ['2026-08-17', 'Energia elétrica', 'Moradia', 'despesa', 'acc-2', 'Boleto', 210, 'Pendente'],
    ['2026-08-17', 'Aporte investimentos', 'Outros', 'despesa', 'acc-2', 'Transferência', 1116, 'Pago'],
    ['2026-08-18', 'YouTube Premium', 'Lazer', 'despesa', 'acc-1', 'Cartão de crédito', 24.90, 'Pago'],
    ['2026-08-18', 'Restaurante japonês', 'Alimentação', 'despesa', 'acc-1', 'Cartão de crédito', 132, 'Pago'],
    ['2026-08-19', 'Farmácia', 'Saúde', 'despesa', 'acc-1', 'Pix', 45, 'Pago'],
    ['2026-08-20', 'Amazon Prime', 'Lazer', 'despesa', 'acc-1', 'Cartão de crédito', 14.90, 'Pago'],
    ['2026-08-20', 'Água', 'Moradia', 'despesa', 'acc-2', 'Boleto', 95, 'Pendente'],
    ['2026-08-22', 'Fatura cartão Nubank', 'Outros', 'despesa', 'acc-1', 'Boleto', 2180, 'Agendado'],
    ['2026-07-01', 'Salário', 'Outros', 'receita', 'acc-2', 'Transferência', 6200, 'Pago'],
    ['2026-07-02', 'Supermercado', 'Mercado', 'despesa', 'acc-1', 'Cartão de débito', 260, 'Pago'],
    ['2026-07-03', 'Aluguel', 'Moradia', 'despesa', 'acc-2', 'Transferência', 1150, 'Pago'],
    ['2026-07-03', 'Condomínio', 'Moradia', 'despesa', 'acc-2', 'Boleto', 300, 'Pago'],
    ['2026-07-05', 'Academia', 'Saúde', 'despesa', 'acc-1', 'Cartão de crédito', 99.90, 'Pago'],
    ['2026-07-07', 'Presente aniversário', 'Compras', 'despesa', 'acc-1', 'Cartão de crédito', 220, 'Pago'],
    ['2026-07-09', 'Viagem final de semana', 'Lazer', 'despesa', 'acc-1', 'Cartão de crédito', 480, 'Pago'],
    ['2026-07-11', 'Freelance', 'Outros', 'receita', 'acc-1', 'Pix', 620, 'Pago'],
    ['2026-07-14', 'Netflix', 'Lazer', 'despesa', 'acc-1', 'Cartão de crédito', 44.90, 'Pago'],
    ['2026-07-15', 'Consulta médica', 'Saúde', 'despesa', 'acc-1', 'Pix', 220, 'Pago'],
    ['2026-07-18', 'Combustível', 'Transporte', 'despesa', 'acc-1', 'Cartão de débito', 260, 'Pago'],
    ['2026-07-20', 'Restaurante', 'Alimentação', 'despesa', 'acc-1', 'Cartão de crédito', 145, 'Pago'],
    ['2026-07-22', 'Cosméticos', 'Cuidado e Beleza', 'despesa', 'acc-1', 'Cartão de débito', 130, 'Pago'],
    ['2026-06-01', 'Salário', 'Outros', 'receita', 'acc-2', 'Transferência', 6300, 'Pago'],
    ['2026-06-03', 'Aluguel', 'Moradia', 'despesa', 'acc-2', 'Transferência', 1150, 'Pago'],
    ['2026-06-10', 'Supermercado', 'Mercado', 'despesa', 'acc-1', 'Cartão de débito', 290, 'Pago'],
    ['2026-06-15', 'Manutenção do carro', 'Transporte', 'despesa', 'acc-1', 'Pix', 380, 'Pago'],
    ['2026-06-20', 'Roupas de inverno', 'Compras', 'despesa', 'acc-1', 'Cartão de crédito', 310, 'Pago'],
  ];
  return rows.map((r) => ({
    id: uid(), date: r[0], description: r[1], category: r[2], type: r[3],
    account: r[4], paymentMethod: r[5], amount: r[6], status: r[7],
  }));
}

/* ---------- Persistência (localStorage do navegador) ---------- */

const STORAGE_KEY = 'cerne-app-data-v1';

async function loadAppData() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}
async function saveAppData(data) {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  } catch (e) {
    /* melhor esforço, sem bloquear a interface (ex: modo anônimo com storage bloqueado) */
  }
}

/* ---------- Insights inteligentes ---------- */

function generateInsights({ transactions, goals, monthlyHistory: mh, accounts = [], categoryComparison }) {
  const insights = [];
  const thisMonth = mh[mh.length - 1];
  const lastMonth = mh.length > 1 ? mh[mh.length - 2] : null;

  accounts.filter((a) => a.alertThreshold > 0 && a.balance <= a.alertThreshold).forEach((a) => {
    insights.unshift({
      icon: AlertCircle, tone: 'expense',
      text: `Saldo da conta ${a.bank} está em ${formatBRL(a.balance)}, abaixo do limite de alerta que você definiu (${formatBRL(a.alertThreshold)}).`,
    });
  });

  if (lastMonth) {
    const savedThis = thisMonth.receitas - thisMonth.despesas;
    const savedLast = lastMonth.receitas - lastMonth.despesas;
    const diff = savedThis - savedLast;
    if (Math.abs(diff) > 1) {
      insights.push({
        icon: diff > 0 ? TrendingUp : TrendingDown,
        tone: diff > 0 ? 'income' : 'expense',
        text: diff > 0
          ? `Você economizou ${formatBRL(diff)} a mais em comparação ao mês anterior.`
          : `Você economizou ${formatBRL(Math.abs(diff))} a menos em comparação ao mês anterior.`,
      });
    }
  }

  const catNow = categoryComparison.current;
  const catPrev = categoryComparison.previous;
  const topCategory = Object.entries(catNow).sort((a, b) => b[1] - a[1])[0];
  if (topCategory) {
    insights.push({
      icon: AlertCircle, tone: 'alert',
      text: `Seu maior gasto este mês continua sendo ${topCategory[0].toLowerCase()}, somando ${formatBRL(topCategory[1])}.`,
    });
  }

  let biggestJump = null;
  Object.keys(catNow).forEach((cat) => {
    if (!catPrev[cat]) return;
    const pct = ((catNow[cat] - catPrev[cat]) / catPrev[cat]) * 100;
    if (pct > 5 && (!biggestJump || pct > biggestJump.pct)) biggestJump = { cat, pct };
  });
  if (biggestJump) {
    insights.push({
      icon: TrendingUp, tone: 'expense',
      text: `Seus gastos com ${biggestJump.cat.toLowerCase()} aumentaram ${biggestJump.pct.toFixed(0)}% este mês.`,
    });
  }

  const emergencyGoal = goals.find((g) => g.name.toLowerCase().includes('reserva'));
  if (emergencyGoal && emergencyGoal.current < emergencyGoal.target) {
    const estimate = estimateGoalCompletion(emergencyGoal);
    if (estimate.status === 'ok') {
      insights.push({
        icon: Target, tone: 'goals',
        text: `Se mantiver este ritmo, sua ${emergencyGoal.name.toLowerCase()} será concluída em ${estimate.etaLabel}.`,
      });
    }
  }

  return insights;
}

/* ============================================================
   COMPONENTES COMPARTILHADOS
   ============================================================ */

function RingProgress({ percent, size = 56, strokeWidth = 6, color = 'var(--primary)', trackColor = 'var(--border)', children }) {
  const radius = (size - strokeWidth) / 2;
  const circumference = 2 * Math.PI * radius;
  const safePercent = Math.max(0, Math.min(100, percent || 0));
  const offset = circumference - (safePercent / 100) * circumference;
  return (
    <div className="relative flex items-center justify-center" style={{ width: size, height: size }}>
      <svg width={size} height={size} style={{ transform: 'rotate(-90deg)' }}>
        <circle cx={size / 2} cy={size / 2} r={radius} fill="none" stroke={trackColor} strokeWidth={strokeWidth} />
        <circle
          cx={size / 2} cy={size / 2} r={radius} fill="none" stroke={color} strokeWidth={strokeWidth}
          strokeDasharray={circumference} strokeDashoffset={offset} strokeLinecap="round"
          style={{ transition: 'stroke-dashoffset 0.8s cubic-bezier(0.16,1,0.3,1)' }}
        />
      </svg>
      <div className="absolute inset-0 flex items-center justify-center">{children}</div>
    </div>
  );
}

function ProgressBar({ percent, color = 'var(--primary)', trackColor = 'var(--border)', height = 8 }) {
  const safePercent = Math.max(0, Math.min(100, percent || 0));
  return (
    <div className="w-full rounded-full overflow-hidden" style={{ backgroundColor: trackColor, height }}>
      <div className="h-full rounded-full" style={{ width: `${safePercent}%`, backgroundColor: color, transition: 'width 0.6s cubic-bezier(0.16,1,0.3,1)' }} />
    </div>
  );
}

function Badge({ children, color, soft, icon: Icon }) {
  return (
    <span
      className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium whitespace-nowrap"
      style={{ backgroundColor: soft, color }}
    >
      {Icon && <Icon size={12} strokeWidth={2.5} />}
      {children}
    </span>
  );
}

function CategoryBadge({ category }) {
  const cat = CATEGORIES[category] || CATEGORIES['Outros'];
  return <Badge color={cat.color} soft={cat.soft} icon={cat.icon}>{category}</Badge>;
}

function StatusBadge({ status }) {
  const map = {
    'Pago': { color: 'var(--income)', soft: 'var(--income-soft)' },
    'Pendente': { color: 'var(--alert)', soft: 'var(--alert-soft)' },
    'Agendado': { color: 'var(--text-soft)', soft: '#EEEDE8' },
  };
  const s = map[status] || map['Pendente'];
  return <Badge color={s.color} soft={s.soft}>{status}</Badge>;
}

function IconCircle({ icon: Icon, color, soft, size = 40 }) {
  return (
    <div className="rounded-xl flex items-center justify-center shrink-0" style={{ width: size, height: size, backgroundColor: soft }}>
      <Icon size={size * 0.45} color={color} strokeWidth={2.2} />
    </div>
  );
}

function Button({ children, variant = 'primary', size = 'md', icon: Icon, onClick, type = 'button', className = '', disabled }) {
  const base = 'inline-flex items-center justify-center gap-2 rounded-xl font-medium transition-all focus-ring disabled:opacity-50 disabled:cursor-not-allowed';
  const sizes = { sm: 'px-3 py-1.5 text-xs', md: 'px-4 py-2.5 text-sm', lg: 'px-5 py-3 text-sm' };
  const variants = {
    primary: 'text-white hover:brightness-95 active:brightness-90 shadow-soft',
    secondary: 'border hover:bg-[var(--primary-soft)]',
    ghost: 'hover:bg-black/5',
    danger: 'text-white hover:brightness-95',
  };
  const style = variant === 'primary' ? { backgroundColor: 'var(--primary)' }
    : variant === 'secondary' ? { borderColor: 'var(--border)', color: 'var(--text)', backgroundColor: 'var(--card)' }
    : variant === 'danger' ? { backgroundColor: 'var(--expense)' }
    : {};
  return (
    <button type={type} onClick={onClick} disabled={disabled} style={style} className={`${base} ${sizes[size]} ${variants[variant]} ${className}`}>
      {Icon && <Icon size={size === 'sm' ? 14 : 16} />}
      {children}
    </button>
  );
}

function Card({ children, className = '', padding = 'p-6' }) {
  return (
    <div className={`rounded-2xl shadow-soft ${padding} ${className}`} style={{ backgroundColor: 'var(--card)', border: '1px solid var(--border)' }}>
      {children}
    </div>
  );
}

function SectionTitle({ children, action }) {
  return (
    <div className="flex items-center justify-between mb-4">
      <h2 className="font-display text-lg font-semibold" style={{ color: 'var(--text)' }}>{children}</h2>
      {action}
    </div>
  );
}

function EmptyState({ icon: Icon = Search, title, description }) {
  return (
    <div className="flex flex-col items-center justify-center text-center py-14 px-4">
      <div className="rounded-2xl p-4 mb-4" style={{ backgroundColor: 'var(--primary-soft)' }}>
        <Icon size={26} color="var(--primary)" />
      </div>
      <p className="font-medium mb-1" style={{ color: 'var(--text)' }}>{title}</p>
      {description && <p className="text-sm max-w-xs" style={{ color: 'var(--text-soft)' }}>{description}</p>}
    </div>
  );
}

function Skeleton({ className }) {
  return <div className={`skeleton rounded-xl ${className}`} />;
}

function LoadingScreen() {
  return (
    <div className="flex-1 flex flex-col items-center justify-center gap-4" style={{ backgroundColor: 'var(--bg)' }}>
      <div className="relative w-16 h-16 flex items-center justify-center">
        <div className="absolute inset-0 rounded-full ring-pulse" style={{ border: '2px solid var(--primary)' }} />
        <RingProgress percent={70} size={56} strokeWidth={5}>
          <Sprout size={20} color="var(--primary)" />
        </RingProgress>
      </div>
      <p className="text-sm font-medium" style={{ color: 'var(--text-soft)' }}>Carregando o Cerne...</p>
    </div>
  );
}

/* ---------- Modal ---------- */

function Modal({ title, onClose, children, wide }) {
  useEffect(() => {
    function handleEsc(e) { if (e.key === 'Escape') onClose(); }
    window.addEventListener('keydown', handleEsc);
    return () => window.removeEventListener('keydown', handleEsc);
  }, [onClose]);
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ backgroundColor: 'rgba(35,35,35,0.35)' }} onClick={onClose}>
      <div
        className={`w-full ${wide ? 'max-w-2xl' : 'max-w-md'} max-h-[88vh] overflow-y-auto rounded-2xl shadow-soft-lg animate-fade-up`}
        style={{ backgroundColor: 'var(--card)' }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-6 py-4 sticky top-0 z-10" style={{ backgroundColor: 'var(--card)', borderBottom: '1px solid var(--border)' }}>
          <h3 className="font-display text-base font-semibold" style={{ color: 'var(--text)' }}>{title}</h3>
          <button onClick={onClose} className="p-1.5 rounded-lg hover:bg-black/5 focus-ring">
            <X size={18} color="var(--text-soft)" />
          </button>
        </div>
        <div className="p-6">{children}</div>
      </div>
    </div>
  );
}

function ConfirmModal({ title, description, onConfirm, onClose }) {
  return (
    <Modal title={title} onClose={onClose}>
      <p className="text-sm mb-6" style={{ color: 'var(--text-soft)' }}>{description}</p>
      <div className="flex justify-end gap-3">
        <Button variant="secondary" onClick={onClose}>Cancelar</Button>
        <Button variant="danger" onClick={() => { onConfirm(); onClose(); }}>Confirmar</Button>
      </div>
    </Modal>
  );
}

/* ---------- Toasts ---------- */

function ToastContainer({ toasts }) {
  return (
    <div className="fixed bottom-6 right-6 z-[60] flex flex-col gap-2 no-print">
      {toasts.map((t) => (
        <div key={t.id} className="animate-toast flex items-center gap-2.5 px-4 py-3 rounded-xl shadow-soft-lg text-sm font-medium" style={{ backgroundColor: t.type === 'error' ? 'var(--expense)' : 'var(--primary-dark)', color: '#fff' }}>
          {t.type === 'error' ? <AlertCircle size={16} /> : <Check size={16} />}
          {t.message}
        </div>
      ))}
    </div>
  );
}

/* ---------- Formulário: campo de moeda ---------- */

function CurrencyInput({ value, onChange, placeholder = 'R$ 0,00' }) {
  const [display, setDisplay] = useState(value ? String(value) : '');
  useEffect(() => { setDisplay(value ? formatBRL(value).replace('R$', '').trim() : ''); }, [value]);
  function handleChange(e) {
    const digits = e.target.value.replace(/\D/g, '');
    const num = digits ? parseInt(digits, 10) / 100 : 0;
    setDisplay(digits ? formatBRL(num).replace('R$', '').trim() : '');
    onChange(num);
  }
  return (
    <div className="relative">
      <span className="absolute left-3.5 top-1/2 -translate-y-1/2 text-sm" style={{ color: 'var(--text-soft)' }}>R$</span>
      <input
        value={display} onChange={handleChange} placeholder={placeholder} inputMode="numeric"
        className="w-full pl-9 pr-3.5 py-2.5 rounded-xl text-sm focus-ring tabular-nums"
        style={{ border: '1px solid var(--border)', color: 'var(--text)' }}
      />
    </div>
  );
}

function FieldLabel({ children, error }) {
  return (
    <label className="block text-xs font-medium mb-1.5" style={{ color: error ? 'var(--expense)' : 'var(--text-soft)' }}>
      {children}
    </label>
  );
}

const inputClass = 'w-full px-3.5 py-2.5 rounded-xl text-sm focus-ring';
const inputStyle = { border: '1px solid var(--border)', color: 'var(--text)', backgroundColor: 'var(--card)' };

/* ============================================================
   SIDEBAR / HEADER / BANNER
   ============================================================ */

const NAV_ITEMS = [
  { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
  { id: 'transacoes', label: 'Transações', icon: ArrowLeftRight },
  { id: 'receitas', label: 'Receitas', icon: TrendingUp },
  { id: 'despesas', label: 'Despesas', icon: TrendingDown },
  { id: 'contas', label: 'Contas Bancárias', icon: Landmark },
  { id: 'cartoes', label: 'Cartões', icon: CreditCard },
  { id: 'investimentos', label: 'Investimentos', icon: PieChartIcon },
  { id: 'metas', label: 'Metas', icon: Target },
  { id: 'recorrentes', label: 'Despesas Recorrentes', icon: RefreshCw },
  { id: 'relatorios', label: 'Relatórios', icon: BarChart3 },
  { id: 'configuracoes', label: 'Configurações', icon: Settings },
];

function Sidebar({ activePage, setActivePage, sidebarOpen, setSidebarOpen, onNewTransaction, userName }) {
  return (
    <>
      {sidebarOpen && (
        <div className="fixed inset-0 z-40 lg:hidden" style={{ backgroundColor: 'rgba(35,35,35,0.4)' }} onClick={() => setSidebarOpen(false)} />
      )}
      <aside
        className={`fixed lg:static inset-y-0 left-0 z-50 w-72 shrink-0 flex flex-col transition-transform duration-300 ${sidebarOpen ? 'translate-x-0' : '-translate-x-full lg:translate-x-0'}`}
        style={{ backgroundColor: 'var(--card)', borderRight: '1px solid var(--border)' }}
      >
        <div className="flex items-center gap-3 px-6 py-6">
          <div className="rounded-xl p-2" style={{ backgroundColor: 'var(--primary)' }}>
            <Sprout size={20} color="#fff" />
          </div>
          <div>
            <p className="font-display text-lg font-bold leading-none" style={{ color: 'var(--text)' }}>Cerne</p>
            <p className="text-[11px] mt-0.5" style={{ color: 'var(--text-soft)' }}>Gestão financeira</p>
          </div>
          <button onClick={() => setSidebarOpen(false)} className="ml-auto lg:hidden p-1.5 rounded-lg hover:bg-black/5">
            <X size={18} />
          </button>
        </div>

        <div className="mx-6 mb-4 p-3 rounded-xl flex items-center gap-3" style={{ backgroundColor: 'var(--primary-soft)' }}>
          <div className="w-9 h-9 rounded-full flex items-center justify-center font-display font-semibold text-sm shrink-0" style={{ backgroundColor: 'var(--primary)', color: '#fff' }}>
            {userName.split(' ').map((p) => p[0]).slice(0, 2).join('').toUpperCase()}
          </div>
          <div className="min-w-0">
            <p className="text-sm font-medium truncate" style={{ color: 'var(--text)' }}>{userName}</p>
            <p className="text-[11px]" style={{ color: 'var(--text-soft)' }}>Conta Premium</p>
          </div>
        </div>

        <nav className="flex-1 overflow-y-auto px-4 space-y-1">
          {NAV_ITEMS.map((item) => {
            const active = activePage === item.id;
            return (
              <button
                key={item.id}
                onClick={() => { setActivePage(item.id); setSidebarOpen(false); }}
                className="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-medium transition-colors focus-ring"
                style={{
                  backgroundColor: active ? 'var(--primary-soft)' : 'transparent',
                  color: active ? 'var(--primary-dark)' : 'var(--text-soft)',
                }}
              >
                <item.icon size={17} strokeWidth={2.2} />
                <span className="truncate">{item.label}</span>
              </button>
            );
          })}
        </nav>

        <div className="p-4">
          <Button variant="primary" className="w-full" icon={Plus} onClick={onNewTransaction}>Novo lançamento</Button>
        </div>
      </aside>
    </>
  );
}

function PeriodSelector({ period, setPeriod, customRange, setCustomRange }) {
  const [open, setOpen] = useState(false);
  const labels = { mes: 'Este mês', trimestre: 'Trimestre', ano: 'Ano', personalizado: 'Personalizado' };
  return (
    <div className="relative">
      <button onClick={() => setOpen(!open)} className="flex items-center gap-2 px-3.5 py-2 rounded-xl text-sm font-medium focus-ring" style={{ border: '1px solid var(--border)', color: 'var(--text)' }}>
        <CalendarIcon size={15} color="var(--text-soft)" />
        {labels[period]}
        <ChevronDown size={14} color="var(--text-soft)" />
      </button>
      {open && (
        <div className="absolute right-0 mt-2 w-64 rounded-xl shadow-soft-lg p-2 z-20" style={{ backgroundColor: 'var(--card)', border: '1px solid var(--border)' }}>
          {Object.entries(labels).map(([key, label]) => (
            <button
              key={key}
              onClick={() => { setPeriod(key); if (key !== 'personalizado') setOpen(false); }}
              className="w-full text-left px-3 py-2 rounded-lg text-sm hover:bg-[var(--primary-soft)]"
              style={{ color: period === key ? 'var(--primary-dark)' : 'var(--text)', fontWeight: period === key ? 600 : 400 }}
            >
              {label}
            </button>
          ))}
          {period === 'personalizado' && (
            <div className="p-2 space-y-2 border-t mt-1" style={{ borderColor: 'var(--border)' }}>
              <input type="date" value={customRange.start} onChange={(e) => setCustomRange((r) => ({ ...r, start: e.target.value }))} className={inputClass} style={inputStyle} />
              <input type="date" value={customRange.end} onChange={(e) => setCustomRange((r) => ({ ...r, end: e.target.value }))} className={inputClass} style={inputStyle} />
              <Button size="sm" className="w-full" onClick={() => setOpen(false)}>Aplicar</Button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function Header({ userName, period, setPeriod, customRange, setCustomRange, search, setSearch, setSidebarOpen }) {
  const [notifOpen, setNotifOpen] = useState(false);
  const today = new Date().toLocaleDateString('pt-BR', { weekday: 'long', day: 'numeric', month: 'long' });
  const notifications = [
    { text: 'A fatura do cartão Nubank fecha em 3 dias.', time: '2h atrás' },
    { text: 'Você atingiu 91% da meta de reserva de emergência.', time: '1 dia atrás' },
    { text: 'Conta de água vence amanhã.', time: '1 dia atrás' },
  ];
  return (
    <header className="sticky top-0 z-30 flex items-center gap-3 px-4 md:px-6 lg:px-8 py-4 no-print" style={{ backgroundColor: 'var(--bg)', borderBottom: '1px solid var(--border)' }}>
      <button onClick={() => setSidebarOpen(true)} className="lg:hidden p-2 rounded-xl hover:bg-black/5">
        <Menu size={20} color="var(--text)" />
      </button>
      <div className="hidden md:block min-w-0">
        <p className="font-display text-base font-semibold truncate" style={{ color: 'var(--text)' }}>{getGreeting()}, {userName.split(' ')[0]}</p>
        <p className="text-xs capitalize" style={{ color: 'var(--text-soft)' }}>{today}</p>
      </div>
      <div className="flex-1 flex justify-center px-2">
        <div className="relative w-full max-w-sm">
          <Search size={15} className="absolute left-3.5 top-1/2 -translate-y-1/2" color="var(--text-soft)" />
          <input
            value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Buscar lançamentos, contas, cartões..."
            className="w-full pl-9 pr-3 py-2 rounded-xl text-sm focus-ring" style={{ border: '1px solid var(--border)', backgroundColor: 'var(--card)', color: 'var(--text)' }}
          />
        </div>
      </div>
      <div className="hidden sm:block">
        <PeriodSelector period={period} setPeriod={setPeriod} customRange={customRange} setCustomRange={setCustomRange} />
      </div>
      <div className="relative">
        <button onClick={() => setNotifOpen(!notifOpen)} className="relative p-2.5 rounded-xl hover:bg-black/5 focus-ring">
          <Bell size={18} color="var(--text-soft)" />
          <span className="absolute top-1.5 right-1.5 w-2 h-2 rounded-full" style={{ backgroundColor: 'var(--alert)' }} />
        </button>
        {notifOpen && (
          <div className="absolute right-0 mt-2 w-72 rounded-xl shadow-soft-lg p-2 z-20" style={{ backgroundColor: 'var(--card)', border: '1px solid var(--border)' }}>
            {notifications.map((n, i) => (
              <div key={i} className="px-3 py-2.5 rounded-lg hover:bg-black/5">
                <p className="text-sm" style={{ color: 'var(--text)' }}>{n.text}</p>
                <p className="text-[11px] mt-0.5" style={{ color: 'var(--text-soft)' }}>{n.time}</p>
              </div>
            ))}
          </div>
        )}
      </div>
      <div className="w-9 h-9 rounded-full flex items-center justify-center font-display font-semibold text-sm shrink-0" style={{ backgroundColor: 'var(--primary)', color: '#fff' }}>
        {userName.split(' ').map((p) => p[0]).slice(0, 2).join('').toUpperCase()}
      </div>
    </header>
  );
}

function Banner({ insight, onDismiss }) {
  if (!insight) return null;
  const Icon = insight.icon;
  return (
    <div className="mx-4 md:mx-6 lg:mx-8 mt-4 mb-2 no-print">
      <div className="flex items-center gap-3 px-4 py-3 rounded-xl" style={{ backgroundColor: 'var(--primary-soft)' }}>
        <Icon size={16} color="var(--primary-dark)" className="shrink-0" />
        <p className="text-sm flex-1" style={{ color: 'var(--primary-dark)' }}>{insight.text}</p>
        <button onClick={onDismiss} className="p-1 rounded-lg hover:bg-black/5 shrink-0"><X size={14} color="var(--primary-dark)" /></button>
      </div>
    </div>
  );
}

/* ============================================================
   DASHBOARD — cartões de indicadores
   ============================================================ */

function computeKPIs(period, customRange, mh, transactions, accounts, goals, caixinhas, monthlySavingsTarget) {
  let receitas = 0, despesas = 0;
  if (period === 'personalizado' && customRange.start && customRange.end) {
    const filtered = transactions.filter((t) => t.status === 'Pago' && t.date >= customRange.start && t.date <= customRange.end);
    receitas = filtered.filter((t) => t.type === 'receita').reduce((s, t) => s + t.amount, 0);
    despesas = filtered.filter((t) => t.type === 'despesa').reduce((s, t) => s + t.amount, 0);
  } else {
    const monthsToTake = period === 'mes' ? 1 : period === 'trimestre' ? 3 : 12;
    const slice = mh.slice(-monthsToTake);
    receitas = slice.reduce((s, m) => s + m.receitas, 0);
    despesas = slice.reduce((s, m) => s + m.despesas, 0);
  }
  const saldoDisponivel = accounts.reduce((s, a) => s + a.balance, 0);
  const patrimonio = mh.length ? mh[mh.length - 1].patrimonio : saldoDisponivel;
  const economiaAcumulada = caixinhas.reduce((s, c) => s + c.balance, 0) + goals.reduce((s, g) => s + g.current, 0);
  const metaMensal = { current: mh.length ? thisMonthSaved(mh) : 0, target: monthlySavingsTarget > 0 ? monthlySavingsTarget : 0 };

  const last = mh[mh.length - 1];
  const prev = mh.length > 1 ? mh[mh.length - 2] : null;
  const pctChange = (curr, before) => (before ? ((curr - before) / Math.abs(before)) * 100 : null);
  const growth = prev
    ? { receitas: pctChange(last.receitas, prev.receitas), despesas: pctChange(last.despesas, prev.despesas), patrimonio: pctChange(last.patrimonio, prev.patrimonio) }
    : { receitas: null, despesas: null, patrimonio: null };

  return { receitas, despesas, saldo: receitas - despesas, saldoDisponivel, patrimonio, economiaAcumulada, metaMensal, growth };
}
function thisMonthSaved(mh) {
  const m = mh[mh.length - 1];
  return m.receitas - m.despesas;
}

function StatCard({ title, value, description, icon: Icon, color, soft, growth, isRing, ringPercent }) {
  return (
    <Card padding="p-5" className="animate-fade-up">
      <div className="flex items-start justify-between mb-3">
        <IconCircle icon={Icon} color={color} soft={soft} />
        {growth != null && (
          <span className="flex items-center gap-0.5 text-xs font-semibold" style={{ color: growth >= 0 ? 'var(--income)' : 'var(--expense)' }}>
            {growth >= 0 ? <ArrowUpRight size={14} /> : <ArrowDownRight size={14} />}
            {Math.abs(growth).toFixed(0)}%
          </span>
        )}
      </div>
      <p className="text-xs font-medium mb-1" style={{ color: 'var(--text-soft)' }}>{title}</p>
      {isRing ? (
        <div className="flex items-center gap-3">
          <RingProgress percent={ringPercent} size={44} strokeWidth={5} color={color}>
            <span className="font-display text-xs font-bold" style={{ color }}>{ringPercent.toFixed(0)}%</span>
          </RingProgress>
          <p className="font-display text-lg font-bold tabular-nums" style={{ color: 'var(--text)' }}>{value}</p>
        </div>
      ) : (
        <p className="font-display text-2xl font-bold tabular-nums mb-1" style={{ color: 'var(--text)' }}>{value}</p>
      )}
      {description && <p className="text-xs mt-1" style={{ color: 'var(--text-soft)' }}>{description}</p>}
    </Card>
  );
}

function KPIRow({ kpis }) {
  const metaPercent = kpis.metaMensal.target > 0 ? Math.min(100, (kpis.metaMensal.current / kpis.metaMensal.target) * 100) : 0;
  const cards = [
    { title: 'Saldo disponível', value: formatBRL(kpis.saldoDisponivel), description: 'Soma de todas as contas', icon: Wallet, color: 'var(--primary)', soft: 'var(--primary-soft)' },
    { title: 'Receitas do período', value: formatBRL(kpis.receitas), description: 'Entradas no período selecionado', icon: TrendingUp, color: 'var(--income)', soft: 'var(--income-soft)', growth: kpis.growth.receitas },
    { title: 'Despesas do período', value: formatBRL(kpis.despesas), description: 'Saídas no período selecionado', icon: TrendingDown, color: 'var(--expense)', soft: 'var(--expense-soft)', growth: kpis.growth.despesas != null ? -kpis.growth.despesas : null },
    { title: 'Economia acumulada', value: formatBRL(kpis.economiaAcumulada), description: 'Em caixinhas + metas', icon: PiggyBank, color: 'var(--goals)', soft: 'var(--goals-soft)' },
    { title: 'Meta mensal', value: formatBRL(kpis.metaMensal.current), description: kpis.metaMensal.target > 0 ? `de ${formatBRL(kpis.metaMensal.target)} planejados` : 'defina uma meta em Configurações', icon: Target, color: 'var(--alert)', soft: 'var(--alert-soft)', isRing: true, ringPercent: metaPercent },
    { title: 'Patrimônio total', value: formatBRL(kpis.patrimonio), description: 'Contas + investimentos', icon: Landmark, color: 'var(--invest)', soft: 'var(--invest-soft)', growth: kpis.growth.patrimonio },
  ];
  return (
    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
      {cards.map((c, i) => <StatCard key={i} {...c} />)}
    </div>
  );
}

/* ---------- Gráficos do dashboard ---------- */

function EvolutionChart({ data }) {
  const [visible, setVisible] = useState({ receitas: true, despesas: true, saldo: true, patrimonio: true });
  const chartData = data.map((m) => ({ ...m, saldo: m.receitas - m.despesas }));
  function toggle(key) { setVisible((v) => ({ ...v, [key]: !v[key] })); }
  const series = [
    { key: 'receitas', name: 'Receitas', color: '#5A8F5A' },
    { key: 'despesas', name: 'Despesas', color: '#B66B6B' },
    { key: 'saldo', name: 'Saldo líquido', color: '#5D7052' },
    { key: 'patrimonio', name: 'Patrimônio', color: '#6B7FAF' },
  ];
  return (
    <Card className="animate-fade-up">
      <SectionTitle>Evolução financeira (12 meses)</SectionTitle>
      <div className="flex flex-wrap gap-2 mb-4">
        {series.map((s) => (
          <button
            key={s.key} onClick={() => toggle(s.key)}
            className="flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium border"
            style={{ borderColor: visible[s.key] ? s.color : 'var(--border)', color: visible[s.key] ? s.color : 'var(--text-soft)', opacity: visible[s.key] ? 1 : 0.5 }}
          >
            <span className="w-2 h-2 rounded-full" style={{ backgroundColor: s.color }} />
            {s.name}
          </button>
        ))}
      </div>
      <div style={{ width: '100%', height: 280 }}>
        <ResponsiveContainer>
          <LineChart data={chartData} margin={{ left: -18, right: 8, top: 6 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" vertical={false} />
            <XAxis dataKey="month" tick={{ fontSize: 11, fill: '#6D6D6D' }} axisLine={{ stroke: '#E7E4DE' }} tickLine={false} />
            <YAxis tick={{ fontSize: 11, fill: '#6D6D6D' }} axisLine={false} tickLine={false} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} />
            <Tooltip
              formatter={(value, name) => [formatBRL(value), name]}
              contentStyle={{ borderRadius: 12, border: '1px solid #E7E4DE', fontSize: 12, fontFamily: 'Plus Jakarta Sans' }}
            />
            {series.map((s) => visible[s.key] && (
              <Line key={s.key} type="monotone" dataKey={s.key} name={s.name} stroke={s.color} strokeWidth={2.5} dot={false} activeDot={{ r: 5 }} animationDuration={800} />
            ))}
          </LineChart>
        </ResponsiveContainer>
      </div>
    </Card>
  );
}

function CategoryDonut({ data, title = 'Gastos por categoria', subtitle = 'neste mês' }) {
  const entries = Object.entries(data).sort((a, b) => b[1] - a[1]);
  const total = entries.reduce((s, [, v]) => s + v, 0);
  const chartData = entries.map(([name, value]) => ({ name, value, color: (CATEGORIES[name] || CATEGORIES['Outros']).color }));
  if (entries.length === 0) {
    return (
      <Card className="animate-fade-up">
        <SectionTitle>{title} <span className="text-xs font-normal" style={{ color: 'var(--text-soft)' }}>{subtitle}</span></SectionTitle>
        <EmptyState icon={PieChartIcon} title="Nenhum gasto registrado" description="Assim que você lançar despesas pagas neste período, elas aparecem aqui." />
      </Card>
    );
  }
  return (
    <Card className="animate-fade-up">
      <SectionTitle>{title} <span className="text-xs font-normal" style={{ color: 'var(--text-soft)' }}>{subtitle}</span></SectionTitle>
      <div className="flex flex-col sm:flex-row items-center gap-4">
        <div style={{ width: 170, height: 170 }} className="shrink-0">
          <ResponsiveContainer>
            <PieChart>
              <Pie data={chartData} dataKey="value" nameKey="name" innerRadius={52} outerRadius={78} paddingAngle={2} animationDuration={700}>
                {chartData.map((entry, i) => <Cell key={i} fill={entry.color} stroke="none" />)}
              </Pie>
              <Tooltip formatter={(value) => formatBRL(value)} contentStyle={{ borderRadius: 12, border: '1px solid #E7E4DE', fontSize: 12 }} />
            </PieChart>
          </ResponsiveContainer>
        </div>
        <div className="w-full space-y-2">
          {chartData.map((entry, i) => (
            <div key={i} className="flex items-center gap-2 text-sm">
              <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: entry.color }} />
              <span className="flex-1 truncate" style={{ color: 'var(--text)' }}>{entry.name}</span>
              <span className="tabular-nums font-medium" style={{ color: 'var(--text-soft)' }}>{((entry.value / total) * 100).toFixed(0)}%</span>
              <span className="tabular-nums w-20 text-right shrink-0" style={{ color: 'var(--text)' }}>{formatBRL(entry.value)}</span>
            </div>
          ))}
        </div>
      </div>
    </Card>
  );
}

/* ---------- Tabela de transações (compartilhada) ---------- */

function TransactionRow({ tx, accounts, onEdit, onDelete }) {
  const accName = accounts.find((a) => a.id === tx.account)?.bank || 'Conta removida';
  return (
    <tr className="text-sm hover:bg-black/[0.02]">
      <td className="py-3 pr-3 whitespace-nowrap" style={{ color: 'var(--text-soft)' }}>{formatDate(tx.date)}</td>
      <td className="py-3 pr-3" style={{ color: 'var(--text)' }}>{tx.description}</td>
      <td className="py-3 pr-3"><CategoryBadge category={tx.category} /></td>
      <td className="py-3 pr-3 whitespace-nowrap" style={{ color: 'var(--text-soft)' }}>{accName}</td>
      <td className="py-3 pr-3 whitespace-nowrap" style={{ color: 'var(--text-soft)' }}>{tx.paymentMethod}</td>
      <td className="py-3 pr-3 tabular-nums font-medium whitespace-nowrap" style={{ color: tx.type === 'receita' ? 'var(--income)' : 'var(--expense)' }}>
        {tx.type === 'receita' ? '+' : '-'} {formatBRL(tx.amount)}
      </td>
      <td className="py-3 pr-3"><StatusBadge status={tx.status} /></td>
      <td className="py-3 pr-1">
        <div className="flex items-center gap-1">
          <button onClick={() => onEdit(tx)} className="p-1.5 rounded-lg hover:bg-black/5"><Pencil size={14} color="var(--text-soft)" /></button>
          <button onClick={() => onDelete(tx)} className="p-1.5 rounded-lg hover:bg-black/5"><Trash2 size={14} color="var(--expense)" /></button>
        </div>
      </td>
    </tr>
  );
}

const TABLE_HEAD = ['Data', 'Descrição', 'Categoria', 'Conta', 'Forma de pagamento', 'Valor', 'Status', ''];

function RecentTransactions({ transactions, accounts, onEdit, onDelete, onSeeAll }) {
  const recent = [...transactions].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 7);
  return (
    <Card className="animate-fade-up" padding="p-5 sm:p-6">
      <SectionTitle action={<button onClick={onSeeAll} className="text-xs font-medium flex items-center gap-1" style={{ color: 'var(--primary)' }}>Ver todas <ArrowRight size={12} /></button>}>
        Transações recentes
      </SectionTitle>
      <div className="overflow-x-auto -mx-1">
        <table className="w-full min-w-[640px]">
          <thead>
            <tr className="text-left text-xs" style={{ color: 'var(--text-soft)' }}>
              {TABLE_HEAD.map((h, i) => <th key={i} className="pb-2 pr-3 font-medium">{h}</th>)}
            </tr>
          </thead>
          <tbody style={{ borderTop: '1px solid var(--border)' }}>
            {recent.map((tx) => <TransactionRow key={tx.id} tx={tx} accounts={accounts} onEdit={onEdit} onDelete={onDelete} />)}
          </tbody>
        </table>
      </div>
    </Card>
  );
}

function MonthSummaryPanel({ transactions, kpis }) {
  const today = new Date();
  const thisMonthTx = transactions.filter((t) => isSameMonth(t.date, today.getFullYear(), today.getMonth()));
  const receitasList = thisMonthTx.filter((t) => t.type === 'receita');
  const despesasList = thisMonthTx.filter((t) => t.type === 'despesa');
  const maiorReceita = receitasList.sort((a, b) => b.amount - a.amount)[0];
  const maiorDespesa = despesasList.sort((a, b) => b.amount - a.amount)[0];
  const diasNoMes = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  const mediaDiaria = kpis.despesas / diasNoMes;
  const metaPercent = kpis.metaMensal.target > 0 ? Math.min(100, (kpis.metaMensal.current / kpis.metaMensal.target) * 100) : 0;
  const rows = [
    { label: 'Maior receita', value: maiorReceita ? formatBRL(maiorReceita.amount) : '—', sub: maiorReceita?.description },
    { label: 'Maior despesa', value: maiorDespesa ? formatBRL(maiorDespesa.amount) : '—', sub: maiorDespesa?.description },
    { label: 'Média diária de gastos', value: formatBRL(mediaDiaria) },
    { label: 'Economia acumulada', value: formatBRL(kpis.economiaAcumulada) },
    { label: 'Saldo previsto (fechamento)', value: formatBRL(kpis.saldoDisponivel + kpis.saldo) },
  ];
  return (
    <Card className="animate-fade-up h-fit">
      <SectionTitle>Resumo do mês</SectionTitle>
      <div className="space-y-4">
        {rows.map((r, i) => (
          <div key={i} className="flex items-center justify-between text-sm">
            <div className="min-w-0">
              <p style={{ color: 'var(--text-soft)' }}>{r.label}</p>
              {r.sub && <p className="text-xs truncate" style={{ color: 'var(--text-soft)', opacity: 0.7 }}>{r.sub}</p>}
            </div>
            <span className="tabular-nums font-medium shrink-0 ml-2" style={{ color: 'var(--text)' }}>{r.value}</span>
          </div>
        ))}
        <div className="pt-3" style={{ borderTop: '1px solid var(--border)' }}>
          <div className="flex items-center justify-between text-sm mb-2">
            <span style={{ color: 'var(--text-soft)' }}>Meta do mês atingida</span>
            <span className="font-semibold" style={{ color: 'var(--primary)' }}>{metaPercent.toFixed(0)}%</span>
          </div>
          <ProgressBar percent={metaPercent} color="var(--primary)" />
        </div>
      </div>
    </Card>
  );
}

/* ---------- Metas (preview + cards) ---------- */

function GoalCard({ goal, onAddFunds, onDelete, compact }) {
  const Icon = GOAL_ICONS[goal.icon] || Target;
  const percent = Math.min(100, (goal.current / goal.target) * 100);
  const [adding, setAdding] = useState(false);
  const [amount, setAmount] = useState(0);
  const [showHistory, setShowHistory] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const deadlineLabel = new Date(goal.deadline + 'T00:00:00').toLocaleDateString('pt-BR', { month: 'short', year: 'numeric' });
  const estimate = useMemo(() => estimateGoalCompletion(goal), [goal.current, goal.target, goal.history]);
  const history = goal.history || [];
  return (
    <Card className="animate-fade-up" padding="p-5">
      <div className="flex items-start gap-3 mb-4">
        <IconCircle icon={Icon} color="var(--goals)" soft="var(--goals-soft)" />
        <div className="min-w-0 flex-1">
          <p className="font-medium text-sm truncate" style={{ color: 'var(--text)' }}>{goal.name}</p>
          <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Previsão: {deadlineLabel}</p>
        </div>
        {!compact && (
          <button onClick={() => setConfirmDelete(true)} className="p-1 rounded-lg hover:bg-black/5 shrink-0"><Trash2 size={14} color="var(--text-soft)" /></button>
        )}
      </div>
      <div className="flex items-end justify-between mb-2">
        <span className="font-display text-xl font-bold tabular-nums" style={{ color: 'var(--text)' }}>{formatBRL(goal.current)}</span>
        <span className="text-xs tabular-nums" style={{ color: 'var(--text-soft)' }}>de {formatBRL(goal.target)}</span>
      </div>
      <ProgressBar percent={percent} color="var(--goals)" />
      <div className="flex items-center justify-between mt-2">
        <span className="text-xs font-semibold" style={{ color: 'var(--goals)' }}>{percent.toFixed(0)}% concluído</span>
        {!compact && !adding && (
          <button onClick={() => setAdding(true)} className="text-xs font-medium flex items-center gap-1" style={{ color: 'var(--primary)' }}>
            <Plus size={12} /> Adicionar valor
          </button>
        )}
      </div>
      {adding && (
        <div className="flex items-center gap-2 mt-3">
          <CurrencyInput value={amount} onChange={setAmount} />
          <Button size="sm" onClick={() => { if (amount > 0) { onAddFunds(goal, amount); } setAdding(false); setAmount(0); }}>OK</Button>
          <button onClick={() => { setAdding(false); setAmount(0); }} className="p-2"><X size={14} color="var(--text-soft)" /></button>
        </div>
      )}
      {!compact && (
        <div className="mt-3 pt-3" style={{ borderTop: '1px solid var(--border)' }}>
          {estimate.status === 'ok' && (
            <p className="text-xs flex items-start gap-1.5" style={{ color: 'var(--text-soft)' }}>
              <TrendingUp size={13} className="mt-0.5 shrink-0" color="var(--goals)" />
              No ritmo atual (~{formatBRL(estimate.avgMonthly)}/mês), a meta deve ser concluída em <strong style={{ color: 'var(--text)' }}>{estimate.etaLabel}</strong>.
            </p>
          )}
          {estimate.status === 'done' && (
            <p className="text-xs flex items-center gap-1.5" style={{ color: 'var(--income)' }}><Check size={13} /> Meta já atingida!</p>
          )}
          {estimate.status === 'no-data' && (
            <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Adicione valores para estimar quando a meta será concluída.</p>
          )}
          {history.length > 0 && (
            <>
              <button onClick={() => setShowHistory((s) => !s)} className="text-xs font-medium mt-2 flex items-center gap-1" style={{ color: 'var(--primary)' }}>
                {showHistory ? <ChevronUp size={12} /> : <ChevronDown size={12} />} Histórico de aportes ({history.length})
              </button>
              {showHistory && (
                <div className="space-y-1 mt-2 max-h-32 overflow-y-auto pr-1">
                  {[...history].reverse().map((h, i) => (
                    <div key={i} className="flex items-center justify-between text-xs px-2 py-1.5 rounded-lg" style={{ backgroundColor: 'var(--bg)' }}>
                      <span style={{ color: 'var(--text-soft)' }}>{formatDate(h.date)}</span>
                      <span className="tabular-nums font-medium" style={{ color: 'var(--goals)' }}>+{formatBRL(h.amount)}</span>
                    </div>
                  ))}
                </div>
              )}
            </>
          )}
        </div>
      )}
      {confirmDelete && (
        <ConfirmModal
          title="Excluir meta"
          description={`Tem certeza que deseja excluir a meta "${goal.name}"? O histórico de aportes será perdido. Essa ação não pode ser desfeita.`}
          onConfirm={() => { onDelete(goal); setConfirmDelete(false); }}
          onClose={() => setConfirmDelete(false)}
        />
      )}
    </Card>
  );
}

function GoalForm({ onSave, onClose }) {
  const [form, setForm] = useState({ name: '', target: 0, current: 0, deadline: '', icon: 'Wallet' });
  const [errors, setErrors] = useState({});
  function validate() {
    const e = {};
    if (!form.name.trim()) e.name = 'Informe um nome';
    if (!form.target || form.target <= 0) e.target = 'Informe um valor objetivo';
    if (!form.deadline) e.deadline = 'Informe uma data prevista';
    setErrors(e);
    return Object.keys(e).length === 0;
  }
  return (
    <Modal title="Nova meta" onClose={onClose}>
      <div className="space-y-4">
        <div>
          <FieldLabel error={errors.name}>Nome da meta</FieldLabel>
          <input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Viagem para a praia" />
          {errors.name && <p className="text-xs mt-1" style={{ color: 'var(--expense)' }}>{errors.name}</p>}
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <FieldLabel error={errors.target}>Valor objetivo</FieldLabel>
            <CurrencyInput value={form.target} onChange={(v) => setForm({ ...form, target: v })} />
            {errors.target && <p className="text-xs mt-1" style={{ color: 'var(--expense)' }}>{errors.target}</p>}
          </div>
          <div>
            <FieldLabel>Valor já guardado</FieldLabel>
            <CurrencyInput value={form.current} onChange={(v) => setForm({ ...form, current: v })} />
          </div>
        </div>
        <div>
          <FieldLabel error={errors.deadline}>Previsão de conclusão</FieldLabel>
          <input type="date" value={form.deadline} onChange={(e) => setForm({ ...form, deadline: e.target.value })} className={inputClass} style={inputStyle} />
          {errors.deadline && <p className="text-xs mt-1" style={{ color: 'var(--expense)' }}>{errors.deadline}</p>}
        </div>
        <div>
          <FieldLabel>Ícone</FieldLabel>
          <div className="flex gap-2">
            {Object.entries(GOAL_ICONS).map(([key, Icon]) => (
              <button key={key} onClick={() => setForm({ ...form, icon: key })} className="p-2.5 rounded-xl" style={{ backgroundColor: form.icon === key ? 'var(--primary-soft)' : 'transparent', border: '1px solid var(--border)' }}>
                <Icon size={16} color={form.icon === key ? 'var(--primary)' : 'var(--text-soft)'} />
              </button>
            ))}
          </div>
        </div>
        <div className="flex justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={() => { if (validate()) onSave(form); }}>Criar meta</Button>
        </div>
      </div>
    </Modal>
  );
}

function GoalsSection({ goals, onAddFunds, onDelete, onSeeAll, compact }) {
  return (
    <Card className="animate-fade-up" padding="p-5 sm:p-6">
      <SectionTitle action={onSeeAll && <button onClick={onSeeAll} className="text-xs font-medium flex items-center gap-1" style={{ color: 'var(--primary)' }}>Ver todas <ArrowRight size={12} /></button>}>
        Metas financeiras
      </SectionTitle>
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-4">
        {goals.map((g) => <GoalCard key={g.id} goal={g} onAddFunds={onAddFunds} onDelete={onDelete} compact={compact} />)}
      </div>
    </Card>
  );
}

/* ---------- Calendário financeiro ---------- */

function FinancialCalendar({ cards, transactions }) {
  const [viewDate, setViewDate] = useState(() => { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), 1); });
  const year = viewDate.getFullYear(), month = viewDate.getMonth();
  const cells = getMonthGrid(year, month);
  const weekDays = ['D', 'S', 'T', 'Q', 'Q', 'S', 'S'];

  const events = [];
  cards.forEach((c) => {
    events.push({ day: c.closingDay, label: `Fechamento fatura ${c.bank}`, type: 'fatura', amount: c.invoice });
    const adjustedDue = getAdjustedDueDate(year, month, c.dueDay);
    if (adjustedDue.getMonth() === month) {
      const wasAdjusted = adjustedDue.getDate() !== Math.min(c.dueDay, new Date(year, month + 1, 0).getDate());
      events.push({ day: adjustedDue.getDate(), label: `Vencimento fatura ${c.bank}${wasAdjusted ? ' (antecipado p/ dia útil)' : ''}`, type: 'vencimento', amount: c.invoice });
    }
  });
  transactions.filter((t) => t.status !== 'Pago' && isSameMonth(t.date, year, month)).forEach((t) => {
    const d = new Date(t.date + 'T00:00:00');
    events.push({ day: d.getDate(), label: t.description, type: t.type === 'receita' ? 'recebimento' : 'boleto', amount: t.amount });
  });
  events.sort((a, b) => a.day - b.day);
  const eventDays = new Set(events.map((e) => e.day));

  return (
    <Card className="animate-fade-up">
      <SectionTitle
        action={
          <div className="flex items-center gap-1">
            <button onClick={() => setViewDate(new Date(year, month - 1, 1))} className="p-1.5 rounded-lg hover:bg-black/5"><ChevronLeft size={16} /></button>
            <span className="text-xs font-medium capitalize w-24 text-center" style={{ color: 'var(--text-soft)' }}>{viewDate.toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' })}</span>
            <button onClick={() => setViewDate(new Date(year, month + 1, 1))} className="p-1.5 rounded-lg hover:bg-black/5"><ChevronRight size={16} /></button>
          </div>
        }
      >
        Calendário financeiro
      </SectionTitle>
      <div className="grid grid-cols-7 gap-1 mb-3">
        {weekDays.map((d, i) => <div key={i} className="text-center text-[11px] font-medium py-1" style={{ color: 'var(--text-soft)' }}>{d}</div>)}
        {cells.map((c) => (
          <div key={c.key} className="aspect-square flex items-center justify-center relative">
            {c.day && (
              <span className="text-xs w-7 h-7 flex items-center justify-center rounded-full" style={{ color: 'var(--text)' }}>
                {c.day}
                {eventDays.has(c.day) && <span className="absolute bottom-0.5 w-1.5 h-1.5 rounded-full" style={{ backgroundColor: 'var(--alert)' }} />}
              </span>
            )}
          </div>
        ))}
      </div>
      <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
        {events.map((e, i) => (
          <div key={i} className="flex items-center justify-between text-xs px-3 py-2 rounded-lg" style={{ backgroundColor: 'var(--bg)' }}>
            <div className="flex items-center gap-2 min-w-0">
              <span className="font-display font-semibold w-6 text-center shrink-0" style={{ color: 'var(--primary)' }}>{e.day}</span>
              <span className="truncate" style={{ color: 'var(--text)' }}>{e.label}</span>
            </div>
            <span className="tabular-nums font-medium shrink-0 ml-2" style={{ color: e.type === 'recebimento' ? 'var(--income)' : 'var(--text-soft)' }}>{formatBRL(e.amount)}</span>
          </div>
        ))}
        {events.length === 0 && <p className="text-xs text-center py-4" style={{ color: 'var(--text-soft)' }}>Nenhum evento próximo neste mês.</p>}
      </div>
    </Card>
  );
}

/* ---------- Cartões (preview) ---------- */

function CreditCardVisual({ card, gradient }) {
  const percentUsed = card.limit > 0 ? (card.used / card.limit) * 100 : 0;
  const barColor = percentUsed > 85 ? 'var(--expense)' : percentUsed > 65 ? 'var(--alert)' : 'var(--income)';
  const nextDue = useMemo(() => getNextCardDueDate(card), [card.dueDay]);
  const dueWasAdjusted = nextDue.getDate() !== card.dueDay;
  const nextDueLabel = nextDue.toLocaleDateString('pt-BR', { day: '2-digit', month: 'short' });
  return (
    <div className="rounded-2xl p-5 text-white shadow-soft-lg" style={{ background: gradient }}>
      <div className="flex items-start justify-between mb-6">
        <div>
          <p className="text-sm font-medium opacity-90">{card.bank}</p>
          <p className="text-xs opacity-70">{card.brand}</p>
        </div>
        <CreditCard size={22} className="opacity-80" />
      </div>
      <p className="text-xs opacity-70 mb-1">Fatura atual</p>
      <p className="font-display text-2xl font-bold tabular-nums mb-4">{formatBRL(card.invoice)}</p>
      <div className="mb-2">
        <div className="w-full h-1.5 rounded-full bg-white/20 overflow-hidden">
          <div className="h-full rounded-full" style={{ width: `${Math.min(100, percentUsed)}%`, backgroundColor: barColor }} />
        </div>
      </div>
      <div className="flex items-center justify-between text-xs opacity-80">
        <span>{formatBRL(card.used)} de {formatBRL(card.limit)}</span>
        <span>{percentUsed.toFixed(0)}% usado</span>
      </div>
      <div className="flex items-center justify-between text-xs opacity-70 mt-3 pt-3" style={{ borderTop: '1px solid rgba(255,255,255,0.2)' }}>
        <span>Fecha dia {card.closingDay}</span>
        <span className="text-right">
          Próx. vencimento: {nextDueLabel}
          {dueWasAdjusted && <span className="block opacity-70">(dia {card.dueDay}, antecipado p/ dia útil)</span>}
        </span>
      </div>
    </div>
  );
}

const CARD_GRADIENTS = [
  'linear-gradient(135deg, #3E4B37 0%, #232323 100%)',
  'linear-gradient(135deg, #5D7052 0%, #3E4B37 100%)',
  'linear-gradient(135deg, #6D6558 0%, #3A362E 100%)',
];

function CardsPreview({ cards, onSeeAll }) {
  return (
    <Card className="animate-fade-up" padding="p-5 sm:p-6">
      <SectionTitle action={<button onClick={onSeeAll} className="text-xs font-medium flex items-center gap-1" style={{ color: 'var(--primary)' }}>Ver todos <ArrowRight size={12} /></button>}>
        Cartões de crédito
      </SectionTitle>
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">
        {cards.map((c, i) => <CreditCardVisual key={c.id} card={c} gradient={CARD_GRADIENTS[i % CARD_GRADIENTS.length]} />)}
      </div>
    </Card>
  );
}

/* ---------- Assinaturas / Despesas recorrentes (preview) ---------- */

function RecurringPreview({ recurring, onSeeAll }) {
  const total = recurring.reduce((s, r) => s + r.value, 0);
  return (
    <Card className="animate-fade-up" padding="p-5 sm:p-6">
      <SectionTitle action={<button onClick={onSeeAll} className="text-xs font-medium flex items-center gap-1" style={{ color: 'var(--primary)' }}>Ver todas <ArrowRight size={12} /></button>}>
        Despesas recorrentes <span className="text-xs font-normal" style={{ color: 'var(--text-soft)' }}>{formatBRL(total)}/mês</span>
      </SectionTitle>
      <div className="space-y-3">
        {recurring.slice(0, 5).map((r) => {
          const cat = CATEGORIES[r.category] || CATEGORIES['Outros'];
          return (
            <div key={r.id} className="flex items-center gap-3">
              <IconCircle icon={cat.icon} color={cat.color} soft={cat.soft} size={34} />
              <div className="flex-1 min-w-0">
                <p className="text-sm font-medium truncate" style={{ color: 'var(--text)' }}>{r.name}</p>
                <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Renova dia {r.renewalDay}</p>
              </div>
              <span className="text-sm tabular-nums font-medium" style={{ color: 'var(--text)' }}>{formatBRL(r.value)}</span>
            </div>
          );
        })}
      </div>
    </Card>
  );
}

/* ---------- Investimentos (preview) ---------- */

function InvestmentsPreview({ investments, total, onSeeAll }) {
  const totalInvested = investments.reduce((s, i) => s + i.invested, 0);
  const gain = total - totalInvested;
  const gainPercent = totalInvested > 0 ? (gain / totalInvested) * 100 : 0;
  return (
    <Card className="animate-fade-up" padding="p-5 sm:p-6">
      <SectionTitle action={<button onClick={onSeeAll} className="text-xs font-medium flex items-center gap-1" style={{ color: 'var(--primary)' }}>Ver detalhes <ArrowRight size={12} /></button>}>
        Investimentos
      </SectionTitle>
      {investments.length === 0 ? (
        <EmptyState icon={PieChartIcon} title="Nenhum investimento cadastrado" description="Adicione seus investimentos para acompanhar o valor total e a rentabilidade aqui." />
      ) : (
        <>
          <div className="flex items-center justify-between mb-4">
            <div>
              <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Patrimônio investido</p>
              <p className="font-display text-2xl font-bold tabular-nums" style={{ color: 'var(--text)' }}>{formatBRL(total)}</p>
            </div>
            {totalInvested > 0 && (
              <Badge color={gain >= 0 ? 'var(--income)' : 'var(--expense)'} soft={gain >= 0 ? 'var(--income-soft)' : 'var(--expense-soft)'} icon={gain >= 0 ? ArrowUpRight : ArrowDownRight}>
                {gain >= 0 ? '+' : ''}{gainPercent.toFixed(1)}%
              </Badge>
            )}
          </div>
          <div className="space-y-2.5">
            {investments.slice(0, 4).map((inv) => (
              <div key={inv.id} className="flex items-center justify-between text-sm gap-2">
                <span className="flex items-center gap-2 min-w-0">
                  <span className="w-2 h-2 rounded-full shrink-0" style={{ backgroundColor: INVESTMENT_CATEGORIES[inv.category] || '#A8A398' }} />
                  <span className="truncate" style={{ color: 'var(--text)' }}>{inv.name}</span>
                </span>
                <span className="tabular-nums shrink-0" style={{ color: 'var(--text-soft)' }}>{formatBRL(inv.currentValue)}</span>
              </div>
            ))}
          </div>
        </>
      )}
    </Card>
  );
}

/* ---------- Insights inteligentes ---------- */

function InsightsSection({ insights }) {
  const toneColors = {
    income: { color: 'var(--income)', soft: 'var(--income-soft)' },
    expense: { color: 'var(--expense)', soft: 'var(--expense-soft)' },
    invest: { color: 'var(--invest)', soft: 'var(--invest-soft)' },
    goals: { color: 'var(--goals)', soft: 'var(--goals-soft)' },
    alert: { color: 'var(--alert)', soft: 'var(--alert-soft)' },
  };
  return (
    <Card className="animate-fade-up" padding="p-5 sm:p-6">
      <SectionTitle>
        <span className="flex items-center gap-2"><Sparkles size={16} color="var(--primary)" /> Insights inteligentes</span>
      </SectionTitle>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        {insights.map((ins, i) => {
          const t = toneColors[ins.tone] || toneColors.alert;
          return (
            <div key={i} className="flex items-start gap-3 p-3.5 rounded-xl" style={{ backgroundColor: t.soft }}>
              <ins.icon size={16} color={t.color} className="shrink-0 mt-0.5" />
              <p className="text-sm" style={{ color: 'var(--text)' }}>{ins.text}</p>
            </div>
          );
        })}
      </div>
    </Card>
  );
}

/* ============================================================
   PÁGINA: DASHBOARD
   ============================================================ */

function DashboardPage({ data, actions }) {
  const kpis = computeKPIs(data.period, data.customRange, data.monthlyHistory, data.transactions, data.accounts, data.goals, data.caixinhas, data.settings.monthlySavingsTarget);
  return (
    <div className="space-y-6">
      <KPIRow kpis={kpis} />
      <div className="grid grid-cols-1 lg:grid-cols-5 gap-6">
        <div className="lg:col-span-3"><EvolutionChart data={data.monthlyHistory} /></div>
        <div className="lg:col-span-2"><CategoryDonut data={data.categoryComparison.current} /></div>
      </div>
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className="lg:col-span-2">
          <RecentTransactions transactions={data.transactions} accounts={data.accounts} onEdit={actions.editTransaction} onDelete={actions.deleteTransaction} onSeeAll={() => actions.goTo('transacoes')} />
        </div>
        <MonthSummaryPanel transactions={data.transactions} kpis={kpis} />
      </div>
      <GoalsSection goals={data.goals} onAddFunds={actions.addGoalFunds} onDelete={actions.deleteGoal} onSeeAll={() => actions.goTo('metas')} compact />
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <FinancialCalendar cards={data.cards} transactions={data.transactions} />
        <CardsPreview cards={data.cards} onSeeAll={() => actions.goTo('cartoes')} />
      </div>
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <RecurringPreview recurring={data.recurring} onSeeAll={() => actions.goTo('recorrentes')} />
        <InvestmentsPreview investments={data.investments} total={data.investmentsTotal} onSeeAll={() => actions.goTo('investimentos')} />
      </div>
      <InsightsSection insights={data.insights} />
    </div>
  );
}

/* ============================================================
   FORMULÁRIO DE TRANSAÇÃO
   ============================================================ */

function TransactionForm({ initial, accounts, onSave, onClose }) {
  const [form, setForm] = useState(initial || {
    type: 'despesa', description: '', amount: 0, category: 'Mercado', account: accounts[0]?.id || '',
    paymentMethod: 'Pix', date: new Date().toISOString().slice(0, 10), status: 'Pago',
    isSalary: false, grossSalary: 0, dependents: 0,
  });
  const [errors, setErrors] = useState({});

  const cltBreakdown = useMemo(
    () => (form.isSalary ? calcCLTNetSalary(form.grossSalary, form.dependents) : null),
    [form.isSalary, form.grossSalary, form.dependents]
  );

  useEffect(() => {
    if (form.isSalary && cltBreakdown) {
      setForm((f) => ({ ...f, amount: cltBreakdown.net, inssAmount: cltBreakdown.inss, irrfAmount: cltBreakdown.irrf }));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [form.isSalary, cltBreakdown?.net]);

  function validate() {
    const e = {};
    if (!form.description.trim()) e.description = 'Informe uma descrição';
    if (!form.amount || form.amount <= 0) e.amount = 'Informe um valor maior que zero';
    if (!form.date) e.date = 'Informe uma data';
    if (!form.account) e.account = 'Selecione uma conta';
    if (form.isSalary && (!form.grossSalary || form.grossSalary <= 0)) e.grossSalary = 'Informe o salário bruto';
    setErrors(e);
    return Object.keys(e).length === 0;
  }

  return (
    <Modal title={initial ? 'Editar lançamento' : 'Novo lançamento'} onClose={onClose} wide>
      <div className="space-y-4">
        <div className="flex gap-2">
          {['despesa', 'receita'].map((t) => (
            <button
              key={t} onClick={() => setForm({ ...form, type: t, isSalary: t === 'receita' ? form.isSalary : false })}
              className="flex-1 py-2.5 rounded-xl text-sm font-medium border transition-colors"
              style={{
                borderColor: form.type === t ? (t === 'receita' ? 'var(--income)' : 'var(--expense)') : 'var(--border)',
                backgroundColor: form.type === t ? (t === 'receita' ? 'var(--income-soft)' : 'var(--expense-soft)') : 'transparent',
                color: form.type === t ? (t === 'receita' ? 'var(--income)' : 'var(--expense)') : 'var(--text-soft)',
              }}
            >
              {t === 'receita' ? 'Receita' : 'Despesa'}
            </button>
          ))}
        </div>
        {form.type === 'receita' && (
          <label className="flex items-center gap-2 text-sm cursor-pointer select-none" style={{ color: 'var(--text)' }}>
            <input type="checkbox" checked={!!form.isSalary} onChange={(e) => setForm({ ...form, isSalary: e.target.checked, description: e.target.checked && !form.description ? 'Salário' : form.description })} className="focus-ring" />
            É salário (CLT)? Calculamos o líquido estimado com INSS e IRRF.
          </label>
        )}
        <div>
          <FieldLabel error={errors.description}>Descrição</FieldLabel>
          <input value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Supermercado" />
          {errors.description && <p className="text-xs mt-1" style={{ color: 'var(--expense)' }}>{errors.description}</p>}
        </div>

        {form.isSalary && (
          <div className="rounded-xl p-4 space-y-3" style={{ backgroundColor: 'var(--income-soft)' }}>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <FieldLabel error={errors.grossSalary}>Salário bruto</FieldLabel>
                <CurrencyInput value={form.grossSalary} onChange={(v) => setForm({ ...form, grossSalary: v })} />
              </div>
              <div>
                <FieldLabel>Dependentes (IR)</FieldLabel>
                <input type="number" min={0} value={form.dependents} onChange={(e) => setForm({ ...form, dependents: Number(e.target.value) })} className={inputClass} style={inputStyle} />
              </div>
            </div>
            {cltBreakdown && (
              <div className="text-xs space-y-1 pt-2" style={{ borderTop: '1px solid rgba(0,0,0,0.08)' }}>
                <div className="flex justify-between" style={{ color: 'var(--text-soft)' }}><span>INSS</span><span className="tabular-nums">− {formatBRL(cltBreakdown.inss)}</span></div>
                <div className="flex justify-between" style={{ color: 'var(--text-soft)' }}><span>IRRF</span><span className="tabular-nums">− {formatBRL(cltBreakdown.irrf)}</span></div>
                <div className="flex justify-between font-semibold pt-1" style={{ color: 'var(--income)' }}><span>Líquido estimado</span><span className="tabular-nums">{formatBRL(cltBreakdown.net)}</span></div>
                <p className="pt-1" style={{ color: 'var(--text-soft)' }}>Estimativa com base nas tabelas de INSS e IRRF de 2026 (já considerando a redução da Lei 15.270/2025). Outros descontos do holerite (plano de saúde, pensão, adiantamentos etc.) não entram nesta conta.</p>
              </div>
            )}
          </div>
        )}

        <div className="grid grid-cols-2 gap-3">
          <div>
            <FieldLabel error={errors.amount}>Valor</FieldLabel>
            {form.isSalary ? (
              <div className={inputClass} style={{ ...inputStyle, display: 'flex', alignItems: 'center', color: 'var(--text-soft)' }}>{formatBRL(form.amount)} <span className="ml-1 text-xs">(líquido calculado)</span></div>
            ) : (
              <CurrencyInput value={form.amount} onChange={(v) => setForm({ ...form, amount: v })} />
            )}
            {errors.amount && <p className="text-xs mt-1" style={{ color: 'var(--expense)' }}>{errors.amount}</p>}
          </div>
          <div>
            <FieldLabel error={errors.date}>Data</FieldLabel>
            <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} className={inputClass} style={inputStyle} />
            {errors.date && <p className="text-xs mt-1" style={{ color: 'var(--expense)' }}>{errors.date}</p>}
          </div>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <FieldLabel>Categoria</FieldLabel>
            <select value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} className={inputClass} style={inputStyle}>
              {CATEGORY_NAMES.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <FieldLabel error={errors.account}>Conta</FieldLabel>
            <select value={form.account} onChange={(e) => setForm({ ...form, account: e.target.value })} className={inputClass} style={inputStyle}>
              {accounts.map((a) => <option key={a.id} value={a.id}>{a.bank} ({a.type})</option>)}
            </select>
          </div>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <FieldLabel>Forma de pagamento</FieldLabel>
            <select value={form.paymentMethod} onChange={(e) => setForm({ ...form, paymentMethod: e.target.value })} className={inputClass} style={inputStyle}>
              {PAYMENT_METHODS.map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
          </div>
          <div>
            <FieldLabel>Status</FieldLabel>
            <select value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })} className={inputClass} style={inputStyle}>
              {STATUS_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
        </div>
        <div className="flex justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={() => { if (validate()) onSave(form); }}>{initial ? 'Salvar alterações' : 'Adicionar lançamento'}</Button>
        </div>
      </div>
    </Modal>
  );
}

/* ---------- Modal: revisão da importação de fatura (CSV) ---------- */

function ImportReviewModal({ parsed, accounts, onConfirm, onClose }) {
  const nubankAccount = accounts.find((a) => a.bank.toLowerCase().includes('nubank')) || accounts[0];
  const [targetAccount, setTargetAccount] = useState(nubankAccount?.id || '');
  const [rows, setRows] = useState(parsed.purchases);

  const duplicateCount = rows.filter((r) => r.isDuplicate).length;
  const selectedCount = rows.filter((r) => r.include).length;
  const selectedTotal = rows.filter((r) => r.include).reduce((s, r) => s + r.amount, 0);

  function toggleRow(rowId) {
    setRows((prev) => prev.map((r) => (r.rowId === rowId ? { ...r, include: !r.include } : r)));
  }
  function setRowCategory(rowId, category) {
    setRows((prev) => prev.map((r) => (r.rowId === rowId ? { ...r, category } : r)));
  }

  function handleConfirm() {
    const toImport = rows
      .filter((r) => r.include)
      .map((r) => ({
        id: uid(),
        date: r.date,
        description: r.installment ? `${r.description} (parcela ${r.installment.current}/${r.installment.total})` : r.description,
        category: r.category,
        type: 'despesa',
        account: targetAccount,
        paymentMethod: 'Cartão de crédito',
        amount: r.amount,
        status: 'Pago',
      }));
    onConfirm(toImport, { duplicates: duplicateCount, payments: parsed.payments.length, paymentsTotal: parsed.paymentsTotal });
  }

  return (
    <Modal title="Importar lançamentos da fatura" onClose={onClose} wide>
      <div className="space-y-4">
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          <div className="rounded-xl p-3" style={{ backgroundColor: 'var(--primary-soft)' }}>
            <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Novos lançamentos</p>
            <p className="font-display text-lg font-semibold" style={{ color: 'var(--text)' }}>{selectedCount}</p>
          </div>
          <div className="rounded-xl p-3" style={{ backgroundColor: 'var(--expense-soft)' }}>
            <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Total selecionado</p>
            <p className="font-display text-lg font-semibold" style={{ color: 'var(--expense)' }}>{formatBRL(selectedTotal)}</p>
          </div>
          <div className="rounded-xl p-3" style={{ backgroundColor: '#EEEDE8' }}>
            <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Duplicadas ignoradas</p>
            <p className="font-display text-lg font-semibold" style={{ color: 'var(--text)' }}>{duplicateCount}</p>
          </div>
          <div className="rounded-xl p-3" style={{ backgroundColor: 'var(--alert-soft)' }}>
            <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Adiantamentos de fatura</p>
            <p className="font-display text-lg font-semibold" style={{ color: 'var(--alert)' }}>{formatBRL(parsed.paymentsTotal)}</p>
          </div>
        </div>

        {parsed.payments.length > 0 && (
          <div className="rounded-xl p-3 text-xs flex items-start gap-2" style={{ backgroundColor: 'var(--alert-soft)', color: 'var(--text-soft)' }}>
            <Info size={14} className="mt-0.5 shrink-0" color="var(--alert)" />
            <span>
              Identificamos {parsed.payments.length} lançamento(s) de "Pagamento recebido" (seus adiantamentos/parcelas da fatura), totalizando {formatBRL(parsed.paymentsTotal)}.
              Eles não são importados como gasto novo, pois isso duplicaria o valor que já sai da sua conta ao pagar a fatura.
            </span>
          </div>
        )}

        <div>
          <FieldLabel>Lançar na conta</FieldLabel>
          <select value={targetAccount} onChange={(e) => setTargetAccount(e.target.value)} className={inputClass} style={inputStyle}>
            {accounts.map((a) => <option key={a.id} value={a.id}>{a.bank} ({a.type})</option>)}
          </select>
        </div>

        {rows.length === 0 ? (
          <EmptyState icon={Upload} title="Nenhuma compra encontrada" description="O arquivo só continha pagamentos/adiantamentos de fatura." />
        ) : (
          <div className="overflow-x-auto -mx-1 max-h-[40vh] overflow-y-auto">
            <table className="w-full min-w-[640px]">
              <thead>
                <tr className="text-left text-xs sticky top-0" style={{ color: 'var(--text-soft)', backgroundColor: 'var(--card)' }}>
                  <th className="pb-2 pr-2 font-medium w-8"></th>
                  <th className="pb-2 pr-3 font-medium">Data</th>
                  <th className="pb-2 pr-3 font-medium">Descrição</th>
                  <th className="pb-2 pr-3 font-medium">Categoria</th>
                  <th className="pb-2 pr-3 font-medium text-right">Valor</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r) => (
                  <tr key={r.rowId} className="text-sm" style={{ borderTop: '1px solid var(--border)', opacity: r.include ? 1 : 0.5 }}>
                    <td className="py-2 pr-2">
                      <input type="checkbox" checked={r.include} onChange={() => toggleRow(r.rowId)} className="focus-ring" />
                    </td>
                    <td className="py-2 pr-3 whitespace-nowrap" style={{ color: 'var(--text-soft)' }}>{formatDate(r.date)}</td>
                    <td className="py-2 pr-3">
                      <span style={{ color: 'var(--text)' }}>{r.description}</span>
                      {r.installment && <span className="ml-2 text-xs" style={{ color: 'var(--text-soft)' }}>parcela {r.installment.current}/{r.installment.total}</span>}
                      {r.isDuplicate && <span className="ml-2 text-xs font-medium" style={{ color: 'var(--alert)' }}>possível duplicata</span>}
                    </td>
                    <td className="py-2 pr-3">
                      <select value={r.category} onChange={(e) => setRowCategory(r.rowId, e.target.value)} className="px-2 py-1.5 rounded-lg text-xs focus-ring" style={inputStyle}>
                        {CATEGORY_NAMES.map((c) => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </td>
                    <td className="py-2 pr-3 text-right tabular-nums" style={{ color: 'var(--expense)' }}>{formatBRL(r.amount)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        <div className="flex justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={handleConfirm} disabled={selectedCount === 0 || !targetAccount}>Importar {selectedCount > 0 ? `(${selectedCount})` : ''}</Button>
        </div>
      </div>
    </Modal>
  );
}

/* ============================================================
   PÁGINA: TRANSAÇÕES / RECEITAS / DESPESAS (componente genérico)
   ============================================================ */

function TransactionsPage({ filterType, title, transactions, accounts, onAdd, onEdit, onDelete, onImport }) {
  const [search, setSearch] = useState('');
  const [categoryFilter, setCategoryFilter] = useState('all');
  const [accountFilter, setAccountFilter] = useState('all');
  const [statusFilter, setStatusFilter] = useState('all');
  const [sortConfig, setSortConfig] = useState({ key: 'date', direction: 'desc' });
  const [page, setPage] = useState(1);
  const [editing, setEditing] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const fileInputRef = useRef(null);
  const [importPreview, setImportPreview] = useState(null);
  const pageSize = 8;

  const base = filterType === 'all' ? transactions : transactions.filter((t) => t.type === filterType);

  const filtered = useMemo(() => base.filter((t) => {
    const matchesSearch = t.description.toLowerCase().includes(search.toLowerCase());
    const matchesCategory = categoryFilter === 'all' || t.category === categoryFilter;
    const matchesAccount = accountFilter === 'all' || t.account === accountFilter;
    const matchesStatus = statusFilter === 'all' || t.status === statusFilter;
    return matchesSearch && matchesCategory && matchesAccount && matchesStatus;
  }), [base, search, categoryFilter, accountFilter, statusFilter]);

  const sorted = useMemo(() => {
    const arr = [...filtered];
    arr.sort((a, b) => {
      let av = a[sortConfig.key], bv = b[sortConfig.key];
      if (sortConfig.key === 'date') { av = new Date(av); bv = new Date(bv); }
      if (av < bv) return sortConfig.direction === 'asc' ? -1 : 1;
      if (av > bv) return sortConfig.direction === 'asc' ? 1 : -1;
      return 0;
    });
    return arr;
  }, [filtered, sortConfig]);

  useEffect(() => { setPage(1); }, [search, categoryFilter, accountFilter, statusFilter]);

  const totalPages = Math.max(1, Math.ceil(sorted.length / pageSize));
  useEffect(() => { if (page > totalPages) setPage(totalPages); }, [totalPages, page]);
  const pageData = sorted.slice((page - 1) * pageSize, page * pageSize);
  const totalReceitas = filtered.filter((t) => t.type === 'receita').reduce((s, t) => s + t.amount, 0);
  const totalDespesas = filtered.filter((t) => t.type === 'despesa').reduce((s, t) => s + t.amount, 0);

  function handleSort(key) {
    setSortConfig((prev) => ({ key, direction: prev.key === key && prev.direction === 'asc' ? 'desc' : 'asc' }));
  }

  function handleImportFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    Papa.parse(file, {
      header: true, skipEmptyLines: true,
      complete: (results) => {
        const parsed = parseNubankFaturaCSV(results.data, transactions);
        if (parsed.purchases.length === 0 && parsed.payments.length === 0) {
          onImport([], { error: true });
        } else {
          setImportPreview(parsed);
        }
      },
      error: () => onImport([], { error: true }),
    });
    e.target.value = '';
  }

  function handleConfirmImport(rows, meta) {
    setImportPreview(null);
    onImport(rows, meta);
  }

  function exportCSV() {
    const rows = sorted.map((t) => ({ Data: formatDate(t.date), Descrição: t.description, Categoria: t.category, Tipo: t.type, Conta: accounts.find((a) => a.id === t.account)?.bank || 'Conta removida', 'Forma de pagamento': t.paymentMethod, Valor: t.amount, Status: t.status }));
    const csv = Papa.unparse(rows);
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url; link.download = `${title.toLowerCase().replace(/\s/g, '-')}.csv`; link.click();
    URL.revokeObjectURL(url);
  }
  function exportExcel() {
    const rows = sorted.map((t) => ({ Data: formatDate(t.date), Descrição: t.description, Categoria: t.category, Tipo: t.type, Conta: accounts.find((a) => a.id === t.account)?.bank || 'Conta removida', 'Forma de pagamento': t.paymentMethod, Valor: t.amount, Status: t.status }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Transações');
    XLSX.writeFile(wb, `${title.toLowerCase().replace(/\s/g, '-')}.xlsx`);
  }

  const sortableColumns = [
    { key: 'date', label: 'Data' }, { key: 'description', label: 'Descrição' }, { key: 'category', label: 'Categoria' },
    { key: 'amount', label: 'Valor' }, { key: 'status', label: 'Status' },
  ];

  return (
    <div className="space-y-6 print-area">
      {filterType !== 'all' && (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <StatCard title={`Total de ${title.toLowerCase()}`} value={formatBRL(filterType === 'receita' ? totalReceitas : totalDespesas)} description="No filtro atual" icon={filterType === 'receita' ? TrendingUp : TrendingDown} color={filterType === 'receita' ? 'var(--income)' : 'var(--expense)'} soft={filterType === 'receita' ? 'var(--income-soft)' : 'var(--expense-soft)'} />
          <StatCard title="Lançamentos no filtro" value={String(filtered.length)} description="Itens encontrados" icon={ArrowLeftRight} color="var(--invest)" soft="var(--invest-soft)" />
        </div>
      )}

      <Card padding="p-4 sm:p-5">
        <div className="flex flex-wrap items-center gap-2 no-print">
          <div className="relative flex-1 min-w-[180px]">
            <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2" color="var(--text-soft)" />
            <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Pesquisar..." className="w-full pl-8 pr-3 py-2 rounded-xl text-sm focus-ring" style={inputStyle} />
          </div>
          <select value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)} className="px-3 py-2 rounded-xl text-sm focus-ring" style={inputStyle}>
            <option value="all">Todas categorias</option>
            {CATEGORY_NAMES.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
          <select value={accountFilter} onChange={(e) => setAccountFilter(e.target.value)} className="px-3 py-2 rounded-xl text-sm focus-ring" style={inputStyle}>
            <option value="all">Todas contas</option>
            {accounts.map((a) => <option key={a.id} value={a.id}>{a.bank}</option>)}
          </select>
          <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} className="px-3 py-2 rounded-xl text-sm focus-ring" style={inputStyle}>
            <option value="all">Todos status</option>
            {STATUS_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <div className="flex-1" />
          <input type="file" accept=".csv" ref={fileInputRef} onChange={handleImportFile} className="hidden" />
          <Button variant="secondary" size="sm" icon={Upload} onClick={() => fileInputRef.current.click()}>Importar fatura</Button>
          <Button variant="secondary" size="sm" icon={Download} onClick={exportCSV}>CSV</Button>
          <Button variant="secondary" size="sm" icon={Download} onClick={exportExcel}>Excel</Button>
          <Button variant="secondary" size="sm" icon={FileText} onClick={() => window.print()}>Imprimir</Button>
          <Button size="sm" icon={Plus} onClick={() => setShowForm(true)}>Novo</Button>
        </div>

        {pageData.length === 0 ? (
          <EmptyState icon={Search} title="Nenhuma transação encontrada" description="Ajuste os filtros ou adicione um novo lançamento." />
        ) : (
          <div className="overflow-x-auto mt-4 -mx-1">
            <table className="w-full min-w-[720px]">
              <thead>
                <tr className="text-left text-xs" style={{ color: 'var(--text-soft)' }}>
                  {sortableColumns.map((c) => (
                    <th key={c.key} className="pb-2 pr-3 font-medium cursor-pointer select-none" onClick={() => handleSort(c.key)}>
                      <span className="flex items-center gap-1">{c.label} {sortConfig.key === c.key && (sortConfig.direction === 'asc' ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}</span>
                    </th>
                  ))}
                  <th className="pb-2 pr-3 font-medium">Conta</th>
                  <th className="pb-2 pr-3 font-medium">Pagamento</th>
                  <th className="pb-2 pr-1 font-medium no-print"></th>
                </tr>
              </thead>
              <tbody style={{ borderTop: '1px solid var(--border)' }}>
                {pageData.map((tx) => (
                  <tr key={tx.id} className="text-sm hover:bg-black/[0.02]">
                    <td className="py-3 pr-3 whitespace-nowrap" style={{ color: 'var(--text-soft)' }}>{formatDate(tx.date)}</td>
                    <td className="py-3 pr-3" style={{ color: 'var(--text)' }}>{tx.description}</td>
                    <td className="py-3 pr-3"><CategoryBadge category={tx.category} /></td>
                    <td className="py-3 pr-3 tabular-nums font-medium whitespace-nowrap" style={{ color: tx.type === 'receita' ? 'var(--income)' : 'var(--expense)' }}>{tx.type === 'receita' ? '+' : '-'} {formatBRL(tx.amount)}</td>
                    <td className="py-3 pr-3"><StatusBadge status={tx.status} /></td>
                    <td className="py-3 pr-3 whitespace-nowrap" style={{ color: 'var(--text-soft)' }}>{accounts.find((a) => a.id === tx.account)?.bank || 'Conta removida'}</td>
                    <td className="py-3 pr-3 whitespace-nowrap" style={{ color: 'var(--text-soft)' }}>{tx.paymentMethod}</td>
                    <td className="py-3 pr-1 no-print">
                      <div className="flex items-center gap-1">
                        <button onClick={() => setEditing(tx)} className="p-1.5 rounded-lg hover:bg-black/5"><Pencil size={14} color="var(--text-soft)" /></button>
                        <button onClick={() => onDelete(tx)} className="p-1.5 rounded-lg hover:bg-black/5"><Trash2 size={14} color="var(--expense)" /></button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {totalPages > 1 && (
          <div className="flex items-center justify-between mt-4 no-print">
            <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Página {page} de {totalPages} · {sorted.length} lançamentos</p>
            <div className="flex items-center gap-1">
              <button disabled={page === 1} onClick={() => setPage((p) => p - 1)} className="p-2 rounded-lg hover:bg-black/5 disabled:opacity-40"><ChevronLeft size={16} /></button>
              <button disabled={page === totalPages} onClick={() => setPage((p) => p + 1)} className="p-2 rounded-lg hover:bg-black/5 disabled:opacity-40"><ChevronRight size={16} /></button>
            </div>
          </div>
        )}
      </Card>

      {showForm && <TransactionForm accounts={accounts} onSave={(f) => { onAdd(f); setShowForm(false); }} onClose={() => setShowForm(false)} />}
      {editing && <TransactionForm initial={editing} accounts={accounts} onSave={(f) => { onEdit(f); setEditing(null); }} onClose={() => setEditing(null)} />}
      {importPreview && <ImportReviewModal parsed={importPreview} accounts={accounts} onConfirm={handleConfirmImport} onClose={() => setImportPreview(null)} />}
    </div>
  );
}

/* ============================================================
   PÁGINA: CONTAS BANCÁRIAS + CAIXINHAS
   ============================================================ */

function AccountForm({ onSave, onClose }) {
  const [form, setForm] = useState({ bank: '', type: 'Conta Corrente', balance: 0, alertThreshold: 0 });
  const [error, setError] = useState('');
  return (
    <Modal title="Nova conta" onClose={onClose}>
      <div className="space-y-4">
        <div>
          <FieldLabel error={error}>Banco</FieldLabel>
          <input value={form.bank} onChange={(e) => setForm({ ...form, bank: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Nubank" />
        </div>
        <div>
          <FieldLabel>Tipo</FieldLabel>
          <select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })} className={inputClass} style={inputStyle}>
            <option>Conta Corrente</option>
            <option>Poupança</option>
          </select>
        </div>
        <div>
          <FieldLabel>Saldo inicial</FieldLabel>
          <CurrencyInput value={form.balance} onChange={(v) => setForm({ ...form, balance: v })} />
        </div>
        <div>
          <FieldLabel>Alerta de saldo baixo (opcional)</FieldLabel>
          <CurrencyInput value={form.alertThreshold} onChange={(v) => setForm({ ...form, alertThreshold: v })} />
          <p className="text-xs mt-1" style={{ color: 'var(--text-soft)' }}>Se o saldo cair até esse valor, a conta é sinalizada como zona de risco. Deixe 0 para não usar.</p>
        </div>
        <div className="flex justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={() => { if (!form.bank.trim()) { setError('Informe o nome do banco'); return; } onSave(form); }}>Adicionar conta</Button>
        </div>
      </div>
    </Modal>
  );
}

function CaixinhaForm({ accounts, onSave, onClose }) {
  const [form, setForm] = useState({ name: '', balance: 0, accountId: accounts[0]?.id || '' });
  const [error, setError] = useState('');
  return (
    <Modal title="Nova caixinha" onClose={onClose}>
      <div className="space-y-4">
        <div>
          <FieldLabel error={error}>Nome da caixinha</FieldLabel>
          <input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Reserva para o Natal" />
        </div>
        <div>
          <FieldLabel>Valor inicial</FieldLabel>
          <CurrencyInput value={form.balance} onChange={(v) => setForm({ ...form, balance: v })} />
        </div>
        <div>
          <FieldLabel>Conta vinculada</FieldLabel>
          <select value={form.accountId} onChange={(e) => setForm({ ...form, accountId: e.target.value })} className={inputClass} style={inputStyle}>
            {accounts.map((a) => <option key={a.id} value={a.id}>{a.bank} ({a.type})</option>)}
          </select>
        </div>
        <div className="flex justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={() => { if (!form.name.trim()) { setError('Informe um nome'); return; } onSave(form); }}>Criar caixinha</Button>
        </div>
      </div>
    </Modal>
  );
}

function CaixinhaCard({ caixinha, accountName, onAddFunds, onDelete }) {
  const [adding, setAdding] = useState(false);
  const [amount, setAmount] = useState(0);
  const [confirmDelete, setConfirmDelete] = useState(false);
  return (
    <Card padding="p-4">
      <div className="flex items-start gap-3 mb-3">
        <IconCircle icon={PiggyBank} color="var(--goals)" soft="var(--goals-soft)" size={36} />
        <div className="min-w-0 flex-1">
          <p className="text-sm font-medium truncate" style={{ color: 'var(--text)' }}>{caixinha.name}</p>
          <p className="text-xs" style={{ color: 'var(--text-soft)' }}>{accountName}</p>
        </div>
        <button onClick={() => setConfirmDelete(true)} className="p-1 rounded-lg hover:bg-black/5"><Trash2 size={13} color="var(--text-soft)" /></button>
      </div>
      <p className="font-display text-lg font-bold tabular-nums mb-2" style={{ color: 'var(--text)' }}>{formatBRL(caixinha.balance)}</p>
      {!adding ? (
        <button onClick={() => setAdding(true)} className="text-xs font-medium flex items-center gap-1" style={{ color: 'var(--primary)' }}><Plus size={12} /> Guardar valor</button>
      ) : (
        <div className="flex items-center gap-2">
          <CurrencyInput value={amount} onChange={setAmount} />
          <Button size="sm" onClick={() => { if (amount > 0) onAddFunds(caixinha, amount); setAdding(false); setAmount(0); }}>OK</Button>
        </div>
      )}
      {confirmDelete && (
        <ConfirmModal
          title="Excluir caixinha"
          description={`Tem certeza que deseja excluir a caixinha "${caixinha.name}"? O saldo guardado nela (${formatBRL(caixinha.balance)}) não será somado de volta à conta automaticamente.`}
          onConfirm={() => { onDelete(caixinha); setConfirmDelete(false); }}
          onClose={() => setConfirmDelete(false)}
        />
      )}
    </Card>
  );
}

function AccountCard({ acc, onDelete, onSetThreshold, onSetBalance, usageCount }) {
  const Icon = ACCOUNTS_ICONS[acc.type] || Wallet;
  const [editingThreshold, setEditingThreshold] = useState(false);
  const [thresholdValue, setThresholdValue] = useState(acc.alertThreshold || 0);
  const [editingBalance, setEditingBalance] = useState(false);
  const [balanceValue, setBalanceValue] = useState(acc.balance);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const atRisk = acc.alertThreshold > 0 && acc.balance <= acc.alertThreshold;
  return (
    <Card className="animate-fade-up" style={atRisk ? { border: '1px solid var(--expense)' } : undefined}>
      <div className="flex items-start justify-between mb-4">
        <IconCircle icon={Icon} color="var(--invest)" soft="var(--invest-soft)" />
        <button onClick={() => setConfirmDelete(true)} className="p-1.5 rounded-lg hover:bg-black/5"><Trash2 size={14} color="var(--text-soft)" /></button>
      </div>
      <p className="text-sm font-medium" style={{ color: 'var(--text)' }}>{acc.bank}</p>
      <Badge color="var(--invest)" soft="var(--invest-soft)">{acc.type}</Badge>
      {!editingBalance ? (
        <div className="flex items-center gap-1.5 mt-3">
          <p className="font-display text-xl font-bold tabular-nums" style={{ color: atRisk ? 'var(--expense)' : 'var(--text)' }}>{formatBRL(acc.balance)}</p>
          <button onClick={() => { setBalanceValue(acc.balance); setEditingBalance(true); }} className="p-1 rounded-lg hover:bg-black/5"><Pencil size={12} color="var(--text-soft)" /></button>
        </div>
      ) : (
        <div className="flex items-center gap-2 mt-3">
          <CurrencyInput value={balanceValue} onChange={setBalanceValue} />
          <Button size="sm" onClick={() => { onSetBalance(acc, balanceValue); setEditingBalance(false); }}>OK</Button>
          <button onClick={() => setEditingBalance(false)} className="p-2"><X size={14} color="var(--text-soft)" /></button>
        </div>
      )}
      {atRisk && (
        <p className="text-xs flex items-center gap-1 mt-1" style={{ color: 'var(--expense)' }}>
          <AlertCircle size={12} /> Saldo em zona de risco (abaixo de {formatBRL(acc.alertThreshold)})
        </p>
      )}
      {!editingThreshold ? (
        <button onClick={() => setEditingThreshold(true)} className="text-xs font-medium mt-3 flex items-center gap-1" style={{ color: 'var(--primary)' }}>
          <Bell size={12} /> {acc.alertThreshold > 0 ? `Alerta em ${formatBRL(acc.alertThreshold)}` : 'Definir alerta de saldo baixo'}
        </button>
      ) : (
        <div className="flex items-center gap-2 mt-3">
          <CurrencyInput value={thresholdValue} onChange={setThresholdValue} />
          <Button size="sm" onClick={() => { onSetThreshold(acc, thresholdValue); setEditingThreshold(false); }}>OK</Button>
          <button onClick={() => setEditingThreshold(false)} className="p-2"><X size={14} color="var(--text-soft)" /></button>
        </div>
      )}
      {confirmDelete && (
        <ConfirmModal
          title="Excluir conta"
          description={usageCount > 0
            ? `Esta conta tem ${usageCount} lançamento(s) ou caixinha(s) vinculados. Eles não serão excluídos, mas ficarão sem uma conta associada. Deseja continuar?`
            : `Tem certeza que deseja excluir "${acc.bank}"? Essa ação não pode ser desfeita.`}
          onConfirm={() => { onDelete(acc); setConfirmDelete(false); }}
          onClose={() => setConfirmDelete(false)}
        />
      )}
    </Card>
  );
}

function AccountsPage({ accounts, caixinhas, transactions, onAddAccount, onDeleteAccount, onSetAccountThreshold, onSetAccountBalance, onAddCaixinha, onDeleteCaixinha, onAddCaixinhaFunds }) {
  const [showAccountForm, setShowAccountForm] = useState(false);
  const [showCaixinhaForm, setShowCaixinhaForm] = useState(false);
  const total = accounts.reduce((s, a) => s + a.balance, 0);
  return (
    <div className="space-y-6">
      <Card>
        <div className="flex items-center justify-between mb-1">
          <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Saldo total em contas</p>
        </div>
        <p className="font-display text-3xl font-bold tabular-nums" style={{ color: 'var(--text)' }}>{formatBRL(total)}</p>
      </Card>

      <div>
        <SectionTitle action={<Button size="sm" icon={Plus} onClick={() => setShowAccountForm(true)}>Nova conta</Button>}>Suas contas</SectionTitle>
        <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">
          {accounts.map((acc) => {
            const usageCount = transactions.filter((t) => t.account === acc.id).length + caixinhas.filter((c) => c.accountId === acc.id).length;
            return <AccountCard key={acc.id} acc={acc} onDelete={onDeleteAccount} onSetThreshold={onSetAccountThreshold} onSetBalance={onSetAccountBalance} usageCount={usageCount} />;
          })}
        </div>
      </div>

      <div>
        <SectionTitle action={<Button size="sm" icon={Plus} onClick={() => setShowCaixinhaForm(true)}>Nova caixinha</Button>}>
          Caixinhas <span className="text-xs font-normal" style={{ color: 'var(--text-soft)' }}>reservas flexíveis dentro das suas contas</span>
        </SectionTitle>
        {caixinhas.length === 0 ? (
          <Card><EmptyState icon={PiggyBank} title="Nenhuma caixinha criada" description="Crie caixinhas para separar dinheiro para objetivos do dia a dia." /></Card>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">
            {caixinhas.map((cx) => (
              <CaixinhaCard key={cx.id} caixinha={cx} accountName={accounts.find((a) => a.id === cx.accountId)?.bank || ''} onAddFunds={onAddCaixinhaFunds} onDelete={onDeleteCaixinha} />
            ))}
          </div>
        )}
      </div>

      {showAccountForm && <AccountForm onSave={(f) => { onAddAccount(f); setShowAccountForm(false); }} onClose={() => setShowAccountForm(false)} />}
      {showCaixinhaForm && <CaixinhaForm accounts={accounts} onSave={(f) => { onAddCaixinha(f); setShowCaixinhaForm(false); }} onClose={() => setShowCaixinhaForm(false)} />}
    </div>
  );
}

/* ============================================================
   PÁGINA: CARTÕES
   ============================================================ */

function CardForm({ onSave, onClose }) {
  const [form, setForm] = useState({ bank: '', brand: '', limit: 0, used: 0, invoice: 0, closingDay: 1, dueDay: 10 });
  const [errors, setErrors] = useState({});
  const clampDay = (v) => Math.min(31, Math.max(1, Math.round(v) || 1));
  function validate() {
    const e = {};
    if (!form.bank.trim()) e.bank = 'Informe o banco';
    if (!form.limit || form.limit <= 0) e.limit = 'Informe um limite maior que zero';
    setErrors(e);
    return Object.keys(e).length === 0;
  }
  return (
    <Modal title="Novo cartão" onClose={onClose}>
      <div className="space-y-4">
        <div className="grid grid-cols-2 gap-3">
          <div>
            <FieldLabel error={errors.bank}>Banco</FieldLabel>
            <input value={form.bank} onChange={(e) => setForm({ ...form, bank: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Nubank" />
          </div>
          <div>
            <FieldLabel>Bandeira</FieldLabel>
            <input value={form.brand} onChange={(e) => setForm({ ...form, brand: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Mastercard" />
          </div>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <FieldLabel error={errors.limit}>Limite total</FieldLabel>
            <CurrencyInput value={form.limit} onChange={(v) => setForm({ ...form, limit: v })} />
          </div>
          <div><FieldLabel>Fatura atual</FieldLabel><CurrencyInput value={form.invoice} onChange={(v) => setForm({ ...form, invoice: v, used: v })} /></div>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div><FieldLabel>Dia do fechamento</FieldLabel><input type="number" min={1} max={31} value={form.closingDay} onChange={(e) => setForm({ ...form, closingDay: Number(e.target.value) })} onBlur={(e) => setForm({ ...form, closingDay: clampDay(Number(e.target.value)) })} className={inputClass} style={inputStyle} /></div>
          <div><FieldLabel>Dia do vencimento</FieldLabel><input type="number" min={1} max={31} value={form.dueDay} onChange={(e) => setForm({ ...form, dueDay: Number(e.target.value) })} onBlur={(e) => setForm({ ...form, dueDay: clampDay(Number(e.target.value)) })} className={inputClass} style={inputStyle} /></div>
        </div>
        <div className="flex justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={() => { if (!validate()) return; onSave({ ...form, closingDay: clampDay(form.closingDay), dueDay: clampDay(form.dueDay) }); }}>Adicionar cartão</Button>
        </div>
      </div>
    </Modal>
  );
}

/* ---------- Vale-benefícios (ex: Flash — alimentação + transporte no mesmo cartão) ---------- */

function BenefitForm({ onSave, onClose }) {
  const [form, setForm] = useState({ provider: 'Flash', foodBalance: 0, mobilityBalance: 0 });
  const [error, setError] = useState('');
  return (
    <Modal title="Novo vale-benefícios" onClose={onClose}>
      <div className="space-y-4">
        <div>
          <FieldLabel error={error}>Fornecedor</FieldLabel>
          <input value={form.provider} onChange={(e) => setForm({ ...form, provider: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Flash" />
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div><FieldLabel>Saldo alimentação</FieldLabel><CurrencyInput value={form.foodBalance} onChange={(v) => setForm({ ...form, foodBalance: v })} /></div>
          <div><FieldLabel>Saldo transporte</FieldLabel><CurrencyInput value={form.mobilityBalance} onChange={(v) => setForm({ ...form, mobilityBalance: v })} /></div>
        </div>
        <div className="flex justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={() => { if (!form.provider.trim()) { setError('Informe o fornecedor'); return; } onSave(form); }}>Adicionar</Button>
        </div>
      </div>
    </Modal>
  );
}

function BenefitBalance({ label, icon: Icon, value, onAdjust }) {
  const [adjusting, setAdjusting] = useState(false);
  const [amount, setAmount] = useState(0);
  return (
    <div className="rounded-xl p-3" style={{ backgroundColor: 'var(--bg)' }}>
      <div className="flex items-center gap-1.5 mb-1 text-xs" style={{ color: 'var(--text-soft)' }}><Icon size={13} /> {label}</div>
      <p className="font-display text-base font-bold tabular-nums" style={{ color: 'var(--text)' }}>{formatBRL(value)}</p>
      {!adjusting ? (
        <button onClick={() => setAdjusting(true)} className="text-[11px] font-medium mt-1.5" style={{ color: 'var(--primary)' }}>Ajustar saldo</button>
      ) : (
        <div className="mt-2 space-y-1.5">
          <CurrencyInput value={amount} onChange={setAmount} />
          <div className="flex gap-1.5">
            <Button size="sm" variant="secondary" className="flex-1" onClick={() => { if (amount > 0) onAdjust(amount); setAdjusting(false); setAmount(0); }}>+ Recarga</Button>
            <Button size="sm" variant="secondary" className="flex-1" onClick={() => { if (amount > 0) onAdjust(-amount); setAdjusting(false); setAmount(0); }}>− Uso</Button>
          </div>
        </div>
      )}
    </div>
  );
}

function BenefitCard({ benefit, onAdjust, onDelete }) {
  return (
    <Card padding="p-4">
      <div className="flex items-start gap-3 mb-3">
        <IconCircle icon={CreditCard} color="var(--goals)" soft="var(--goals-soft)" size={36} />
        <div className="min-w-0 flex-1">
          <p className="text-sm font-medium truncate" style={{ color: 'var(--text)' }}>{benefit.provider}</p>
          <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Vale-benefícios</p>
        </div>
        <button onClick={() => onDelete(benefit)} className="p-1 rounded-lg hover:bg-black/5"><Trash2 size={13} color="var(--text-soft)" /></button>
      </div>
      <div className="grid grid-cols-2 gap-3">
        <BenefitBalance label="Alimentação" icon={Utensils} value={benefit.foodBalance} onAdjust={(delta) => onAdjust(benefit, 'foodBalance', delta)} />
        <BenefitBalance label="Transporte" icon={Car} value={benefit.mobilityBalance} onAdjust={(delta) => onAdjust(benefit, 'mobilityBalance', delta)} />
      </div>
    </Card>
  );
}

function CardsPage({ cards, onAdd, onDelete, benefits, onAddBenefit, onDeleteBenefit, onAdjustBenefit }) {
  const [showForm, setShowForm] = useState(false);
  const [showBenefitForm, setShowBenefitForm] = useState(false);
  const [confirmDeleteCard, setConfirmDeleteCard] = useState(null);
  const [confirmDeleteBenefit, setConfirmDeleteBenefit] = useState(null);
  return (
    <div className="space-y-6">
      <SectionTitle action={<Button size="sm" icon={Plus} onClick={() => setShowForm(true)}>Novo cartão</Button>}>Seus cartões</SectionTitle>
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-5">
        {cards.map((c, i) => (
          <div key={c.id} className="relative group">
            <CreditCardVisual card={c} gradient={CARD_GRADIENTS[i % CARD_GRADIENTS.length]} />
            <button onClick={() => setConfirmDeleteCard(c)} className="absolute top-4 right-4 p-1.5 rounded-lg bg-black/20 opacity-0 group-hover:opacity-100 transition-opacity">
              <Trash2 size={14} color="#fff" />
            </button>
          </div>
        ))}
      </div>

      <SectionTitle action={<Button size="sm" icon={Plus} onClick={() => setShowBenefitForm(true)}>Novo vale-benefícios</Button>}>
        Vale-benefícios <span className="text-xs font-normal" style={{ color: 'var(--text-soft)' }}>alimentação e transporte no mesmo cartão</span>
      </SectionTitle>
      {benefits.length === 0 ? (
        <Card><EmptyState icon={Utensils} title="Nenhum vale-benefícios cadastrado" description="Adicione seu cartão Flash (ou similar) para acompanhar alimentação e transporte separadamente." /></Card>
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">
          {benefits.map((b) => (
            <BenefitCard key={b.id} benefit={b} onAdjust={onAdjustBenefit} onDelete={() => setConfirmDeleteBenefit(b)} />
          ))}
        </div>
      )}

      {showForm && <CardForm onSave={(f) => { onAdd(f); setShowForm(false); }} onClose={() => setShowForm(false)} />}
      {showBenefitForm && <BenefitForm onSave={(f) => { onAddBenefit(f); setShowBenefitForm(false); }} onClose={() => setShowBenefitForm(false)} />}
      {confirmDeleteCard && (
        <ConfirmModal
          title="Excluir cartão"
          description={`Tem certeza que deseja excluir o cartão "${confirmDeleteCard.bank}"? Essa ação não pode ser desfeita.`}
          onConfirm={() => { onDelete(confirmDeleteCard); setConfirmDeleteCard(null); }}
          onClose={() => setConfirmDeleteCard(null)}
        />
      )}
      {confirmDeleteBenefit && (
        <ConfirmModal
          title="Excluir vale-benefícios"
          description={`Tem certeza que deseja excluir "${confirmDeleteBenefit.provider}"? Essa ação não pode ser desfeita.`}
          onConfirm={() => { onDeleteBenefit(confirmDeleteBenefit); setConfirmDeleteBenefit(null); }}
          onClose={() => setConfirmDeleteBenefit(null)}
        />
      )}
    </div>
  );
}

/* ============================================================
   PÁGINA: INVESTIMENTOS
   ============================================================ */

function InvestmentForm({ initial, onSave, onClose }) {
  const [form, setForm] = useState(initial || { name: '', category: INVESTMENT_CATEGORY_NAMES[0], invested: 0, currentValue: 0 });
  const [error, setError] = useState('');
  return (
    <Modal title={initial ? 'Editar investimento' : 'Novo investimento'} onClose={onClose}>
      <div className="space-y-4">
        <div>
          <FieldLabel error={error}>Nome</FieldLabel>
          <input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Tesouro Selic" />
        </div>
        <div>
          <FieldLabel>Categoria</FieldLabel>
          <select value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} className={inputClass} style={inputStyle}>
            {INVESTMENT_CATEGORY_NAMES.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div><FieldLabel>Valor investido</FieldLabel><CurrencyInput value={form.invested} onChange={(v) => setForm({ ...form, invested: v })} /></div>
          <div><FieldLabel>Valor atual</FieldLabel><CurrencyInput value={form.currentValue} onChange={(v) => setForm({ ...form, currentValue: v })} /></div>
        </div>
        <p className="text-xs" style={{ color: 'var(--text-soft)' }}>A rentabilidade é calculada automaticamente a partir da diferença entre esses dois valores.</p>
        <div className="flex justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={() => { if (!form.name.trim()) { setError('Informe o nome'); return; } onSave(form); }}>{initial ? 'Salvar' : 'Adicionar'}</Button>
        </div>
      </div>
    </Modal>
  );
}

function InvestmentsPage({ investments, onAdd, onEdit, onDelete }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);

  const totalInvested = investments.reduce((s, i) => s + i.invested, 0);
  const totalCurrent = investments.reduce((s, i) => s + i.currentValue, 0);
  const gain = totalCurrent - totalInvested;
  const gainPercent = totalInvested > 0 ? (gain / totalInvested) * 100 : 0;

  const byCategory = {};
  investments.forEach((i) => { byCategory[i.category] = (byCategory[i.category] || 0) + i.currentValue; });
  const allocation = Object.entries(byCategory).map(([name, value]) => ({ name, value, color: INVESTMENT_CATEGORIES[name] || '#A8A398' }));

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <StatCard title="Valor investido" value={formatBRL(totalInvested)} description="Total aportado" icon={Wallet} color="var(--primary)" soft="var(--primary-soft)" />
        <StatCard title="Valor atual" value={formatBRL(totalCurrent)} description="Valor de mercado hoje" icon={PiggyBank} color="var(--invest)" soft="var(--invest-soft)" />
        <StatCard title="Rentabilidade" value={`${gain >= 0 ? '+' : ''}${gainPercent.toFixed(1)}%`} description={formatBRL(gain)} icon={gain >= 0 ? TrendingUp : TrendingDown} color={gain >= 0 ? 'var(--income)' : 'var(--expense)'} soft={gain >= 0 ? 'var(--income-soft)' : 'var(--expense-soft)'} />
      </div>

      {investments.length > 0 && (
        <Card className="max-w-xs">
          <SectionTitle>Distribuição por categoria</SectionTitle>
          <div style={{ width: '100%', height: 180 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie data={allocation} dataKey="value" nameKey="name" innerRadius={48} outerRadius={72} paddingAngle={2}>
                  {allocation.map((a, i) => <Cell key={i} fill={a.color} />)}
                </Pie>
                <Tooltip formatter={(v) => formatBRL(v)} contentStyle={{ borderRadius: 12, border: '1px solid #E7E4DE', fontSize: 12 }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div className="space-y-2 mt-2">
            {allocation.map((a, i) => (
              <div key={i} className="flex items-center gap-2 text-sm">
                <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: a.color }} />
                <span className="flex-1 truncate" style={{ color: 'var(--text)' }}>{a.name}</span>
                <span className="tabular-nums" style={{ color: 'var(--text-soft)' }}>{totalCurrent > 0 ? ((a.value / totalCurrent) * 100).toFixed(0) : 0}%</span>
              </div>
            ))}
          </div>
        </Card>
      )}

      <Card>
        <SectionTitle action={<Button size="sm" icon={Plus} onClick={() => setShowForm(true)}>Novo investimento</Button>}>Carteira detalhada</SectionTitle>
        {investments.length === 0 ? (
          <EmptyState icon={PiggyBank} title="Nenhum investimento cadastrado" description="Adicione seus investimentos para acompanhar valor e rentabilidade reais." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full min-w-[560px]">
              <thead>
                <tr className="text-left text-xs" style={{ color: 'var(--text-soft)' }}>
                  <th className="pb-2 pr-3 font-medium">Ativo</th>
                  <th className="pb-2 pr-3 font-medium">Investido</th>
                  <th className="pb-2 pr-3 font-medium">Atual</th>
                  <th className="pb-2 pr-3 font-medium">Rentabilidade</th>
                  <th className="pb-2 pr-3 font-medium w-16"></th>
                </tr>
              </thead>
              <tbody style={{ borderTop: '1px solid var(--border)' }}>
                {investments.map((inv) => {
                  const g = inv.currentValue - inv.invested;
                  const gp = inv.invested > 0 ? (g / inv.invested) * 100 : 0;
                  return (
                    <tr key={inv.id} className="text-sm">
                      <td className="py-3 pr-3" style={{ color: 'var(--text)' }}><span className="inline-flex items-center gap-2"><span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: INVESTMENT_CATEGORIES[inv.category] || '#A8A398' }} />{inv.name}</span></td>
                      <td className="py-3 pr-3 tabular-nums" style={{ color: 'var(--text-soft)' }}>{formatBRL(inv.invested)}</td>
                      <td className="py-3 pr-3 tabular-nums" style={{ color: 'var(--text)' }}>{formatBRL(inv.currentValue)}</td>
                      <td className="py-3 pr-3 tabular-nums font-medium" style={{ color: g >= 0 ? 'var(--income)' : 'var(--expense)' }}>{g >= 0 ? '+' : ''}{gp.toFixed(1)}%</td>
                      <td className="py-3 pr-3">
                        <div className="flex items-center gap-1">
                          <button onClick={() => setEditing(inv)} className="p-1.5 rounded-lg hover:bg-black/5"><Pencil size={13} color="var(--text-soft)" /></button>
                          <button onClick={() => setConfirmDelete(inv)} className="p-1.5 rounded-lg hover:bg-black/5"><Trash2 size={13} color="var(--expense)" /></button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      {showForm && <InvestmentForm onSave={(f) => { onAdd(f); setShowForm(false); }} onClose={() => setShowForm(false)} />}
      {editing && <InvestmentForm initial={editing} onSave={(f) => { onEdit(f); setEditing(null); }} onClose={() => setEditing(null)} />}
      {confirmDelete && (
        <ConfirmModal
          title="Excluir investimento"
          description={`Tem certeza que deseja excluir "${confirmDelete.name}"? Essa ação não pode ser desfeita.`}
          onConfirm={() => { onDelete(confirmDelete); setConfirmDelete(null); }}
          onClose={() => setConfirmDelete(null)}
        />
      )}
    </div>
  );
}

/* ============================================================
   PÁGINA: METAS
   ============================================================ */

function GoalsPage({ goals, onAdd, onAddFunds, onDelete }) {
  const [showForm, setShowForm] = useState(false);
  return (
    <div className="space-y-6">
      <SectionTitle action={<Button size="sm" icon={Plus} onClick={() => setShowForm(true)}>Nova meta</Button>}>Suas metas financeiras</SectionTitle>
      {goals.length === 0 ? (
        <Card><EmptyState icon={Target} title="Nenhuma meta criada" description="Crie metas para acompanhar seus objetivos financeiros." /></Card>
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">
          {goals.map((g) => <GoalCard key={g.id} goal={g} onAddFunds={onAddFunds} onDelete={onDelete} />)}
        </div>
      )}
      {showForm && <GoalForm onSave={(f) => { onAdd(f); setShowForm(false); }} onClose={() => setShowForm(false)} />}
    </div>
  );
}

/* ============================================================
   PÁGINA: DESPESAS RECORRENTES
   ============================================================ */

function RecurringForm({ onSave, onClose }) {
  const [form, setForm] = useState({ name: '', category: 'Lazer', value: 0, renewalDay: 1 });
  const [error, setError] = useState('');
  return (
    <Modal title="Nova despesa recorrente" onClose={onClose}>
      <div className="space-y-4">
        <div>
          <FieldLabel error={error}>Nome</FieldLabel>
          <input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Academia" />
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <FieldLabel>Categoria</FieldLabel>
            <select value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} className={inputClass} style={inputStyle}>
              {CATEGORY_NAMES.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <FieldLabel>Dia da renovação</FieldLabel>
            <input type="number" min={1} max={31} value={form.renewalDay} onChange={(e) => setForm({ ...form, renewalDay: Number(e.target.value) })} className={inputClass} style={inputStyle} />
          </div>
        </div>
        <div>
          <FieldLabel>Valor mensal</FieldLabel>
          <CurrencyInput value={form.value} onChange={(v) => setForm({ ...form, value: v })} />
        </div>
        <div className="flex justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={() => { if (!form.name.trim()) { setError('Informe um nome'); return; } onSave(form); }}>Adicionar</Button>
        </div>
      </div>
    </Modal>
  );
}

function RecurringExpensesPage({ recurring, accounts, onAdd, onDelete, onLaunchNow }) {
  const [showForm, setShowForm] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const total = recurring.reduce((s, r) => s + r.value, 0);
  const byCategory = {};
  recurring.forEach((r) => { byCategory[r.category] = (byCategory[r.category] || 0) + r.value; });
  const chartData = Object.entries(byCategory).map(([name, value]) => ({ name, value, color: (CATEGORIES[name] || CATEGORIES['Outros']).color }));

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <StatCard title="Total recorrente mensal" value={formatBRL(total)} description={`${recurring.length} assinaturas e despesas fixas`} icon={RefreshCw} color="var(--primary)" soft="var(--primary-soft)" />
        <StatCard title="Total anual estimado" value={formatBRL(total * 12)} description="Projeção com base no valor atual" icon={CalendarIcon} color="var(--invest)" soft="var(--invest-soft)" />
      </div>
      <p className="text-xs -mt-2" style={{ color: 'var(--text-soft)' }}>
        Esta lista é só um controle de assinaturas — ela não entra sozinha nas suas despesas. Use "Lançar agora" para registrar o mês como um lançamento de verdade.
      </p>

      <div className="grid grid-cols-1 lg:grid-cols-5 gap-6">
        <Card className="lg:col-span-2">
          <SectionTitle>Por categoria <span className="text-xs font-normal" style={{ color: 'var(--text-soft)' }}>onde você mais gasta de forma fixa</span></SectionTitle>
          <div style={{ width: '100%', height: 200 }}>
            <ResponsiveContainer>
              <BarChart data={chartData} layout="vertical" margin={{ left: 0, right: 16 }}>
                <XAxis type="number" tick={{ fontSize: 11, fill: '#6D6D6D' }} axisLine={false} tickLine={false} tickFormatter={(v) => `R$${v}`} />
                <YAxis type="category" dataKey="name" tick={{ fontSize: 11, fill: '#6D6D6D' }} axisLine={false} tickLine={false} width={110} />
                <Tooltip formatter={(v) => formatBRL(v)} contentStyle={{ borderRadius: 12, border: '1px solid #E7E4DE', fontSize: 12 }} />
                <Bar dataKey="value" radius={[0, 6, 6, 0]}>
                  {chartData.map((entry, i) => <Cell key={i} fill={entry.color} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </Card>
        <Card className="lg:col-span-3">
          <SectionTitle action={<Button size="sm" icon={Plus} onClick={() => setShowForm(true)}>Nova despesa</Button>}>Assinaturas e despesas fixas</SectionTitle>
          <div className="space-y-3">
            {recurring.map((r) => {
              const cat = CATEGORIES[r.category] || CATEGORIES['Outros'];
              return (
                <div key={r.id} className="flex items-center gap-3 p-2.5 rounded-xl hover:bg-black/[0.02]">
                  <IconCircle icon={cat.icon} color={cat.color} soft={cat.soft} size={38} />
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-medium truncate" style={{ color: 'var(--text)' }}>{r.name}</p>
                    <div className="flex items-center gap-2 mt-0.5">
                      <CategoryBadge category={r.category} />
                      <span className="text-xs" style={{ color: 'var(--text-soft)' }}>Renova dia {r.renewalDay}</span>
                    </div>
                  </div>
                  <span className="text-sm tabular-nums font-medium shrink-0" style={{ color: 'var(--text)' }}>{formatBRL(r.value)}</span>
                  <button
                    onClick={() => onLaunchNow(r, accounts[0]?.id)}
                    disabled={!accounts[0]}
                    className="text-xs font-medium shrink-0 px-2.5 py-1.5 rounded-lg"
                    style={{ color: 'var(--primary)', backgroundColor: 'var(--primary-soft)' }}
                  >
                    Lançar agora
                  </button>
                  <button onClick={() => setConfirmDelete(r)} className="p-1.5 rounded-lg hover:bg-black/5 shrink-0"><Trash2 size={14} color="var(--text-soft)" /></button>
                </div>
              );
            })}
          </div>
        </Card>
      </div>
      {showForm && <RecurringForm onSave={(f) => { onAdd(f); setShowForm(false); }} onClose={() => setShowForm(false)} />}
      {confirmDelete && (
        <ConfirmModal
          title="Excluir despesa recorrente"
          description={`Tem certeza que deseja excluir "${confirmDelete.name}" da lista? Essa ação não pode ser desfeita.`}
          onConfirm={() => { onDelete(confirmDelete); setConfirmDelete(null); }}
          onClose={() => setConfirmDelete(null)}
        />
      )}
    </div>
  );
}

/* ============================================================
   PÁGINA: RELATÓRIOS
   ============================================================ */

function ReportsPage({ monthlyHistory, categoryComparison }) {
  const barData = monthlyHistory.map((m) => ({ month: m.month, Receitas: m.receitas, Despesas: m.despesas }));
  function exportSummary() {
    const rows = monthlyHistory.map((m) => ({ Mês: m.month, Receitas: m.receitas, Despesas: m.despesas, Saldo: m.receitas - m.despesas, Patrimônio: m.patrimonio }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Resumo mensal');
    XLSX.writeFile(wb, 'relatorio-financeiro.xlsx');
  }
  return (
    <div className="space-y-6 print-area">
      <div className="flex justify-end gap-2 no-print">
        <Button variant="secondary" size="sm" icon={Download} onClick={exportSummary}>Exportar Excel</Button>
        <Button variant="secondary" size="sm" icon={FileText} onClick={() => window.print()}>Imprimir / PDF</Button>
      </div>
      <Card>
        <SectionTitle>Receitas vs. Despesas (12 meses)</SectionTitle>
        <div style={{ width: '100%', height: 280 }}>
          <ResponsiveContainer>
            <BarChart data={barData}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" vertical={false} />
              <XAxis dataKey="month" tick={{ fontSize: 11, fill: '#6D6D6D' }} axisLine={{ stroke: '#E7E4DE' }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: '#6D6D6D' }} axisLine={false} tickLine={false} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} />
              <Tooltip formatter={(v) => formatBRL(v)} contentStyle={{ borderRadius: 12, border: '1px solid #E7E4DE', fontSize: 12 }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="Receitas" fill="#5A8F5A" radius={[6, 6, 0, 0]} />
              <Bar dataKey="Despesas" fill="#B66B6B" radius={[6, 6, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </Card>
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <CategoryDonut data={categoryComparison.current} title="Gastos por categoria" subtitle="mês atual" />
        <CategoryDonut data={categoryComparison.previous} title="Gastos por categoria" subtitle="mês anterior" />
      </div>
    </div>
  );
}

/* ============================================================
   PÁGINA: CONFIGURAÇÕES
   ============================================================ */

function ToggleSwitch({ checked, onChange }) {
  return (
    <button onClick={() => onChange(!checked)} className="w-11 h-6 rounded-full p-0.5 transition-colors focus-ring" style={{ backgroundColor: checked ? 'var(--primary)' : 'var(--border)' }}>
      <div className="w-5 h-5 rounded-full bg-white shadow-sm transition-transform" style={{ transform: checked ? 'translateX(20px)' : 'translateX(0)' }} />
    </button>
  );
}

function SettingsPage({ userName, settings, onChangeName, onChangeSettings, onReset, dropboxConnected, dropboxBusy, dropboxLastBackup, dropboxSyncError, onConnectDropbox, onDisconnectDropbox, onBackupNow, onRestoreFromDropbox }) {
  const [name, setName] = useState(userName);
  const [showConfirmReset, setShowConfirmReset] = useState(false);
  const [showConfirmRestore, setShowConfirmRestore] = useState(false);
  const dropboxReady = isDropboxConfigured();
  return (
    <div className="space-y-6 max-w-2xl">
      <Card>
        <SectionTitle>Perfil</SectionTitle>
        <div className="space-y-4">
          <div>
            <FieldLabel>Nome</FieldLabel>
            <input value={name} onChange={(e) => setName(e.target.value)} onBlur={() => onChangeName(name)} className={inputClass} style={inputStyle} />
          </div>
          <div>
            <FieldLabel>E-mail</FieldLabel>
            <input type="email" placeholder="seu@email.com" className={inputClass} style={inputStyle} />
          </div>
        </div>
      </Card>

      <Card>
        <SectionTitle>Aparência</SectionTitle>
        <div className="space-y-4">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium" style={{ color: 'var(--text)' }}>Modo claro</p>
              <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Tema atual da aplicação</p>
            </div>
            <Badge color="var(--primary)" soft="var(--primary-soft)">Ativo</Badge>
          </div>
          <div className="flex items-center justify-between opacity-50">
            <div>
              <p className="text-sm font-medium" style={{ color: 'var(--text)' }}>Modo escuro</p>
              <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Em breve</p>
            </div>
            <ToggleSwitch checked={false} onChange={() => {}} />
          </div>
        </div>
      </Card>

      <Card>
        <SectionTitle>Metas gerais</SectionTitle>
        <div>
          <FieldLabel>Meta mensal de economia</FieldLabel>
          <CurrencyInput value={settings.monthlySavingsTarget || 0} onChange={(v) => onChangeSettings({ ...settings, monthlySavingsTarget: v })} />
          <p className="text-xs mt-1" style={{ color: 'var(--text-soft)' }}>Usado no card "Meta mensal" do dashboard (receitas − despesas do mês). Deixe 0 para ocultar a meta.</p>
        </div>
      </Card>

      <Card>
        <SectionTitle>Notificações</SectionTitle>
        <div className="space-y-4">
          {[
            { key: 'notifyDueDates', label: 'Vencimentos e faturas', desc: 'Avisos sobre boletos e cartões' },
            { key: 'notifyGoals', label: 'Progresso de metas', desc: 'Quando você se aproxima de uma meta' },
            { key: 'notifyInsights', label: 'Insights inteligentes', desc: 'Recomendações automáticas' },
          ].map((item) => (
            <div key={item.key} className="flex items-center justify-between">
              <div>
                <p className="text-sm font-medium" style={{ color: 'var(--text)' }}>{item.label}</p>
                <p className="text-xs" style={{ color: 'var(--text-soft)' }}>{item.desc}</p>
              </div>
              <ToggleSwitch checked={settings[item.key] !== false} onChange={(v) => onChangeSettings({ ...settings, [item.key]: v })} />
            </div>
          ))}
        </div>
      </Card>

      <Card>
        <SectionTitle>Backup na nuvem</SectionTitle>
        {!dropboxReady ? (
          <p className="text-sm" style={{ color: 'var(--text-soft)' }}>
            Backup automático via Dropbox ainda não configurado neste app (falta o App key). Veja o README do projeto.
          </p>
        ) : !dropboxConnected ? (
          <>
            <p className="text-sm mb-4" style={{ color: 'var(--text-soft)' }}>
              Conecte sua conta Dropbox para manter backups automáticos dos seus dados. A cada alteração, um arquivo é salvo — os 5 mais recentes ficam guardados, os antigos são apagados sozinhos.
            </p>
            <Button icon={Cloud} onClick={onConnectDropbox}>Conectar Dropbox</Button>
          </>
        ) : (
          <div className="space-y-4">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Badge color={dropboxSyncError ? 'var(--expense)' : 'var(--income)'} soft={dropboxSyncError ? 'var(--expense-soft)' : 'var(--income-soft)'}>
                  {dropboxSyncError ? 'Falha no backup' : 'Conectado'}
                </Badge>
                <span className="text-xs" style={{ color: 'var(--text-soft)' }}>
                  {dropboxLastBackup ? `Último backup: ${dropboxLastBackup.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}` : 'Backup automático a cada alteração'}
                </span>
              </div>
              <button onClick={onDisconnectDropbox} className="text-xs font-medium" style={{ color: 'var(--expense)' }}>Desconectar</button>
            </div>
            {dropboxSyncError && (
              <p className="text-xs flex items-start gap-1.5" style={{ color: 'var(--expense)' }}>
                <AlertCircle size={13} className="mt-0.5 shrink-0" />
                O último backup automático falhou. Pode ser conexão instável ou o acesso ao Dropbox ter expirado — tente "Fazer backup agora" ou desconecte e conecte de novo.
              </p>
            )}
            <div className="flex flex-wrap gap-3">
              <Button variant="secondary" size="sm" onClick={onBackupNow} disabled={dropboxBusy}>{dropboxBusy ? 'Enviando...' : 'Fazer backup agora'}</Button>
              <Button variant="secondary" size="sm" onClick={() => setShowConfirmRestore(true)} disabled={dropboxBusy}>Restaurar último backup</Button>
            </div>
          </div>
        )}
      </Card>

      <Card>
        <SectionTitle>Dados</SectionTitle>
        <p className="text-sm mb-4" style={{ color: 'var(--text-soft)' }}>Seus dados ficam salvos automaticamente neste aplicativo. Você pode restaurar os dados de exemplo a qualquer momento.</p>
        <Button variant="secondary" onClick={() => setShowConfirmReset(true)}>Restaurar dados de exemplo</Button>
      </Card>

      {showConfirmReset && (
        <ConfirmModal
          title="Restaurar dados de exemplo"
          description="Isso substituirá suas transações, contas, cartões, metas, caixinhas e despesas recorrentes atuais pelos dados de demonstração. Essa ação não pode ser desfeita."
          onConfirm={onReset}
          onClose={() => setShowConfirmReset(false)}
        />
      )}
      {showConfirmRestore && (
        <ConfirmModal
          title="Restaurar último backup do Dropbox"
          description="Isso substituirá os dados deste navegador pelo backup mais recente salvo no Dropbox. Essa ação não pode ser desfeita."
          onConfirm={() => { onRestoreFromDropbox(); setShowConfirmRestore(false); }}
          onClose={() => setShowConfirmRestore(false)}
        />
      )}
    </div>
  );
}

/* ============================================================
   APLICAÇÃO PRINCIPAL
   ============================================================ */

export default function App() {
  const [isLoading, setIsLoading] = useState(true);
  const [activePage, setActivePage] = useState('dashboard');
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [period, setPeriod] = useState('mes');
  const [customRange, setCustomRange] = useState({ start: '', end: '' });
  const [search, setSearch] = useState('');
  const [userName, setUserName] = useState('Thiago Coura');
  const [settings, setSettings] = useState({ notifyDueDates: true, notifyGoals: true, notifyInsights: true, monthlySavingsTarget: 2500 });

  const [transactions, setTransactions] = useState([]);
  const [accounts, setAccounts] = useState([]);
  const [cards, setCards] = useState([]);
  const [goals, setGoals] = useState([]);
  const [caixinhas, setCaixinhas] = useState([]);
  const [recurring, setRecurring] = useState([]);
  const [benefits, setBenefits] = useState([]);
  const [investments, setInvestments] = useState([]);

  const [modal, setModal] = useState(null);
  const [toasts, setToasts] = useState([]);
  const [bannerDismissed, setBannerDismissed] = useState(false);

  const [dropboxConnected, setDropboxConnected] = useState(isDropboxConnected());
  const [dropboxBusy, setDropboxBusy] = useState(false);
  const [dropboxLastBackup, setDropboxLastBackup] = useState(null);
  const [dropboxSyncError, setDropboxSyncError] = useState(false);

  function applyLoadedData(loaded) {
    setTransactions(loaded.transactions || buildInitialTransactions());
    setAccounts(loaded.accounts || initialAccounts);
    setCards(loaded.cards || initialCards);
    setGoals(loaded.goals || initialGoals);
    setCaixinhas(loaded.caixinhas || initialCaixinhas);
    setRecurring(loaded.recurring || initialRecurring);
    setBenefits(loaded.benefits || initialBenefits);
    setInvestments(loaded.investments || initialInvestments);
    setUserName(loaded.userName || 'Thiago Coura');
    setSettings(loaded.settings || { notifyDueDates: true, notifyGoals: true, notifyInsights: true, monthlySavingsTarget: 2500 });
  }

  useEffect(() => {
    (async () => {
      // Se voltamos de um login do Dropbox, finaliza a troca do código por token.
      const redirectResult = await handleDropboxRedirect().catch((e) => ({ status: 'error', message: e?.message }));
      if (redirectResult.status === 'connected') {
        setDropboxConnected(true);
        addToastSafe('Dropbox conectado com sucesso.');
      } else if (redirectResult.status === 'error') {
        addToastSafe(`Não foi possível conectar ao Dropbox: ${redirectResult.message || 'tente novamente.'}`, 'error');
      }

      const loaded = await loadAppData();
      if (loaded) {
        applyLoadedData(loaded);
      } else if (isDropboxConnected()) {
        // Sem dados neste navegador, mas com Dropbox conectado: tenta puxar o backup mais recente
        // (cenário de "abrir em outro aparelho e continuar de onde parou").
        try {
          const remote = await downloadLatestBackup();
          if (remote) {
            applyLoadedData(remote.data);
            persistImmediate(remote.data);
            addToastSafe(`Dados restaurados do backup do Dropbox (${remote.filename}).`);
          } else {
            applySampleData();
          }
        } catch {
          applySampleData();
        }
      } else {
        applySampleData();
      }
      setIsLoading(false);
    })();

    function applySampleData() {
      setTransactions(buildInitialTransactions());
      setAccounts(initialAccounts);
      setCards(initialCards);
      setGoals(initialGoals);
      setCaixinhas(initialCaixinhas);
      setRecurring(initialRecurring);
      setBenefits(initialBenefits);
      setInvestments(initialInvestments);
    }
    function persistImmediate(d) { saveAppData(d); }
    function addToastSafe(message, type = 'success') { setToasts((prev) => [...prev, { id: uid(), message, type }]); }
  }, []);

  function addToast(message, type = 'success') {
    const id = uid();
    setToasts((prev) => [...prev, { id, message, type }]);
    setTimeout(() => setToasts((prev) => prev.filter((t) => t.id !== id)), 3500);
  }

  function persist(partial) {
    const fullState = {
      transactions: partial.transactions ?? transactions,
      accounts: partial.accounts ?? accounts,
      cards: partial.cards ?? cards,
      goals: partial.goals ?? goals,
      caixinhas: partial.caixinhas ?? caixinhas,
      recurring: partial.recurring ?? recurring,
      benefits: partial.benefits ?? benefits,
      investments: partial.investments ?? investments,
      userName: partial.userName ?? userName,
      settings: partial.settings ?? settings,
    };
    saveAppData(fullState);
    scheduleBackup(() => fullState, (err, info) => {
      if (err) { setDropboxSyncError(true); }
      else if (info) { setDropboxSyncError(false); setDropboxLastBackup(new Date()); }
    });
  }

  async function connectDropbox() {
    await startDropboxConnect();
  }
  function disconnectDropboxAccount() {
    disconnectDropbox();
    setDropboxConnected(false);
    setDropboxLastBackup(null);
    addToast('Dropbox desconectado. Seus dados continuam salvos neste navegador.');
  }
  async function backupNowToDropbox() {
    setDropboxBusy(true);
    try {
      await uploadBackup(JSON.stringify({ transactions, accounts, cards, goals, caixinhas, recurring, benefits, investments, userName, settings }));
      setDropboxSyncError(false);
      setDropboxLastBackup(new Date());
      addToast('Backup enviado para o Dropbox.');
    } catch (e) {
      setDropboxSyncError(true);
      addToast('Não foi possível enviar o backup. Verifique sua conexão com o Dropbox.', 'error');
    } finally {
      setDropboxBusy(false);
    }
  }
  async function restoreFromDropbox() {
    setDropboxBusy(true);
    try {
      const remote = await downloadLatestBackup();
      if (!remote) { addToast('Nenhum backup encontrado no Dropbox ainda.', 'error'); return; }
      applyLoadedData(remote.data);
      saveAppData(remote.data);
      addToast(`Dados restaurados do backup "${remote.filename}".`);
    } catch (e) {
      addToast('Não foi possível restaurar o backup do Dropbox.', 'error');
    } finally {
      setDropboxBusy(false);
    }
  }

  function goTo(page) { setActivePage(page); window.scrollTo({ top: 0 }); }

  /* ---- transações ---- */
  function addTransaction(form) {
    const newTx = { ...form, id: uid() };
    const updated = [newTx, ...transactions];
    const updatedAccounts = reapplyAccountEffect(accounts, null, newTx);
    setTransactions(updated); setAccounts(updatedAccounts); persist({ transactions: updated, accounts: updatedAccounts });
    addToast('Lançamento adicionado com sucesso.');
  }
  function editTransaction(form) {
    const oldTx = transactions.find((t) => t.id === form.id);
    const updated = transactions.map((t) => (t.id === form.id ? form : t));
    const updatedAccounts = reapplyAccountEffect(accounts, oldTx, form);
    setTransactions(updated); setAccounts(updatedAccounts); persist({ transactions: updated, accounts: updatedAccounts });
    addToast('Lançamento atualizado.');
  }
  function deleteTransaction(tx) {
    const updated = transactions.filter((t) => t.id !== tx.id);
    const updatedAccounts = reapplyAccountEffect(accounts, tx, null);
    setTransactions(updated); setAccounts(updatedAccounts); persist({ transactions: updated, accounts: updatedAccounts });
    addToast('Lançamento removido.');
  }
  function importTransactions(rows, meta = {}) {
    if (rows.length === 0) {
      if (meta.error) addToast('Não foi possível ler o arquivo. Verifique se é um CSV de fatura do Nubank.', 'error');
      else addToast('Nenhum lançamento novo para importar.', 'error');
      return;
    }
    const updated = [...rows, ...transactions];
    let updatedAccounts = accounts;
    rows.forEach((r) => { updatedAccounts = reapplyAccountEffect(updatedAccounts, null, r); });
    setTransactions(updated); setAccounts(updatedAccounts); persist({ transactions: updated, accounts: updatedAccounts });
    let message = `${rows.length} lançamento(s) importado(s) com sucesso.`;
    if (meta.duplicates) message += ` ${meta.duplicates} duplicata(s) ignorada(s).`;
    if (meta.payments) message += ` ${meta.payments} adiantamento(s) de fatura (${formatBRL(meta.paymentsTotal || 0)}) não importado(s).`;
    addToast(message);
  }

  /* ---- contas ---- */
  function addAccount(form) {
    const updated = [...accounts, { ...form, id: uid() }];
    setAccounts(updated); persist({ accounts: updated });
    addToast('Conta adicionada.');
  }
  function deleteAccount(acc) {
    const updated = accounts.filter((a) => a.id !== acc.id);
    setAccounts(updated); persist({ accounts: updated });
    addToast('Conta removida.');
  }
  function setAccountThreshold(acc, value) {
    const updated = accounts.map((a) => (a.id === acc.id ? { ...a, alertThreshold: value } : a));
    setAccounts(updated); persist({ accounts: updated });
    addToast(value > 0 ? `Alerta de saldo baixo definido para ${formatBRL(value)}.` : 'Alerta de saldo baixo desativado.');
  }
  function setAccountBalance(acc, value) {
    const updated = accounts.map((a) => (a.id === acc.id ? { ...a, balance: value } : a));
    setAccounts(updated); persist({ accounts: updated });
    addToast('Saldo ajustado.');
  }

  /* ---- caixinhas ---- */
  function addCaixinha(form) {
    const updated = [...caixinhas, { ...form, id: uid() }];
    setCaixinhas(updated); persist({ caixinhas: updated });
    addToast('Caixinha criada.');
  }
  function deleteCaixinha(cx) {
    const updated = caixinhas.filter((c) => c.id !== cx.id);
    setCaixinhas(updated); persist({ caixinhas: updated });
    addToast('Caixinha removida.');
  }
  function addCaixinhaFunds(cx, amount) {
    const updated = caixinhas.map((c) => (c.id === cx.id ? { ...c, balance: c.balance + amount } : c));
    setCaixinhas(updated); persist({ caixinhas: updated });
    addToast(`${formatBRL(amount)} guardado(s) em "${cx.name}".`);
  }

  /* ---- cartões ---- */
  function addCard(form) {
    const updated = [...cards, { ...form, id: uid() }];
    setCards(updated); persist({ cards: updated });
    addToast('Cartão adicionado.');
  }
  function deleteCard(card) {
    const updated = cards.filter((c) => c.id !== card.id);
    setCards(updated); persist({ cards: updated });
    addToast('Cartão removido.');
  }

  /* ---- vale-benefícios ---- */
  function addBenefit(form) {
    const updated = [...benefits, { ...form, id: uid() }];
    setBenefits(updated); persist({ benefits: updated });
    addToast('Vale-benefícios adicionado.');
  }
  function deleteBenefit(b) {
    const updated = benefits.filter((x) => x.id !== b.id);
    setBenefits(updated); persist({ benefits: updated });
    addToast('Vale-benefícios removido.');
  }
  function adjustBenefit(b, field, delta) {
    const updated = benefits.map((x) => (x.id === b.id ? { ...x, [field]: Math.max(0, x[field] + delta) } : x));
    setBenefits(updated); persist({ benefits: updated });
    addToast(delta > 0 ? `${formatBRL(delta)} recarregado(s).` : `${formatBRL(Math.abs(delta))} registrado(s) como uso.`);
  }

  /* ---- metas ---- */
  function addGoal(form) {
    const today = new Date().toISOString().slice(0, 10);
    const history = form.current > 0 ? [{ date: today, amount: form.current }] : [];
    const updated = [...goals, { ...form, id: uid(), history }];
    setGoals(updated); persist({ goals: updated });
    addToast('Meta criada.');
  }
  function deleteGoal(goal) {
    const updated = goals.filter((g) => g.id !== goal.id);
    setGoals(updated); persist({ goals: updated });
    addToast('Meta removida.');
  }
  function addGoalFunds(goal, amount) {
    const today = new Date().toISOString().slice(0, 10);
    const updated = goals.map((g) => (g.id === goal.id ? { ...g, current: g.current + amount, history: [...(g.history || []), { date: today, amount }] } : g));
    setGoals(updated); persist({ goals: updated });
    addToast(`${formatBRL(amount)} adicionado(s) à meta "${goal.name}".`);
  }

  /* ---- despesas recorrentes ---- */
  function addRecurring(form) {
    const updated = [...recurring, { ...form, id: uid() }];
    setRecurring(updated); persist({ recurring: updated });
    addToast('Despesa recorrente adicionada.');
  }
  function deleteRecurring(r) {
    const updated = recurring.filter((x) => x.id !== r.id);
    setRecurring(updated); persist({ recurring: updated });
    addToast('Despesa recorrente removida.');
  }
  // Cria um lançamento real de hoje a partir de uma despesa recorrente — elas não entram
  // sozinhas nos totais, isso é o que efetivamente conta a despesa no seu saldo/relatórios.
  function addRecurringAsTransaction(item, accountId) {
    addTransaction({
      description: item.name, amount: item.value, category: item.category, type: 'despesa',
      account: accountId, paymentMethod: 'Não informado', date: new Date().toISOString().slice(0, 10), status: 'Pago',
    });
  }

  /* ---- investimentos ---- */
  function addInvestment(form) {
    const updated = [...investments, { ...form, id: uid() }];
    setInvestments(updated); persist({ investments: updated });
    addToast('Investimento adicionado.');
  }
  function editInvestment(form) {
    const updated = investments.map((i) => (i.id === form.id ? form : i));
    setInvestments(updated); persist({ investments: updated });
    addToast('Investimento atualizado.');
  }
  function deleteInvestment(inv) {
    const updated = investments.filter((i) => i.id !== inv.id);
    setInvestments(updated); persist({ investments: updated });
    addToast('Investimento removido.');
  }

  /* ---- configurações ---- */
  function changeName(name) {
    setUserName(name); persist({ userName: name });
  }
  function changeSettings(newSettings) {
    setSettings(newSettings); persist({ settings: newSettings });
  }
  function resetToSampleData() {
    const t = buildInitialTransactions();
    setTransactions(t); setAccounts(initialAccounts); setCards(initialCards);
    setGoals(initialGoals); setCaixinhas(initialCaixinhas); setRecurring(initialRecurring); setBenefits(initialBenefits); setInvestments(initialInvestments);
    persist({ transactions: t, accounts: initialAccounts, cards: initialCards, goals: initialGoals, caixinhas: initialCaixinhas, recurring: initialRecurring, benefits: initialBenefits, investments: initialInvestments });
    addToast('Dados de exemplo restaurados.');
  }

  const investmentsTotal = useMemo(() => investments.reduce((s, i) => s + i.currentValue, 0), [investments]);
  const realMonthlyHistory = useMemo(
    () => computeMonthlyHistory(transactions, accounts, investmentsTotal),
    [transactions, accounts, investmentsTotal]
  );
  const realCategoryComparison = useMemo(() => {
    const now = new Date();
    const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    return {
      current: computeCategoryTotals(transactions, now.getFullYear(), now.getMonth()),
      previous: computeCategoryTotals(transactions, prev.getFullYear(), prev.getMonth()),
    };
  }, [transactions]);

  const insights = useMemo(() => {
    if (isLoading) return [];
    return generateInsights({ transactions, goals, monthlyHistory: realMonthlyHistory, accounts, categoryComparison: realCategoryComparison });
  }, [isLoading, transactions, goals, accounts, realMonthlyHistory, realCategoryComparison]);

  const filteredForSearch = useMemo(() => {
    if (!search.trim()) return null;
    const q = search.toLowerCase();
    return transactions.filter((t) => t.description.toLowerCase().includes(q) || t.category.toLowerCase().includes(q));
  }, [search, transactions]);

  const data = {
    transactions, accounts, cards, goals, caixinhas, recurring, benefits, investments, settings, period, customRange,
    monthlyHistory: realMonthlyHistory, categoryComparison: realCategoryComparison, investmentsTotal, insights,
  };
  const actions = {
    goTo, editTransaction, deleteTransaction, addGoalFunds, deleteGoal,
    addInvestment, editInvestment, deleteInvestment, addRecurringAsTransaction,
  };

  if (isLoading) {
    return (
      <div className="cerne-root flex h-screen" style={{ backgroundColor: 'var(--bg)' }}>
        <style>{GLOBAL_STYLES}</style>
        <LoadingScreen />
      </div>
    );
  }

  return (
    <div className="cerne-root flex h-screen overflow-hidden" style={{ backgroundColor: 'var(--bg)' }}>
      <style>{GLOBAL_STYLES}</style>
      <Sidebar activePage={activePage} setActivePage={goTo} sidebarOpen={sidebarOpen} setSidebarOpen={setSidebarOpen} onNewTransaction={() => setModal({ type: 'newTransaction' })} userName={userName} />

      <div className="flex-1 flex flex-col overflow-hidden">
        <Header userName={userName} period={period} setPeriod={setPeriod} customRange={customRange} setCustomRange={setCustomRange} search={search} setSearch={setSearch} setSidebarOpen={setSidebarOpen} />
        {!bannerDismissed && insights[0] && <Banner insight={insights[0]} onDismiss={() => setBannerDismissed(true)} />}

        <main className="flex-1 overflow-y-auto px-4 md:px-6 lg:px-8 py-6 print-area">
          {search.trim() && activePage === 'dashboard' ? (
            <Card>
              <SectionTitle>Resultados para "{search}"</SectionTitle>
              {filteredForSearch.length === 0 ? (
                <EmptyState icon={Search} title="Nenhum resultado encontrado" description="Tente buscar por outra descrição ou categoria." />
              ) : (
                <div className="space-y-2">
                  {filteredForSearch.slice(0, 10).map((tx) => (
                    <div key={tx.id} className="flex items-center justify-between text-sm p-2.5 rounded-xl hover:bg-black/[0.02]">
                      <div className="flex items-center gap-3 min-w-0">
                        <span style={{ color: 'var(--text-soft)' }}>{formatDate(tx.date)}</span>
                        <span className="truncate" style={{ color: 'var(--text)' }}>{tx.description}</span>
                        <CategoryBadge category={tx.category} />
                      </div>
                      <span className="tabular-nums font-medium" style={{ color: tx.type === 'receita' ? 'var(--income)' : 'var(--expense)' }}>{formatBRL(tx.amount)}</span>
                    </div>
                  ))}
                </div>
              )}
            </Card>
          ) : (
            <>
              {activePage === 'dashboard' && <DashboardPage data={data} actions={actions} />}
              {activePage === 'transacoes' && <TransactionsPage filterType="all" title="Transações" transactions={transactions} accounts={accounts} onAdd={addTransaction} onEdit={editTransaction} onDelete={deleteTransaction} onImport={importTransactions} />}
              {activePage === 'receitas' && <TransactionsPage filterType="receita" title="Receitas" transactions={transactions} accounts={accounts} onAdd={addTransaction} onEdit={editTransaction} onDelete={deleteTransaction} onImport={importTransactions} />}
              {activePage === 'despesas' && <TransactionsPage filterType="despesa" title="Despesas" transactions={transactions} accounts={accounts} onAdd={addTransaction} onEdit={editTransaction} onDelete={deleteTransaction} onImport={importTransactions} />}
              {activePage === 'contas' && <AccountsPage accounts={accounts} caixinhas={caixinhas} transactions={transactions} onAddAccount={addAccount} onDeleteAccount={deleteAccount} onSetAccountThreshold={setAccountThreshold} onSetAccountBalance={setAccountBalance} onAddCaixinha={addCaixinha} onDeleteCaixinha={deleteCaixinha} onAddCaixinhaFunds={addCaixinhaFunds} />}
              {activePage === 'cartoes' && <CardsPage cards={cards} onAdd={addCard} onDelete={deleteCard} benefits={benefits} onAddBenefit={addBenefit} onDeleteBenefit={deleteBenefit} onAdjustBenefit={adjustBenefit} />}
              {activePage === 'investimentos' && <InvestmentsPage investments={investments} onAdd={addInvestment} onEdit={editInvestment} onDelete={deleteInvestment} />}
              {activePage === 'metas' && <GoalsPage goals={goals} onAdd={addGoal} onAddFunds={addGoalFunds} onDelete={deleteGoal} />}
              {activePage === 'recorrentes' && <RecurringExpensesPage recurring={recurring} accounts={accounts} onAdd={addRecurring} onDelete={deleteRecurring} onLaunchNow={addRecurringAsTransaction} />}
              {activePage === 'relatorios' && <ReportsPage monthlyHistory={data.monthlyHistory} categoryComparison={data.categoryComparison} />}
              {activePage === 'configuracoes' && (
                <SettingsPage
                  userName={userName} settings={settings} onChangeName={changeName} onChangeSettings={changeSettings} onReset={resetToSampleData}
                  dropboxConnected={dropboxConnected} dropboxBusy={dropboxBusy} dropboxLastBackup={dropboxLastBackup} dropboxSyncError={dropboxSyncError}
                  onConnectDropbox={connectDropbox} onDisconnectDropbox={disconnectDropboxAccount}
                  onBackupNow={backupNowToDropbox} onRestoreFromDropbox={restoreFromDropbox}
                />
              )}
            </>
          )}
        </main>
      </div>

      {modal?.type === 'newTransaction' && (
        <TransactionForm accounts={accounts} onSave={(f) => { addTransaction(f); setModal(null); }} onClose={() => setModal(null)} />
      )}

      <ToastContainer toasts={toasts} />
    </div>
  );
}
