import React, { useState, useEffect, useMemo, useRef } from 'react';
import { createPortal } from 'react-dom';
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
  MessageCircle, Moon, Sun, Shirt, Bike, Building2, EyeOff, PersonStanding, Copy, Smartphone, MoreVertical,
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

.cerne-root.dark {
  --primary: #8FA680;
  --primary-dark: #7A9068;
  --primary-soft: #2E3A28;
  --secondary: #8A9B7D;
  --bg: #0D1117;
  --card: #161B22;
  --border: #30363D;
  --text: #E6EDF3;
  --text-soft: #8B949E;
  --income: #7FB37F;
  --income-soft: #24352A;
  --expense: #D98888;
  --expense-soft: #3A2626;
  --invest: #8FA0C9;
  --invest-soft: #262B3A;
  --goals: #C9B896;
  --goals-soft: #332E23;
  --alert: #E6BC6C;
  --alert-soft: #3A2F1A;
}
.cerne-root.dark .shadow-soft,
.cerne-root.dark .shadow-soft-lg { box-shadow: 0 1px 2px rgba(0,0,0,0.25), 0 8px 20px rgba(0,0,0,0.3); }
.cerne-root.dark .skeleton { background: linear-gradient(90deg, #161B22 25%, #21262D 37%, #161B22 63%); background-size: 400px 100%; }
/* Os overlays de hover foram pensados pra fundo claro (escurecem levemente) — no escuro, viram
   quase invisíveis. Troca pra um overlay branco sutil só dentro do tema escuro. */
.cerne-root.dark .hover\:bg-black\/5:hover,
.cerne-root.dark .hover\:bg-black\/10:hover { background-color: rgba(255,255,255,0.08) !important; }
.cerne-root.dark .bg-black\/\[0\.02\] { background-color: rgba(255,255,255,0.03) !important; }
.cerne-root.dark .bg-black\/\[0\.03\] { background-color: rgba(255,255,255,0.04) !important; }

/* Temas de cor de destaque (opcionais, escolhidos em Configurações → Aparência).
   Cada um só troca --primary / --primary-dark / --primary-soft; o resto dos tokens
   (fundo, texto, bordas, cores semânticas) continua vindo de :root / .dark acima. */
.cerne-root.theme-azul { --primary: #4A6FA5; --primary-dark: #3B5A87; --primary-soft: #E6ECF3; }
.cerne-root.theme-azul.dark { --primary: #7CA3D6; --primary-dark: #6690C7; --primary-soft: #26303F; }

.cerne-root.theme-rosa { --primary: #B5677E; --primary-dark: #9A5268; --primary-soft: #F5E9EC; }
.cerne-root.theme-rosa.dark { --primary: #D98CA0; --primary-dark: #C4728A; --primary-soft: #3A2830; }

.cerne-root.theme-vermelho { --primary: #A85D4E; --primary-dark: #8E4B3E; --primary-soft: #F3E7E3; }
.cerne-root.theme-vermelho.dark { --primary: #CC8271; --primary-dark: #B96D5C; --primary-soft: #3A2A24; }

.cerne-root.theme-branco { --primary: #3A3A3A; --primary-dark: #2A2A2A; --primary-soft: #EDEDEC; }
.cerne-root.theme-branco.dark { --primary: #C9C9C9; --primary-dark: #B5B5B5; --primary-soft: #333333; }

.cerne-root.theme-ameixa { --primary: #7D5A82; --primary-dark: #664867; --primary-soft: #EFE7F0; }
.cerne-root.theme-ameixa.dark { --primary: #B08AB8; --primary-dark: #9B76A3; --primary-soft: #332B36; }

.cerne-root, .cerne-root * { font-family: 'Plus Jakarta Sans', ui-sans-serif, system-ui, sans-serif; box-sizing: border-box; }
.font-display { font-family: 'Space Grotesk', ui-sans-serif, system-ui, sans-serif; }

.cerne-root ::-webkit-scrollbar { width: 8px; height: 8px; }
.cerne-root ::-webkit-scrollbar-track { background: transparent; }
.cerne-root ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 8px; }
.cerne-root ::-webkit-scrollbar-thumb:hover { background: var(--secondary); }

.shadow-soft { box-shadow: 0 1px 2px rgba(35,35,35,0.04), 0 4px 14px rgba(35,35,35,0.05); }
.shadow-soft-lg { box-shadow: 0 2px 6px rgba(35,35,35,0.05), 0 14px 28px rgba(35,35,35,0.08); }

.focus-ring:focus { outline: none; box-shadow: 0 0 0 3px color-mix(in srgb, var(--primary) 30%, transparent); }
.focus-ring:focus-visible { outline: none; box-shadow: 0 0 0 3px color-mix(in srgb, var(--primary) 30%, transparent); }

@keyframes fadeSlideUp { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
.animate-fade-up { animation: fadeSlideUp 0.4s cubic-bezier(0.16,1,0.3,1) both; }

@keyframes toastIn { from { opacity: 0; transform: translateY(10px) scale(0.98); } to { opacity: 1; transform: translateY(0) scale(1); } }
.animate-toast { animation: toastIn 0.25s ease-out both; }

@keyframes ringPulse { 0% { transform: scale(0.85); opacity: 0.6; } 100% { transform: scale(1.6); opacity: 0; } }
.ring-pulse { animation: ringPulse 1.8s ease-out infinite; }

@keyframes shimmer { 0% { background-position: -200px 0; } 100% { background-position: 200px 0; } }
.skeleton { background: linear-gradient(90deg, #EDEAE2 25%, #F7F5EF 37%, #EDEAE2 63%); background-size: 400px 100%; animation: shimmer 1.4s ease-in-out infinite; }

@keyframes shakeError { 10%, 90% { transform: translateX(-1px); } 20%, 80% { transform: translateX(2px); } 30%, 50%, 70% { transform: translateX(-4px); } 40%, 60% { transform: translateX(4px); } }
.animate-shake { animation: shakeError 0.4s cubic-bezier(0.36,0.07,0.19,0.97) both; }

@keyframes fieldFlash { 0%, 100% { box-shadow: 0 0 0 0 rgba(182,107,107,0); } 15%, 45% { box-shadow: 0 0 0 3px rgba(182,107,107,0.35); } }
.animate-field-flash { animation: fieldFlash 1.1s ease-out both; }

/* Botão flutuante de novo lançamento: ao encolher/expandir (rolagem), gira uma volta
   rápida seguida de uma segunda volta mais lenta (0%→35% cobre os primeiros 360°,
   35%→100% cobre os últimos 360° num intervalo bem maior de tempo). Ao voltar ao
   tamanho normal, gira meia volta (180°) no sentido CONTRÁRIO ao do encolhimento —
   como um beliscão/elástico voltando, em vez de continuar girando pra sempre no
   mesmo sentido.  */
@keyframes fabSpinIn {
  0%   { transform: scale(1) rotate(0deg); }
  35%  { transform: scale(0.7) rotate(360deg); }
  100% { transform: scale(0.5) rotate(720deg); }
}
@keyframes fabSpinOut {
  0%   { transform: scale(0.5) rotate(0deg); }
  100% { transform: scale(1) rotate(-180deg); }
}
.fab-spin-in { animation: fabSpinIn 0.7s cubic-bezier(0.4, 0, 0.2, 1) both; }
.fab-spin-out { animation: fabSpinOut 0.4s cubic-bezier(0.34, 1.56, 0.64, 1) both; }

/* Explosão sutil de emojis ao completar uma meta — burst rápido inicial, depois flutuação lenta
   e fade out. --tx/--ty/--dur/--delay são definidos por partícula via style inline. */
@keyframes goalCelebrateParticle {
  0%   { transform: translate(-50%, -50%) translate(0, 0) scale(0.4); opacity: 0; }
  15%  { transform: translate(-50%, -50%) translate(calc(var(--tx) * 0.55), calc(var(--ty) * 0.55)) scale(1.2); opacity: 1; }
  70%  { transform: translate(-50%, -50%) translate(var(--tx), calc(var(--ty) - 36px)) scale(1); opacity: 1; }
  100% { transform: translate(-50%, -50%) translate(var(--tx), calc(var(--ty) - 64px)) scale(0.85); opacity: 0; }
}
.goal-celebrate-particle {
  position: fixed; left: var(--origin-x); top: var(--origin-y); font-size: var(--size);
  animation: goalCelebrateParticle var(--dur) var(--delay) cubic-bezier(0.16,0.85,0.3,1) both;
  pointer-events: none; will-change: transform, opacity; z-index: 60;
}

@media print {
  .no-print { display: none !important; }
  .print-area { padding: 0 !important; overflow: visible !important; }
}

/* Remove o contorno de foco padrão do navegador ao tocar numa fatia do gráfico de pizza —
   sem isso, toques no celular desenham um retângulo preto ao redor da fatia selecionada. */
.recharts-wrapper svg *:focus,
.recharts-pie-sector,
.recharts-sector { outline: none !important; }
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
  'Telefonia': { color: '#6B8FB0', soft: '#E9EFF4', icon: Smartphone },
  'Outros': { color: '#A8A398', soft: '#F1EFEA', icon: MoreHorizontal },
};
const CATEGORY_NAMES = Object.keys(CATEGORIES);

const ACCOUNTS_ICONS = { 'Conta Corrente': Wallet, 'Poupança': PiggyBank };

const PAYMENT_METHODS = ['Pix', 'Cartão de crédito', 'Cartão de débito', 'Dinheiro', 'Transferência', 'Boleto', 'Empréstimo / Pix parcelado', 'Vale-benefícios'];
const STATUS_OPTIONS = ['Pago', 'Pendente'];
// O rótulo "Pago" não faz muito sentido pra uma receita (você recebe, não paga) — o valor
// gravado continua sendo 'Pago' (é o que toda a lógica de saldo/fatura usa), só o texto muda.
function statusLabel(status, type) {
  if (status === 'Pago' && type === 'receita') return 'Recebido';
  return status;
}

/* ---------- Helpers ---------- */

function formatBRL(value) {
  return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value || 0);
}
function formatDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr + 'T00:00:00');
  return d.toLocaleDateString('pt-BR');
}
// Compras parceladas importadas da fatura ANTES do ajuste no fluxo de importação guardavam a
// parcela como texto solto dentro da descrição, em um de dois formatos: "(parcela 3/3)" (o que o
// app passou a escrever) ou "- Parcela 3/3" / "- 3/3" (o formato bruto que vem do extrato do
// Nubank, caso tenha ficado assim por algum motivo) — sem preencher os campos estruturados
// (installmentGroupId/Index/Count) que a tag depende. Esse helper reconhece os dois formatos como
// fallback, pra esses lançamentos antigos também mostrarem a tag corretamente sem reimportar nada.
function parseInstallmentFallback(description) {
  const patterns = [
    /\s*\(parcela\s*(\d+)\s*\/\s*(\d+)\)\s*$/i,
    /\s*-\s*parcela\s+(\d+)\s*\/\s*(\d+)\s*$/i,
    /\s*-\s*(\d+)\s*\/\s*(\d+)\s*$/,
  ];
  for (const re of patterns) {
    const m = description.match(re);
    if (m) return { clean: description.slice(0, m.index).trim(), index: Number(m[1]), count: Number(m[2]) };
  }
  return null;
}
// Dado um lançamento, devolve { desc, index, count } prontos pra exibir — usa os campos
// estruturados quando existem (compra parcelada normal) ou o fallback acima (importação antiga).
function getInstallmentDisplay(t) {
  if (t.installmentGroupId) return { desc: t.description, index: t.installmentIndex, count: t.installmentCount };
  const fallback = parseInstallmentFallback(t.description);
  return fallback ? { desc: fallback.clean, index: fallback.index, count: fallback.count } : { desc: t.description, index: null, count: null };
}
// Mesma data, mas com o ano em 2 dígitos (ex: 23/08/27) — usada nos cards mobile onde a
// data divide espaço com a tag de parcela (3/12), pra sobrar mais respiro sem perder o ano.
function formatDateShortYear(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr + 'T00:00:00');
  return d.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: '2-digit' });
}
function formatDateLong(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  return d.toLocaleDateString('pt-BR', { day: 'numeric', month: 'long', year: 'numeric' });
}
function formatMonthYear(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  const label = d.toLocaleDateString('pt-BR', { month: 'short', year: '2-digit' });
  return label.charAt(0).toUpperCase() + label.slice(1).replace('.', '');
}
// Capitaliza só a primeira letra da frase — diferente do CSS `capitalize`, que capitaliza
// cada palavra e acaba maiusculizando indevidamente conectivos como "de" (ex.: "Terça-feira,
// 19 De Agosto"). Datas por extenso em pt-BR só devem ter a primeira letra maiúscula.
function capitalizeFirst(str) {
  if (!str) return str;
  return str.charAt(0).toUpperCase() + str.slice(1);
}
function relativeTime(date) {
  if (!date) return '';
  const diffMs = Date.now() - date.getTime();
  const mins = Math.round(diffMs / 60000);
  if (mins < 1) return 'agora mesmo';
  if (mins < 60) return `há ${mins} min`;
  const hours = Math.round(mins / 60);
  if (hours < 24) return `há ${hours}h`;
  const days = Math.round(hours / 24);
  return `há ${days}d`;
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

// Catálogo de blocos que o usuário pode ligar/desligar em Configurações → Personalização.
// Cada categoria também tem um "essentialLabel" só pra exibição (o que sempre fica visível ali),
// não é um item de fato — navegação, listas principais e formulários nunca entram nesse catálogo,
// então não tem como desativar sem querer algo indispensável.
const VISIBILITY_SCHEMA = [
  {
    key: 'dashboard', label: 'Dashboard', essentialLabel: 'Navegação principal',
    items: [
      { key: 'kpiSaldo', label: 'Card "Saldo disponível"' },
      { key: 'kpiReceitas', label: 'Card "Receitas do período"' },
      { key: 'kpiDespesas', label: 'Card "Despesas do período"' },
      { key: 'kpiEconomia', label: 'Card "Economia acumulada"' },
      { key: 'kpiMeta', label: 'Card "Meta de economia"' },
      { key: 'kpiPatrimonio', label: 'Card "Patrimônio total"' },
      { key: 'evolutionChart', label: 'Gráfico de evolução financeira' },
      { key: 'categoryDonut', label: 'Gastos por categoria (gráfico)' },
      { key: 'recentTransactions', label: 'Últimas transações' },
      { key: 'monthSummary', label: 'Resumo do mês' },
      { key: 'goalsSection', label: 'Metas financeiras' },
      { key: 'financialCalendar', label: 'Calendário financeiro' },
      { key: 'cardsPreview', label: 'Prévia dos cartões' },
      { key: 'recurringPreview', label: 'Prévia de despesas recorrentes' },
      { key: 'investmentsPreview', label: 'Prévia de investimentos' },
      { key: 'insights', label: 'Insights automáticos' },
    ],
  },
  {
    key: 'transacoes', label: 'Transações', essentialLabel: 'Busca, filtros, tabela e botão "Novo"',
    items: [
      { key: 'faturaBanner', label: 'Aviso com atalho para a Fatura mensal' },
      { key: 'statCards', label: 'Cards de totais (despesas e receitas do filtro atual)' },
    ],
  },
  {
    key: 'contas', label: 'Contas Bancárias', essentialLabel: 'Lista de contas',
    items: [
      { key: 'saldoTotalCard', label: 'Card de saldo total' },
      { key: 'caixinhas', label: 'Seção de caixinhas' },
    ],
  },
  {
    key: 'cartoes', label: 'Cartões', essentialLabel: 'Seus cartões e Fatura mensal',
    items: [
      { key: 'beneficios', label: 'Seção de vale-benefícios' },
    ],
  },
  {
    key: 'investimentos', label: 'Investimentos', essentialLabel: 'Lista de investimentos',
    items: [
      { key: 'kpiInvestido', label: 'Card "Valor investido"' },
      { key: 'kpiAtual', label: 'Card "Valor atual"' },
      { key: 'kpiRentabilidade', label: 'Card "Rentabilidade"' },
      { key: 'allocationChart', label: 'Gráfico de distribuição por categoria' },
    ],
  },
  {
    key: 'recorrentes', label: 'Despesas Recorrentes', essentialLabel: 'Lista de assinaturas e despesas fixas',
    items: [
      { key: 'kpiTotalMensal', label: 'Card "Total recorrente mensal"' },
      { key: 'kpiTotalAnual', label: 'Card "Total anual estimado"' },
      { key: 'categoryChart', label: 'Gráfico por categoria' },
    ],
  },
  {
    key: 'relatorios', label: 'Relatórios', essentialLabel: 'Exportar Excel e imprimir/PDF',
    items: [
      { key: 'categoryComparison', label: 'Comparativo por categoria (mês atual x anterior)' },
    ],
  },
];
// Item ausente em settings.visibility conta como visível — assim dados salvos antes dessa feature
// (ou de futuras adições ao catálogo) continuam mostrando tudo até o usuário desativar algo.
function isVisible(settings, page, item) {
  return settings?.visibility?.[page]?.[item] !== false;
}

/* ---------- Agregações reais (a partir dos lançamentos, contas etc.) ---------- */

// Efeito de um lançamento sobre o saldo da conta: só lançamentos "Pago" já saíram/entraram de fato.
// Uma despesa "no crédito" (com cardId) só debita a conta quando a fatura é paga — não no
// momento da compra. Receitas e despesas fora do cartão (Pix, débito, dinheiro...) continuam
// afetando o saldo na hora, como sempre.
function transactionBalanceEffect(t) {
  if (!t || t.status !== 'Pago') return 0;
  if (t.type === 'despesa' && t.cardId) return 0;
  return t.type === 'receita' ? t.amount : -t.amount;
}
// "Pago"/"Pendente" agora tem dois sentidos diferentes dependendo do lançamento: numa despesa em
// débito/conta, Pago significa que o dinheiro já saiu de fato; numa despesa no crédito, Pago
// significa que a FATURA daquele cartão já foi paga (o gasto em si já aconteceu desde a compra,
// só o pagamento da fatura é que fica pendente). Pra relatórios/gráficos (que respondem "quanto eu
// gastei nesse mês"), uma despesa no crédito sempre conta a partir da compra, esteja a fatura paga
// ou não — é isso que isRealized() representa, diferente do status bruto usado nos badges.
function isRealized(t) {
  if (!t) return false;
  if (t.isInvoicePayment) return false;
  if (t.type === 'despesa' && t.cardId) return true;
  return t.status === 'Pago';
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

// Uma despesa lançada num vale-benefícios (Alimentação/Transporte) desconta o saldo daquele
// benefício na hora — diferente do cartão de crédito, o vale já é um dinheiro pré-carregado,
// não existe "fatura" pra pagar depois.
function benefitBalanceEffect(t) {
  if (!t || t.type !== 'despesa' || !t.benefitId || !t.benefitType) return 0;
  return t.amount;
}
function applyBenefitDelta(benefitsList, benefitId, field, delta) {
  if (!benefitId || !field || !delta) return benefitsList;
  const historyField = field === 'foodBalance' ? 'foodHistory' : 'mobilityHistory';
  return benefitsList.map((b) => {
    if (b.id !== benefitId) return b;
    const newValue = Math.max(0, (b[field] || 0) - delta);
    return { ...b, [field]: newValue };
  });
}
function reapplyBenefitEffect(benefitsList, oldTx, newTx) {
  let updated = benefitsList;
  if (oldTx && benefitBalanceEffect(oldTx)) updated = applyBenefitDelta(updated, oldTx.benefitId, oldTx.benefitType, -benefitBalanceEffect(oldTx));
  if (newTx && benefitBalanceEffect(newTx)) updated = applyBenefitDelta(updated, newTx.benefitId, newTx.benefitType, benefitBalanceEffect(newTx));
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
    if (!isRealized(t)) return;
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
  transactions.filter((t) => t.type === 'despesa' && isRealized(t) && isSameMonth(t.date, year, month)).forEach((t) => {
    totals[t.category] = (totals[t.category] || 0) + t.amount;
  });
  return totals;
}

// Mesma ideia, mas seguindo o período selecionado no cabeçalho do Dashboard (mês/trimestre/ano/
// personalizado) em vez de ficar travado no mês-calendário atual — a mesma janela de tempo que
// os KPIs do topo (Receitas do período, Despesas do período) já usam. Sem isso, ao trocar pra
// "Trimestre" ou "Ano" o donut de categorias continuava só com o mês atual, então categorias com
// gasto só em meses anteriores do período sumiam do gráfico.
function computeCategoryTotalsForPeriod(transactions, period, customRange) {
  const totals = {};
  const add = (t) => { totals[t.category] = (totals[t.category] || 0) + t.amount; };
  if (period === 'personalizado' && customRange.start && customRange.end) {
    transactions.filter((t) => t.type === 'despesa' && isRealized(t) && t.date >= customRange.start && t.date <= customRange.end).forEach(add);
    return totals;
  }
  const monthsToTake = period === 'mes' ? 1 : period === 'trimestre' ? 3 : 12;
  const now = new Date();
  const startStr = ymd(new Date(now.getFullYear(), now.getMonth() - (monthsToTake - 1), 1));
  // Limite superior = fim do mês atual. Sem isso, as recorrências e parcelas já pré-geradas no
  // array de transações (até 12 meses à frente) entravam na soma, contando meses que ainda nem
  // aconteceram e inflando muito o total de categorias com gasto recorrente.
  const endStr = ymd(new Date(now.getFullYear(), now.getMonth() + 1, 0));
  transactions.filter((t) => t.type === 'despesa' && isRealized(t) && t.date >= startStr && t.date <= endStr).forEach(add);
  return totals;
}

// Fatura de um cartão: soma de tudo que foi lançado nele (despesas "Pago" com esse cardId)
// desde o último pagamento registrado. Não é mais um número solto — é sempre o espelho real
// dos lançamentos, então trocar o cartão de uma despesa atualiza a fatura na hora.
// A fatura "atual" (aberta) é tudo que foi lançado no cartão desde o último pagamento até o
// vencimento em aberto mais próximo — nunca inclui parcelas futuras de uma compra parcelada,
// que só entram na fatura quando chegar a vez delas (ou se forem antecipadas).
// Retorna a chave {year, month} (mês 0-indexed) do fechamento da fatura em que uma compra cai,
// respeitando o dia de fechamento do cartão — não o mês-calendário bruto da data de compra.
// Ex: fecha dia 22, compra em 25/jan → cai na fatura que fecha em fevereiro, não na de janeiro.
function getCardInvoiceCycle(card, dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  const closingMonth = d.getMonth() + (d.getDate() > card.closingDay ? 1 : 0);
  const normalized = new Date(d.getFullYear(), closingMonth, 1); // JS normaliza estouro de mês/ano
  return { year: normalized.getFullYear(), month: normalized.getMonth() };
}
function daysInMonth(year, month) { return new Date(year, month + 1, 0).getDate(); }
// Move a data de uma parcela pra fatura anterior/seguinte (direction: -1 ou 1), respeitando o
// fechamento do cartão — útil quando o banco atribuiu a compra a um mês diferente do que a
// pessoa esperava (ex: compra bem perto do fechamento).
function shiftToAdjacentInvoiceCycle(card, dateStr, direction) {
  const cycle = getCardInvoiceCycle(card, dateStr);
  if (direction > 0) {
    const boundary = new Date(cycle.year, cycle.month, Math.min(card.closingDay, daysInMonth(cycle.year, cycle.month)));
    return ymd(addDays(boundary, 1));
  }
  const prevRef = new Date(cycle.year, cycle.month - 1, 1);
  return ymd(new Date(prevRef.getFullYear(), prevRef.getMonth(), Math.min(card.closingDay, daysInMonth(prevRef.getFullYear(), prevRef.getMonth()))));
}

// Acha o FECHAMENTO (não o vencimento) do ciclo em aberto agora — a fronteira real que separa
// "já faz parte dessa fatura" de "só vai entrar na próxima", igual o getCardInvoiceCycle usa pra
// agrupar lançamentos na Fatura mensal. Diferente do vencimento (que pode cair em qualquer dia
// do mês seguinte ao fechamento), essa data é sempre um "dia de fechamento" de algum mês — por
// isso é o valor certo pra usar como paidThroughDate: assim o próximo lançamento já sabe, sem
// ambiguidade, se ficou "dentro" da fatura paga ou "fora" dela (na próxima).
function getOpenInvoiceClosingDate(card, referenceDate = new Date()) {
  const dueDate = getNextCardDueDate(card, referenceDate);
  const dueMonthOffset = card.dueDay < card.closingDay ? 1 : 0;
  const closingRef = new Date(dueDate.getFullYear(), dueDate.getMonth() - dueMonthOffset, 1);
  return new Date(closingRef.getFullYear(), closingRef.getMonth(), Math.min(card.closingDay, daysInMonth(closingRef.getFullYear(), closingRef.getMonth())));
}
// Corrige uma paidThroughDate que ficou desalinhada do ciclo (salva antes desta correção, como a
// data em que a pessoa clicou em pagar, não o fechamento de verdade), "encaixando" ela no
// fechamento de ciclo válido mais próximo — sempre pra trás, nunca pra frente, pra nunca marcar
// como já pago um lançamento que na verdade ainda está em aberto.
function snapToClosingBoundary(card, dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  const sameMonthClosing = new Date(d.getFullYear(), d.getMonth(), Math.min(card.closingDay, daysInMonth(d.getFullYear(), d.getMonth())));
  const closing = sameMonthClosing <= d ? sameMonthClosing : new Date(d.getFullYear(), d.getMonth() - 1, Math.min(card.closingDay, daysInMonth(d.getFullYear(), d.getMonth() - 1)));
  return ymd(closing);
}

function computeCardInvoice(card, transactions, referenceDate = new Date()) {
  const cutoff = ymd(getOpenInvoiceClosingDate(card, referenceDate));
  return transactions
    .filter((t) => t.cardId === card.id && t.type === 'despesa'
      && (!card.paidThroughDate || t.date > card.paidThroughDate)
      && t.date <= cutoff)
    .reduce((s, t) => s + t.amount, 0);
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
// Gera as ocorrências futuras (até `monthsAhead` meses a partir de hoje) de uma despesa
// recorrente como lançamentos de verdade, marcados com recurringId (permite editar/excluir em
// lote depois). Sempre nasce como "Pendente" — mesmo em débito — pra nunca descontar da conta
// silenciosamente um valor que ainda não aconteceu; a pessoa confirma normalmente quando o
// lançamento vence, do mesmo jeito que já funciona pra débito agendado. Nunca gera pra um mês
// que já tem um lançamento com o mesmo nome (evita duplicar uma cobrança lançada manualmente
// antes dessa função existir, ou lançada via "Lançar agora").
function generateRecurringOccurrences(item, cards, existingTransactions, monthsAhead = 12) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const card = item.cardId ? cards.find((c) => c.id === item.cardId) : null;
  const occurrences = [];
  for (let i = 0; i <= monthsAhead; i++) {
    const cursor = new Date(today.getFullYear(), today.getMonth() + i, 1);
    const day = Math.min(item.renewalDay, daysInMonth(cursor.getFullYear(), cursor.getMonth()));
    const occDate = new Date(cursor.getFullYear(), cursor.getMonth(), day);
    if (i === 0 && occDate < today) continue; // já passou esse mês — não gera retroativo
    const alreadyExists = existingTransactions.some((t) =>
      t.type === 'despesa' && t.description === item.name && isSameMonth(t.date, cursor.getFullYear(), cursor.getMonth())
    );
    if (alreadyExists) continue;
    occurrences.push({
      id: uid(), description: item.name, amount: item.value, category: item.category, type: 'despesa',
      account: card ? card.accountId : (item.accountId || null), cardId: card ? card.id : null,
      paymentMethod: card ? 'Cartão de crédito' : 'Não informado',
      date: ymd(occDate), status: 'Pendente', recurringId: item.id,
    });
  }
  return occurrences;
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

// Data de vencimento da fatura em que uma parcela cai, dado o dia da compra e o dia de
// fechamento do cartão: compras até o dia do fechamento entram na fatura que fecha nesse mês;
// compras depois do fechamento só entram na fatura do mês seguinte. Cada parcela seguinte
// (offsetMonths = 0, 1, 2...) empurra o fechamento um mês pra frente a partir daí.
function getInstallmentDueDate(card, purchaseDate, offsetMonths) {
  let closingMonth = purchaseDate.getMonth() + (purchaseDate.getDate() > card.closingDay ? 1 : 0) + offsetMonths;
  const closingRef = new Date(purchaseDate.getFullYear(), closingMonth, 1);
  // Se o dia do vencimento (numericamente) vem antes do dia do fechamento, o vencimento cai
  // no mês seguinte ao fechamento (ex: fecha dia 22, vence dia 5 → vence no mês seguinte).
  const dueMonthOffset = card.dueDay < card.closingDay ? 1 : 0;
  return getAdjustedDueDate(closingRef.getFullYear(), closingRef.getMonth() + dueMonthOffset, card.dueDay);
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

function daysBetween(dateA, dateB) {
  return Math.round((new Date(dateA + 'T00:00:00') - new Date(dateB + 'T00:00:00')) / 86400000);
}

// Semelhança simples entre dois nomes (0 a 1) — quantas palavras em comum, proporcionalmente.
// Não precisa ser um algoritmo sofisticado, só o suficiente pra reconhecer nomes parecidos mas
// não idênticos (ex.: "Steam", digitado à mão, vs "Pag*Steam", do jeito que o Nubank registra).
function textSimilarity(a, b) {
  const wordsA = new Set(normalizeForFingerprint(a).split(' ').filter(Boolean));
  const wordsB = new Set(normalizeForFingerprint(b).split(' ').filter(Boolean));
  if (wordsA.size === 0 || wordsB.size === 0) return 0;
  let common = 0;
  wordsA.forEach((w) => { if (wordsB.has(w)) common++; });
  return common / Math.max(wordsA.size, wordsB.size);
}

// Acha, entre os lançamentos já existentes ainda não "usados" por outra linha da fatura, o
// melhor candidato a duplicata pra uma linha importada — nessa ordem de prioridade, como pedido:
// 1) valor: funciona como filtro (tem que bater, é o sinal mais confiável pra dinheiro);
// 2) data: quanto mais perto, maior a pontuação (aceita até 5 dias de diferença, cobrindo
//    compra x postagem no extrato, ou fuso/final de semana);
// 3) nome: só entra como desempate leve, via semelhança de palavras — não precisa ser idêntico.
function findBestDuplicateMatch(row, pool) {
  const rowAmount = Math.abs(row.amount).toFixed(2);
  let best = null;
  let bestScore = -Infinity;
  pool.forEach((c) => {
    if (c.used || Math.abs(c.amount).toFixed(2) !== rowAmount) return;
    const dayDiff = Math.abs(daysBetween(row.date, c.date));
    if (dayDiff > 5) return;
    const dateScore = (5 - dayDiff) / 5;
    const nameScore = textSimilarity(row.description, c.description);
    const score = dateScore * 10 + nameScore;
    if (score > bestScore) { bestScore = score; best = c; }
  });
  return best;
}

// Faz o parse do CSV de fatura do Nubank (colunas: date,title,amount) separando compras reais
// de pagamentos/adiantamentos de fatura, e sinalizando duplicatas contra o que já existe no app.
function parseNubankFaturaCSV(rows, existingTransactions) {
  const pool = existingTransactions.map((t) => ({ date: t.date, description: t.description, amount: t.amount, used: false }));
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
    const match = findBestDuplicateMatch({ date: rawDate, description, amount }, pool);
    const isDuplicate = !!match;
    // "Exata" quando data e nome normalizado batem plenamente — usada só pra decidir o texto
    // mostrado (duplicata certa vs. possível duplicata), a lógica de seleção é a mesma pras duas.
    const isExactMatch = isDuplicate && match.date === rawDate && normalizeForFingerprint(match.description) === normalizeForFingerprint(description);
    if (match) match.used = true;

    purchases.push({
      rowId: `row-${idx}`,
      date: rawDate,
      description,
      category: guessCategory(description),
      amount: Math.abs(amount),
      installment: installment ? { current: installment.current, total: installment.total } : null,
      isDuplicate,
      isExactMatch,
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
  { id: 'card-1', bank: 'Nubank', brand: 'Mastercard Black', accountId: 'acc-1', limit: 8000, closingDay: 22, dueDay: 29, paidThroughDate: null },
  { id: 'card-2', bank: 'Itaú', brand: 'Visa Click', accountId: 'acc-2', limit: 5000, closingDay: 5, dueDay: 12, paidThroughDate: null },
  { id: 'card-3', bank: 'Inter', brand: 'Gold Mastercard', accountId: 'acc-3', limit: 4000, closingDay: 15, dueDay: 22, paidThroughDate: null },
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
const GOAL_ICONS = { Wallet, Plane, Laptop, Award, Home, Car, Bike, Shirt, GraduationCap, Heart, Briefcase, Sprout, PersonStanding, PiggyBank };

// Temas de cor de destaque, cada um com uma variante para claro e uma para escuro.
// As cores semânticas (receita, despesa, investimento, meta, alerta) não mudam entre temas.
const COLOR_THEMES = {
  default: { label: 'Verde militar', light: '#5D7052', dark: '#8FA680' },
  azul: { label: 'Azul', light: '#4A6FA5', dark: '#7CA3D6' },
  rosa: { label: 'Rosa', light: '#B5677E', dark: '#D98CA0' },
  vermelho: { label: 'Terracota', light: '#A85D4E', dark: '#CC8271' },
  branco: { label: 'Monocromático', light: '#3A3A3A', dark: '#C9C9C9' },
  ameixa: { label: 'Ameixa', light: '#7D5A82', dark: '#B08AB8' },
};

/* ---------- Derivar tonalidades de uma cor-base (usado nos gradientes dos cartões) ---------- */

function hexToRgb(hex) {
  const h = hex.replace('#', '');
  const full = h.length === 3 ? h.split('').map((c) => c + c).join('') : h;
  const n = parseInt(full, 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}
function rgbToHex({ r, g, b }) {
  return '#' + [r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('');
}
// Mistura hexA com hexB; t=0 devolve hexA puro, t=1 devolve hexB puro.
function mixHex(hexA, hexB, t) {
  const a = hexToRgb(hexA), b = hexToRgb(hexB);
  return rgbToHex({ r: a.r + (b.r - a.r) * t, g: a.g + (b.g - a.g) * t, b: a.b + (b.b - a.b) * t });
}

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
  return rows.map((r) => {
    const accountToCard = { 'acc-1': 'card-1', 'acc-2': 'card-2', 'acc-3': 'card-3' };
    return {
      id: uid(), date: r[0], description: r[1], category: r[2], type: r[3],
      account: r[4], paymentMethod: r[5], amount: r[6], status: r[7],
      cardId: r[5] === 'Cartão de crédito' ? accountToCard[r[4]] || null : null,
    };
  });
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

function generateInsights({ transactions, goals, monthlyHistory: mh, accounts = [], cards = [], categoryComparison }) {
  const insights = [];
  const thisMonth = mh[mh.length - 1];
  const lastMonth = mh.length > 1 ? mh[mh.length - 2] : null;

  accounts.filter((a) => a.alertThreshold > 0 && a.balance <= a.alertThreshold).forEach((a) => {
    insights.unshift({
      icon: AlertCircle, tone: 'expense',
      text: `Saldo da conta ${a.bank} está em ${formatBRL(a.balance)}, abaixo do limite de alerta que você definiu (${formatBRL(a.alertThreshold)}).`,
    });
  });

  cards.forEach((c) => {
    const due = getNextCardDueDate(c);
    const daysLeft = Math.round((due - new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate())) / 86400000);
    if (daysLeft >= 0 && daysLeft <= 5) {
      const invoice = computeCardInvoice(c, transactions);
      insights.unshift({
        icon: CalendarIcon, tone: 'alert',
        text: daysLeft === 0
          ? `A fatura do cartão ${c.bank} vence hoje (${formatBRL(invoice)}).`
          : `A fatura do cartão ${c.bank} vence em ${daysLeft} dia(s) (${formatBRL(invoice)}).`,
      });
    }
  });

  goals.filter((g) => g.current < g.target && g.current / g.target >= 0.9).forEach((g) => {
    insights.push({
      icon: Target, tone: 'goals',
      text: `Você atingiu ${((g.current / g.target) * 100).toFixed(0)}% da meta "${g.name}".`,
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

function CategoryBadge({ category, showIcon = true }) {
  const cat = CATEGORIES[category] || CATEGORIES['Outros'];
  return <Badge color={cat.color} soft={cat.soft} icon={showIcon ? cat.icon : undefined}>{category}</Badge>;
}

function StatusBadge({ status, type }) {
  const map = {
    'Pago': { color: 'var(--income)', soft: 'var(--income-soft)' },
    'Pendente': { color: 'var(--alert)', soft: 'var(--alert-soft)' },
  };
  const s = map[status] || map['Pendente'];
  return <Badge color={s.color} soft={s.soft}>{statusLabel(status, type)}</Badge>;
}

function IconCircle({ icon: Icon, color, soft, size = 40 }) {
  return (
    <div className="rounded-xl flex items-center justify-center shrink-0" style={{ width: size, height: size, backgroundColor: soft }}>
      <Icon size={size * 0.45} color={color} strokeWidth={2.2} />
    </div>
  );
}

// Card com gesto de arrastar pra esquerda (swipe) que revela ações rápidas (editar/excluir) por
// trás — usado nas listas mobile no lugar das tabelas largas, que não cabem numa tela de celular.
// touchAction: 'pan-y' deixa o scroll vertical da página funcionando normalmente enquanto captura
// o arrasto horizontal.
// deleteConfirm customiza o texto do ConfirmModal ({ title, description }) exibido antes de
// executar onDelete de verdade — arrastar e tocar em excluir nunca apaga na hora.
function SwipeableRow({ children, onEdit, onDelete, deleteConfirm }) {
  const ACTIONS_WIDTH = 96;
  const [dragX, setDragX] = useState(0);
  const [dragging, setDragging] = useState(false);
  const [confirmingDelete, setConfirmingDelete] = useState(false);
  const startXRef = useRef(0);
  const startDragXRef = useRef(0);

  function handleTouchStart(e) {
    startXRef.current = e.touches[0].clientX;
    startDragXRef.current = dragX;
    setDragging(true);
  }
  function handleTouchMove(e) {
    const delta = e.touches[0].clientX - startXRef.current;
    setDragX(Math.max(-ACTIONS_WIDTH, Math.min(0, startDragXRef.current + delta)));
  }
  function handleTouchEnd() {
    setDragging(false);
    setDragX((cur) => (cur < -ACTIONS_WIDTH / 2 ? -ACTIONS_WIDTH : 0));
  }

  return (
    <div className="relative overflow-hidden rounded-xl">
      <div className="absolute inset-y-0 right-0 flex items-stretch" style={{ width: ACTIONS_WIDTH }}>
        {onEdit && (
          <button onClick={() => { onEdit(); setDragX(0); }} className="flex-1 flex items-center justify-center" style={{ backgroundColor: 'var(--primary)' }} title="Editar">
            <Pencil size={18} color="#fff" />
          </button>
        )}
        {onDelete && (
          <button onClick={() => setConfirmingDelete(true)} className="flex-1 flex items-center justify-center" style={{ backgroundColor: 'var(--expense)' }} title="Excluir">
            <Trash2 size={18} color="#fff" />
          </button>
        )}
      </div>
      <div
        onTouchStart={handleTouchStart} onTouchMove={handleTouchMove} onTouchEnd={handleTouchEnd}
        style={{ transform: `translateX(${dragX}px)`, transition: dragging ? 'none' : 'transform 0.2s ease', backgroundColor: 'var(--card)', touchAction: 'pan-y', position: 'relative' }}
      >
        {children}
      </div>
      {confirmingDelete && (
        <ConfirmModal
          title={deleteConfirm?.title || 'Excluir lançamento'}
          description={deleteConfirm?.description || 'Tem certeza que deseja excluir este item? Essa ação não pode ser desfeita.'}
          onConfirm={() => { onDelete(); setConfirmingDelete(false); setDragX(0); }}
          onClose={() => { setConfirmingDelete(false); setDragX(0); }}
        />
      )}
    </div>
  );
}

function Button({ children, variant = 'primary', size = 'md', icon: Icon, onClick, type = 'button', className = '', disabled }) {
  const base = 'inline-flex items-center justify-center gap-2 rounded-xl font-medium transition-all focus-ring disabled:opacity-50 disabled:cursor-not-allowed';
  const sizes = {
    sm: 'px-3 py-1.5 text-xs',
    // Compacto no mobile (cabe "Importar fatura" + "Novo" + "⋮" numa linha só sem
    // estourar pra uma 3ª linha) e no tamanho normal a partir do breakpoint sm.
    toolbar: 'px-2.5 py-1.5 text-xs sm:px-3 sm:py-2 sm:text-sm',
    md: 'px-4 py-2.5 text-sm', lg: 'px-5 py-3 text-sm',
  };
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
      {Icon && <Icon size={size === 'sm' || size === 'toolbar' ? 14 : 16} />}
      {children}
    </button>
  );
}

// Substitui o <select> nativo: no mobile, um <select> normal abre a caixa de opções do
// próprio sistema operacional (aquela lista cinza padrão, "poluída" e fora do visual do
// app). O Select abaixo tem a MESMA API de um <select> — value, onChange, disabled, filhos
// <option> — só que desenha o próprio menu suspenso, então em qualquer lugar do código dá
// pra trocar a tag <select> por <Select> sem mudar mais nada.
function flattenSelectChildren(children) {
  const items = [];
  React.Children.forEach(children, (child) => {
    if (!child) return;
    if (Array.isArray(child)) { items.push(...flattenSelectChildren(child)); return; }
    if (child.type === 'option') { items.push({ kind: 'option', el: child }); return; }
    if (child.type === 'optgroup') {
      items.push({ kind: 'group', label: child.props.label });
      items.push(...flattenSelectChildren(child.props.children));
    }
  });
  return items;
}

function optionValue(opt) {
  // Um <option> sem o atributo value usa o próprio texto como valor — mesmo comportamento
  // do <select> nativo (usado, por exemplo, em "Conta Corrente" / "Poupança").
  return opt.props.value !== undefined ? opt.props.value : opt.props.children;
}

// Popover genérico usado pelo <Select>, pelo menu "Filtros", pelo menu de exportar, pelo
// seletor de período e pelas notificações: renderiza num portal (document.body) e calcula a
// posição a partir da tela toda (não do layout do pai), então nunca fica cortado por um card/
// modal com "overflow" no caminho. A posição — incluindo altura máxima — é sempre recalculada
// pra caber inteiro na tela (abrindo pra cima se precisar, encostando na borda se precisar),
// então nunca sobra conteúdo inacessível fora da viewport. Fecha ao tocar numa camada
// invisível atrás do painel — nunca ao tocar dentro dele (mesmo arrastando pra rolar).
function Popover({ open, onClose, triggerRef, children, width = 'trigger', align = 'left', className = '', panelStyle = {} }) {
  const [rect, setRect] = useState(null);
  const panelRef = useRef(null);

  useEffect(() => {
    if (!open) { setRect(null); return; }
    function update() {
      if (!triggerRef.current) return;
      const r = triggerRef.current.getBoundingClientRect();
      const margin = 8;
      const vw = window.innerWidth, vh = window.innerHeight;
      const panelWidth = width === 'trigger' ? r.width : Math.min(width, vw - margin * 2);
      let left;
      if (width === 'trigger') left = r.left;
      else if (align === 'center') left = r.left + r.width / 2 - panelWidth / 2;
      else if (align === 'right') left = r.right - panelWidth;
      else left = r.left;
      if (left + panelWidth > vw - margin) left = vw - margin - panelWidth;
      if (left < margin) left = margin;
      const spaceBelow = vh - r.bottom - margin;
      const spaceAbove = r.top - margin;
      const heightGuess = panelRef.current?.offsetHeight ?? 320;
      const openBelow = spaceBelow >= Math.min(heightGuess, 160) || spaceBelow >= spaceAbove;
      if (openBelow) setRect({ left, width: panelWidth, top: r.bottom + 6, maxHeight: Math.max(120, spaceBelow - 6) });
      else setRect({ left, width: panelWidth, bottom: vh - r.top + 6, maxHeight: Math.max(120, spaceAbove - 6) });
    }
    update();
    window.addEventListener('scroll', update, true);
    window.addEventListener('resize', update);
    return () => {
      window.removeEventListener('scroll', update, true);
      window.removeEventListener('resize', update);
    };
  }, [open, triggerRef, width, align]);

  const portalTarget = typeof document !== 'undefined' ? document.body : null;
  if (!open || !rect || !portalTarget) return null;

  return createPortal(
    <>
      <div data-select-portal className="fixed inset-0 z-[99]" onClick={onClose} />
      <div
        ref={panelRef} data-select-portal
        className={`fixed overflow-y-auto overscroll-contain rounded-xl shadow-soft-lg z-[100] animate-fade-up ${className}`}
        style={{ left: rect.left, width: rect.width, top: rect.top, bottom: rect.bottom, maxHeight: rect.maxHeight, backgroundColor: 'var(--card)', border: '1px solid var(--border)', WebkitOverflowScrolling: 'touch', ...panelStyle }}
      >
        {children}
      </div>
    </>,
    portalTarget
  );
}

function Select({ children, value, onChange, className = '', style = {}, disabled = false }) {
  const [open, setOpen] = useState(false);
  const triggerRef = useRef(null);

  const items = flattenSelectChildren(children);
  const selectedItem = items.find((it) => it.kind === 'option' && String(optionValue(it.el)) === String(value));

  function pick(opt) {
    if (opt.props.disabled) return;
    onChange({ target: { value: optionValue(opt) } });
    setOpen(false);
  }

  return (
    <>
      <button
        ref={triggerRef}
        type="button" disabled={disabled} onClick={() => setOpen((v) => !v)}
        className={`${className} flex items-center justify-between gap-2 text-left disabled:cursor-not-allowed`}
        style={style}
      >
        <span className="truncate">{selectedItem ? selectedItem.el.props.children : ''}</span>
        <ChevronDown size={14} className={`shrink-0 transition-transform ${open ? 'rotate-180' : ''}`} style={{ color: 'var(--text-soft)' }} />
      </button>
      <Popover open={open} onClose={() => setOpen(false)} triggerRef={triggerRef} width="trigger" className="rounded-2xl py-1">
        {items.map((it, i) => {
          if (it.kind === 'group') {
            return (
              <p key={`g-${i}`} className={`px-4 text-[11px] font-semibold uppercase tracking-wide ${i > 0 ? 'mt-2 pt-2' : ''} pb-1`} style={{ color: 'var(--text-soft)', borderTop: i > 0 ? '1px solid var(--border)' : 'none' }}>
                {it.label}
              </p>
            );
          }
          const opt = it.el;
          const isSelected = String(optionValue(opt)) === String(value);
          return (
            <button
              key={opt.key ?? i} type="button" disabled={opt.props.disabled} onClick={() => pick(opt)}
              className="w-full text-left px-4 py-3.5 text-base flex items-center justify-between gap-3 hover:bg-[var(--primary-soft)] disabled:opacity-50 disabled:cursor-not-allowed"
              style={{ color: 'var(--text)', borderTop: i > 0 && items[i - 1]?.kind !== 'group' ? '1px solid var(--border)' : 'none' }}
            >
              <span className="truncate">{opt.props.children}</span>
              <span className="w-5 h-5 rounded-full border-2 flex items-center justify-center shrink-0" style={{ borderColor: isSelected ? 'var(--primary)' : 'var(--border)' }}>
                {isSelected && <span className="w-2.5 h-2.5 rounded-full" style={{ backgroundColor: 'var(--primary)' }} />}
              </span>
            </button>
          );
        })}
      </Popover>
    </>
  );
}

const Card = React.forwardRef(function Card({ children, className = '', padding = 'p-6' }, ref) {
  return (
    <div ref={ref} className={`rounded-2xl shadow-soft ${padding} ${className}`} style={{ backgroundColor: 'var(--card)', border: '1px solid var(--border)' }}>
      {children}
    </div>
  );
});

function SectionTitle({ children, subtitle, action }) {
  return (
    <div className="flex items-start justify-between gap-3 mb-4 flex-wrap">
      <div className="min-w-0">
        <h2 className="font-display text-lg font-semibold" style={{ color: 'var(--text)' }}>{children}</h2>
        {subtitle && <p className="text-xs font-normal mt-0.5" style={{ color: 'var(--text-soft)' }}>{subtitle}</p>}
      </div>
      {action && <div className="shrink-0">{action}</div>}
    </div>
  );
}

function EmptyState({ icon: Icon = Search, title, description }) {
  return (
    <div className="flex flex-col items-center justify-center text-center py-8 sm:py-14 px-4">
      <div className="rounded-2xl p-3 sm:p-4 mb-3 sm:mb-4" style={{ backgroundColor: 'var(--primary-soft)' }}>
        <Icon size={24} color="var(--primary)" />
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
          <button onClick={onClose} className="p-2 rounded-lg hover:bg-black/5 focus-ring" aria-label="Fechar" title="Fechar">
            <X size={18} color="var(--text-soft)" />
          </button>
        </div>
        <div className="p-6">{children}</div>
      </div>
    </div>
  );
}

function ConfirmModal({ title, description, onConfirm, onClose, variant = 'danger', confirmLabel = 'Confirmar' }) {
  return (
    <Modal title={title} onClose={onClose}>
      <p className="text-sm mb-6" style={{ color: 'var(--text-soft)' }}>{description}</p>
      <div className="flex flex-wrap justify-end gap-3">
        <Button variant="secondary" onClick={onClose}>Cancelar</Button>
        <Button variant={variant} onClick={() => { onConfirm(); onClose(); }}>{confirmLabel}</Button>
      </div>
    </Modal>
  );
}

function SyncConflictModal({ date, onUseRemote, onKeepLocal }) {
  return (
    <Modal title="Encontramos uma versão mais recente" onClose={onKeepLocal}>
      <div className="space-y-4">
        <p className="text-sm flex items-start gap-2" style={{ color: 'var(--text)' }}>
          <Cloud size={16} className="mt-0.5 shrink-0" color="var(--primary)" />
          <span>Parece que você fez alterações em outro aparelho. O backup mais recente no Dropbox é de <strong>{relativeTime(date)}</strong> — mais novo que os dados deste navegador.</span>
        </p>
        <div className="flex flex-col gap-2">
          <Button onClick={onUseRemote}>Usar a versão do Dropbox</Button>
          <Button variant="secondary" onClick={onKeepLocal}>Manter os dados deste aparelho</Button>
        </div>
        <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Usar a versão do Dropbox substitui os dados deste navegador. Manter os dados daqui reenvia essa versão para o Dropbox, sobrescrevendo o backup mais novo.</p>
      </div>
    </Modal>
  );
}

/* ---------- Toasts ---------- */

// Botão flutuante de novo lançamento — só faz sentido em telas onde a barra lateral não fica
// sempre visível (abaixo do breakpoint lg, ela vira uma gaveta que precisa do menu hambúrguer
// pra abrir). Fica rente à base da tela, no mesmo nível da faixa dos toasts.
// Enquanto o usuário rola a tela ele encolhe pela metade, e volta ao tamanho normal depois
// de 0,75s sem nenhum scroll (ver handleContentScroll no App).
function FAB({ onClick, shrink }) {
  return (
    <button
      onClick={onClick}
      className={`fixed z-20 w-14 h-14 rounded-2xl flex items-center justify-center shadow-soft-lg no-print lg:hidden active:scale-95 ${shrink ? 'fab-spin-in' : 'fab-spin-out'}`}
      style={{ right: '1.25rem', bottom: 'calc(1.5rem + env(safe-area-inset-bottom, 0px))', backgroundColor: 'var(--primary)' }}
      title="Novo lançamento"
    >
      <Plus size={26} color="#fff" />
    </button>
  );
}

function ToastContainer({ toasts }) {
  return (
    <div className="fixed left-4 right-4 sm:left-auto sm:right-6 z-[60] flex flex-col items-stretch sm:items-end gap-2 no-print" style={{ bottom: 'calc(1.5rem + env(safe-area-inset-bottom, 0px))' }}>
      {toasts.map((t) => (
        <div key={t.id} className="animate-toast flex items-center gap-2.5 px-4 py-3 rounded-xl shadow-soft-lg text-sm font-medium sm:max-w-sm" style={{ backgroundColor: t.type === 'error' ? 'var(--expense)' : 'var(--primary-dark)', color: '#fff' }}>
          {t.type === 'error' ? <AlertCircle size={16} className="shrink-0" /> : <Check size={16} className="shrink-0" />}
          <span className="min-w-0">{t.message}</span>
        </div>
      ))}
    </div>
  );
}

/* ---------- Comemoração sutil ao completar uma meta (sem popup — só os emojis) ---------- */

const CELEBRATION_EMOJIS = ['🥳', '🎉', '🎊'];
const CELEBRATION_PARTICLE_COUNT = 20;
const CELEBRATION_DURATION_MS = 3800; // cobre com folga o pior caso de dur+delay das partículas (até ~3.7s)

function GoalCelebration({ origin, onDone }) {
  // Gera as partículas uma única vez (na montagem), não a cada render — senão a animação reinicia sozinha.
  const particles = useMemo(() => {
    return Array.from({ length: CELEBRATION_PARTICLE_COUNT }, (_, i) => {
      const angle = Math.random() * Math.PI * 2;
      const distance = 70 + Math.random() * 170;
      return {
        id: i,
        emoji: CELEBRATION_EMOJIS[Math.floor(Math.random() * CELEBRATION_EMOJIS.length)],
        tx: Math.cos(angle) * distance,
        ty: Math.sin(angle) * distance,
        size: 18 + Math.random() * 16,
        dur: 2.4 + Math.random() * 1.1,
        delay: Math.random() * 0.2,
      };
    });
  }, []);

  useEffect(() => {
    const timer = setTimeout(onDone, CELEBRATION_DURATION_MS);
    return () => clearTimeout(timer);
  }, [onDone]);

  // Sem posição de origem capturada (ex: card fora da tela) — cai num ponto central razoável,
  // pra comemoração nunca quebrar, só ficar um pouco menos "conectada" ao card.
  const x = origin?.x ?? (typeof window !== 'undefined' ? window.innerWidth / 2 : 200);
  const y = origin?.y ?? (typeof window !== 'undefined' ? window.innerHeight / 2.5 : 200);

  return (
    <div className="fixed inset-0 pointer-events-none overflow-hidden" style={{ zIndex: 60 }} aria-hidden="true">
      {particles.map((p) => (
        <span
          key={p.id}
          className="goal-celebrate-particle"
          style={{ '--origin-x': `${x}px`, '--origin-y': `${y}px`, '--tx': `${p.tx}px`, '--ty': `${p.ty}px`, '--size': `${p.size}px`, '--dur': `${p.dur}s`, '--delay': `${p.delay}s` }}
        >
          {p.emoji}
        </span>
      ))}
    </div>
  );
}

/* ---------- Formulário: campo de moeda ---------- */

function CurrencyInput({ value, onChange, placeholder = '0,00' }) {
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
        className="w-full pl-9 pr-3.5 py-2.5 rounded-xl text-base sm:text-sm focus-ring tabular-nums"
        style={{ border: '1px solid var(--border)', color: 'var(--text)', backgroundColor: 'var(--card)' }}
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

// Pra saldo de conta (diferente de valor de lançamento, que é sempre positivo — o sinal vem do
// tipo receita/despesa): aqui o próprio valor pode ser negativo (conta no cheque especial, cartão
// de débito estourado etc). Um teclado numérico de celular não tem tecla de "-", então em vez de
// pedir pra digitar o sinal, um botão alterna entre positivo/negativo e o campo edita só a
// magnitude — funciona igual em qualquer teclado.
function SignedCurrencyInput({ value, onChange, placeholder }) {
  const isNegative = value < 0;
  return (
    <div className="flex items-center gap-2">
      <button
        type="button"
        onClick={() => onChange(-value)}
        title={isNegative ? 'Tornar positivo' : 'Tornar negativo (conta no vermelho)'}
        className="shrink-0 w-10 h-10 rounded-xl flex items-center justify-center font-display font-bold text-lg focus-ring transition-colors"
        style={{
          backgroundColor: isNegative ? 'var(--expense-soft)' : 'var(--bg)',
          color: isNegative ? 'var(--expense)' : 'var(--text-soft)',
          border: '1px solid var(--border)',
        }}
      >
        {isNegative ? '−' : '+'}
      </button>
      <div className="flex-1">
        <CurrencyInput value={Math.abs(value)} onChange={(v) => onChange(isNegative ? -v : v)} placeholder={placeholder} />
      </div>
    </div>
  );
}

const inputClass = 'w-full px-3.5 py-2.5 rounded-xl text-base sm:text-sm focus-ring';
const inputStyle = { border: '1px solid var(--border)', color: 'var(--text)', backgroundColor: 'var(--card)' };

/* ============================================================
   SIDEBAR / HEADER / BANNER
   ============================================================ */

const NAV_ITEMS = [
  { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
  { id: 'contas', label: 'Contas Bancárias', icon: Landmark },
  { id: 'cartoes', label: 'Cartões', icon: CreditCard },
  { id: 'transacoes', label: 'Transações', icon: ArrowLeftRight },
  { id: 'investimentos', label: 'Investimentos', icon: PieChartIcon },
  { id: 'metas', label: 'Metas', icon: Target },
  { id: 'recorrentes', label: 'Despesas Recorrentes', icon: RefreshCw },
  { id: 'relatorios', label: 'Relatórios', icon: BarChart3 },
  { id: 'configuracoes', label: 'Configurações', icon: Settings },
];

function Sidebar({ activePage, setActivePage, sidebarOpen, setSidebarOpen, onNewTransaction, dropboxConnected, dropboxLastBackup, dropboxSyncError, onGoToSettings }) {
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
          <button onClick={() => setSidebarOpen(false)} className="ml-auto lg:hidden p-1.5 rounded-lg hover:bg-black/5" title="Fechar menu">
            <X size={18} />
          </button>
        </div>

        <button onClick={onGoToSettings} className="mx-6 mb-4 p-3 rounded-xl flex items-center gap-2.5 text-left hover:bg-black/[0.03] transition-colors" style={{ backgroundColor: 'var(--primary-soft)' }}>
          {dropboxConnected ? (
            <Cloud size={16} color={dropboxSyncError ? 'var(--expense)' : 'var(--primary)'} className="shrink-0" />
          ) : (
            <EyeOff size={16} color="var(--text-soft)" className="shrink-0" />
          )}
          <div className="min-w-0">
            <p className="text-xs font-medium truncate" style={{ color: dropboxSyncError ? 'var(--expense)' : 'var(--text)' }}>
              {!dropboxConnected ? 'Backup na nuvem desativado' : dropboxSyncError ? 'Falha no último backup' : 'Backup automático ativo'}
            </p>
            <p className="text-[11px]" style={{ color: 'var(--text-soft)' }}>
              {dropboxConnected && dropboxLastBackup ? `Backup: ${relativeTime(dropboxLastBackup)}` : dropboxConnected ? 'Aguardando primeiro backup' : 'Toque para configurar'}
            </p>
          </div>
        </button>

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

// Fecha um dropdown/painel ao clicar fora dele — usado no filtro de período e no sino de
// notificações, que hoje só fecham se o usuário clicar de novo no botão que abriu.
function useClickOutside(ref, onOutside, active) {
  useEffect(() => {
    if (!active) return;
    function handle(e) {
      if (!ref.current || ref.current.contains(e.target)) return;
      // O menu do <Select> é renderizado num portal (fora dessa árvore de elementos), então
      // tocar numa opção dele conta como "fora" pra esse ref — sem essa checagem, selecionar
      // algo dentro de um <Select> aninhado (ex: no popover de Filtros) fechava o popover
      // inteiro antes do clique ser processado.
      if (e.target.closest?.('[data-select-portal]')) return;
      onOutside();
    }
    document.addEventListener('mousedown', handle);
    document.addEventListener('touchstart', handle);
    return () => {
      document.removeEventListener('mousedown', handle);
      document.removeEventListener('touchstart', handle);
    };
  }, [ref, onOutside, active]);
}

function PeriodSelector({ period, setPeriod, customRange, setCustomRange }) {
  const [open, setOpen] = useState(false);
  const triggerRef = useRef(null);
  const labels = { mes: 'Este mês', trimestre: 'Trimestre', ano: 'Ano', personalizado: 'Personalizado' };
  return (
    <>
      <button
        ref={triggerRef}
        onClick={() => setOpen(!open)} title={labels[period]}
        className="flex items-center gap-2 h-11 box-border px-3 sm:px-3.5 rounded-xl text-sm font-medium focus-ring"
        style={{ border: '1px solid var(--border)', color: 'var(--text)' }}
      >
        <CalendarIcon size={15} color="var(--text-soft)" />
        <span className="hidden sm:inline">{labels[period]}</span>
        <ChevronDown size={14} color="var(--text-soft)" className="hidden sm:inline" />
      </button>
      <Popover open={open} onClose={() => setOpen(false)} triggerRef={triggerRef} width={256} align="center" className="p-2">
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
      </Popover>
    </>
  );
}

function Header({ period, setPeriod, customRange, setCustomRange, search, setSearch, setSidebarOpen, insights }) {
  const [notifOpen, setNotifOpen] = useState(false);
  const notifRef = useRef(null);
  const [searchExpanded, setSearchExpanded] = useState(false);
  const [dismissed, setDismissed] = useState(() => {
    try { return new Set(JSON.parse(window.localStorage.getItem('cerne-dismissed-notifications-v1')) || []); } catch { return new Set(); }
  });
  const visibleInsights = insights.filter((n) => !dismissed.has(n.text));
  function dismissNotification(text) {
    const updated = new Set(dismissed); updated.add(text);
    setDismissed(updated);
    try { window.localStorage.setItem('cerne-dismissed-notifications-v1', JSON.stringify([...updated])); } catch { /* melhor esforço */ }
  }
  const today = capitalizeFirst(new Date().toLocaleDateString('pt-BR', { weekday: 'long', day: 'numeric', month: 'long' }));
  return (
    <header className="sticky top-0 z-30 flex items-center gap-2 sm:gap-3 px-3 sm:px-6 lg:px-8 py-3 sm:py-4 no-print" style={{ backgroundColor: 'var(--bg)', borderBottom: '1px solid var(--border)' }}>
      {!searchExpanded && (
        <button onClick={() => setSidebarOpen(true)} className="lg:hidden h-11 w-11 box-border flex items-center justify-center rounded-xl hover:bg-black/5 shrink-0" title="Abrir menu">
          <Menu size={20} color="var(--text)" />
        </button>
      )}
      {!searchExpanded && (
        <div className="hidden md:block min-w-0">
          <p className="font-display text-base font-semibold truncate" style={{ color: 'var(--text)' }}>{getGreeting()}!</p>
          <p className="text-xs" style={{ color: 'var(--text-soft)' }}>{today}</p>
        </div>
      )}

      {/* No mobile, a busca some por padrão (não cabia sem cortar o texto do placeholder) e vira
          só um ícone — ao tocar, expande e toma o espaço dos outros ícones do cabeçalho
          temporariamente. Em telas ≥ sm ela sempre aparece inteira, como antes. */}
      <div className="hidden sm:flex flex-1 min-w-0 justify-center px-1 sm:px-2">
        <div className="relative w-full max-w-sm min-w-0">
          <Search size={15} className="absolute left-3.5 top-1/2 -translate-y-1/2" color="var(--text-soft)" />
          <input
            value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Buscar lançamentos"
            className="w-full pl-9 pr-9 py-2 rounded-xl text-base sm:text-sm focus-ring" style={{ border: '1px solid var(--border)', backgroundColor: 'var(--card)', color: 'var(--text)' }}
          />
          {search && (
            <button
              onClick={() => setSearch('')} title="Limpar busca"
              className="absolute right-1.5 top-1/2 -translate-y-1/2 p-2 rounded-full hover:bg-black/10"
            >
              <Trash2 size={14} color="var(--text-soft)" />
            </button>
          )}
        </div>
      </div>
      {/* Este wrapper fica sempre flex-1 (recolhido ou expandido) pra empurrar o filtro
          mensal e o sino de notificações pro canto direito do cabeçalho — antes, no estado
          recolhido, ele encolhia junto com o ícone e os outros ícones ficavam todos
          amontoados à esquerda. */}
      <div className="sm:hidden flex-1 min-w-0">
        {searchExpanded ? (
          <div className="relative w-full">
            <Search size={15} className="absolute left-3.5 top-1/2 -translate-y-1/2" color="var(--text-soft)" />
            <input
              autoFocus
              value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Buscar lançamentos"
              className="w-full h-11 box-border pl-9 pr-9 text-base focus-ring" style={{ border: '1px solid var(--border)', backgroundColor: 'var(--card)', color: 'var(--text)' }}
            />
            <button
              onClick={() => { setSearch(''); setSearchExpanded(false); }} title="Fechar busca"
              className="absolute right-1.5 top-1/2 -translate-y-1/2 p-2 rounded-full hover:bg-black/10"
            >
              <X size={14} color="var(--text-soft)" />
            </button>
          </div>
        ) : (
          // Estado de repouso: agora mostra a lupa e o texto "Buscar lançamentos" por
          // inteiro, como uma caixa de busca de verdade (não só o ícone) — o toque abre o
          // input de fato editável logo abaixo. Fonte um pouco menor (text-xs) pra caber
          // sem cortar o texto nessa largura reduzida, dividida com os outros ícones do
          // cabeçalho.
          <button
            onClick={() => setSearchExpanded(true)}
            className="flex items-center gap-2 w-full h-11 box-border pl-3.5 pr-3 rounded-xl focus-ring"
            style={{ border: '1px solid var(--border)', backgroundColor: 'var(--card)' }}
            title="Buscar"
          >
            <Search size={15} className="shrink-0" color="var(--text-soft)" />
            <span className="text-xs truncate" style={{ color: 'var(--text-soft)' }}>Buscar lançamentos</span>
          </button>
        )}
      </div>

      <div className={`shrink-0${searchExpanded ? ' hidden sm:block' : ''}`}>
        <PeriodSelector period={period} setPeriod={setPeriod} customRange={customRange} setCustomRange={setCustomRange} />
      </div>
      <div className={`shrink-0${searchExpanded ? ' hidden sm:block' : ''}`}>
        <button ref={notifRef} onClick={() => setNotifOpen(!notifOpen)} className="relative h-11 w-11 box-border flex items-center justify-center rounded-xl hover:bg-black/5 focus-ring" style={{ border: '1px solid var(--border)' }} title="Notificações">
          <Bell size={18} color="var(--text-soft)" />
          {visibleInsights.length > 0 && <span className="absolute top-1.5 right-1.5 w-2 h-2 rounded-full" style={{ backgroundColor: 'var(--alert)' }} />}
        </button>
        <Popover open={notifOpen} onClose={() => setNotifOpen(false)} triggerRef={notifRef} width={288} align="center" className="p-2">
          {visibleInsights.length === 0 ? (
            <div className="px-3 py-4 text-center">
              <p className="text-sm" style={{ color: 'var(--text-soft)' }}>Nenhum aviso por enquanto.</p>
            </div>
          ) : (
            visibleInsights.map((n, i) => {
              const Icon = n.icon || Info;
              return (
                <div key={i} className="flex items-start gap-2.5 px-3 py-2.5 rounded-lg hover:bg-black/5">
                  <Icon size={15} className="mt-0.5 shrink-0" color="var(--text-soft)" />
                  <p className="text-sm flex-1" style={{ color: 'var(--text)' }}>{n.text}</p>
                  <button onClick={() => dismissNotification(n.text)} className="p-2 rounded-lg hover:bg-black/10 shrink-0" title="Dispensar">
                    <X size={13} color="var(--text-soft)" />
                  </button>
                </div>
              );
            })
          )}
        </Popover>
      </div>
    </header>
  );
}

function Banner({ insight, onDismiss }) {
  if (!insight) return null;
  const Icon = insight.icon;
  return (
    <div className="mx-4 md:mx-6 lg:mx-8 mt-3 sm:mt-4 mb-2 no-print">
      <div className="flex items-center gap-2.5 sm:gap-3 px-3.5 sm:px-4 py-2.5 sm:py-3 rounded-xl" style={{ backgroundColor: 'var(--primary-soft)' }}>
        <Icon size={16} color="var(--primary-dark)" className="shrink-0" />
        <p className="text-xs sm:text-sm flex-1 min-w-0" style={{ color: 'var(--primary-dark)' }}>{insight.text}</p>
        <button onClick={onDismiss} className="p-2 rounded-lg hover:bg-black/5 shrink-0" title="Dispensar"><X size={14} color="var(--primary-dark)" /></button>
      </div>
    </div>
  );
}

/* ============================================================
   DASHBOARD — cartões de indicadores
   ============================================================ */

function computeKPIs(period, customRange, mh, transactions, accounts, cards, goals, caixinhas, monthlySavingsTarget) {
  let receitas = 0, despesas = 0, periodStart, periodEnd;
  if (period === 'personalizado' && customRange.start && customRange.end) {
    periodStart = customRange.start; periodEnd = customRange.end;
    const filtered = transactions.filter((t) => isRealized(t) && t.date >= periodStart && t.date <= periodEnd);
    receitas = filtered.filter((t) => t.type === 'receita').reduce((s, t) => s + t.amount, 0);
    despesas = filtered.filter((t) => t.type === 'despesa').reduce((s, t) => s + t.amount, 0);
  } else {
    const monthsToTake = period === 'mes' ? 1 : period === 'trimestre' ? 3 : 12;
    const slice = mh.slice(-monthsToTake);
    receitas = slice.reduce((s, m) => s + m.receitas, 0);
    despesas = slice.reduce((s, m) => s + m.despesas, 0);
    const now = new Date();
    periodStart = ymd(new Date(now.getFullYear(), now.getMonth() - (monthsToTake - 1), 1));
    periodEnd = ymd(new Date(now.getFullYear(), now.getMonth() + 1, 0));
  }
  const saldoDisponivel = accounts.reduce((s, a) => s + a.balance, 0);
  // "Saldo previsto" soma ao saldo atual só o que AINDA NÃO afetou o saldo das contas. Receitas e
  // despesas comuns já contam aqui só quando "Pago" (isRealized), e nesse caso já foram aplicadas
  // à conta na hora — somar de novo em cima do saldo atual contaria esse valor duas vezes. A única
  // parte que realmente ainda não saiu de conta nenhuma são as despesas de cartão cuja fatura
  // ainda está em aberto (o débito só acontece quando a fatura é paga).
  const unpaidCardExpenses = transactions
    .filter((t) => t.type === 'despesa' && t.cardId && t.date >= periodStart && t.date <= periodEnd)
    .filter((t) => { const card = cards.find((c) => c.id === t.cardId); return card && (!card.paidThroughDate || t.date > card.paidThroughDate); })
    .reduce((s, t) => s + t.amount, 0);
  const saldoPrevisto = saldoDisponivel - unpaidCardExpenses;
  const patrimonio = mh.length ? mh[mh.length - 1].patrimonio : saldoDisponivel;
  const economiaAcumulada = caixinhas.reduce((s, c) => s + c.balance, 0) + goals.reduce((s, g) => s + g.current, 0);
  const metaMensalCurrent = mh.length ? thisMonthSaved(mh) : 0;
  const metaMensalTarget = monthlySavingsTarget > 0 ? monthlySavingsTarget : 0;
  const metaMensal = { current: metaMensalCurrent, target: metaMensalTarget, percent: metaMensalTarget > 0 ? Math.min(100, (metaMensalCurrent / metaMensalTarget) * 100) : 0 };

  const last = mh[mh.length - 1];
  const prev = mh.length > 1 ? mh[mh.length - 2] : null;
  const pctChange = (curr, before) => (before ? ((curr - before) / Math.abs(before)) * 100 : null);
  const growth = prev
    ? { receitas: pctChange(last.receitas, prev.receitas), despesas: pctChange(last.despesas, prev.despesas), patrimonio: pctChange(last.patrimonio, prev.patrimonio) }
    : { receitas: null, despesas: null, patrimonio: null };

  return { receitas, despesas, saldo: receitas - despesas, saldoDisponivel, saldoPrevisto, patrimonio, economiaAcumulada, metaMensal, growth };
}
function thisMonthSaved(mh) {
  const m = mh[mh.length - 1];
  return m.receitas - m.despesas;
}

function StatCard({ title, value, description, icon: Icon, color, soft, growth, progressPercent, expandableLabel, expandableContent, masked, onToggleMask }) {
  const [expanded, setExpanded] = useState(false);
  return (
    <Card padding="p-5" className="animate-fade-up">
      <div className="flex items-center justify-between mb-3 gap-2">
        <div className="flex items-center gap-2.5 min-w-0">
          <IconCircle icon={Icon} color={color} soft={soft} />
          <p className="text-xs font-medium truncate" style={{ color: 'var(--text-soft)' }}>{title}</p>
        </div>
        {growth != null && (
          <span className="flex items-center gap-0.5 text-xs font-semibold shrink-0" style={{ color: growth >= 0 ? 'var(--income)' : 'var(--expense)' }}>
            {growth >= 0 ? <ArrowUpRight size={14} /> : <ArrowDownRight size={14} />}
            {Math.abs(growth).toFixed(0)}%
          </span>
        )}
      </div>
      <div className="flex items-center justify-between gap-2 mb-1">
        <p className="font-display text-2xl font-bold tabular-nums truncate" style={{ color: 'var(--text)' }}>{masked ? '••••••' : value}</p>
        {onToggleMask && (
          <button onClick={onToggleMask} className="p-1.5 rounded-lg hover:bg-black/5 shrink-0" title={masked ? 'Mostrar saldo' : 'Ocultar saldo'}>
            {masked ? <EyeOff size={15} color="var(--text-soft)" /> : <Eye size={15} color="var(--text-soft)" />}
          </button>
        )}
      </div>
      {progressPercent != null && (
        <div className="my-2"><ProgressBar percent={progressPercent} color={color} /></div>
      )}
      {description && <p className="text-xs mt-1" style={{ color: 'var(--text-soft)' }}>{description}</p>}
      {expandableContent && (
        <>
          <button onClick={() => setExpanded((e) => !e)} className="text-xs font-medium mt-2 flex items-center gap-1" style={{ color }}>
            {expanded ? 'Ocultar' : (expandableLabel || 'Ver detalhes')} {expanded ? <ChevronUp size={12} /> : <ChevronDown size={12} />}
          </button>
          {expanded && <div className="mt-2 pt-2 space-y-1.5" style={{ borderTop: '1px solid var(--border)' }}>{expandableContent}</div>}
        </>
      )}
    </Card>
  );
}

function KPIRow({ kpis, settings, accounts, caixinhas = [], goals = [] }) {
  const [hideBalance, setHideBalance] = useState(false);
  const metaPercent = kpis.metaMensal.percent;
  const saldoBreakdown = accounts && accounts.length > 0 && (
    <>
      {accounts.map((a) => (
        <div key={a.id} className="flex items-center justify-between gap-2 text-xs">
          <span className="truncate" style={{ color: 'var(--text-soft)' }}>{a.bank}</span>
          <span className="tabular-nums font-medium shrink-0" style={{ color: a.balance < 0 ? 'var(--expense)' : 'var(--text)' }}>{hideBalance ? '••••••' : formatBRL(a.balance)}</span>
        </div>
      ))}
    </>
  );
  const totalCaixinhas = caixinhas.reduce((s, c) => s + c.balance, 0);
  const totalMetas = goals.reduce((s, g) => s + g.current, 0);
  const economiaBreakdown = (caixinhas.length > 0 || goals.length > 0) && (
    <>
      <div className="flex items-center justify-between gap-2 text-xs">
        <span style={{ color: 'var(--text-soft)' }}>Em caixinhas (separado de verdade)</span>
        <span className="tabular-nums font-medium shrink-0" style={{ color: 'var(--text)' }}>{formatBRL(totalCaixinhas)}</span>
      </div>
      <div className="flex items-center justify-between gap-2 text-xs">
        <span style={{ color: 'var(--text-soft)' }}>Em metas (acompanhamento)</span>
        <span className="tabular-nums font-medium shrink-0" style={{ color: 'var(--text)' }}>{formatBRL(totalMetas)}</span>
      </div>
    </>
  );
  const allCards = [
    { key: 'kpiSaldo', title: 'Saldo disponível', value: formatBRL(kpis.saldoDisponivel), description: 'Soma de todas as contas', icon: Wallet, color: 'var(--primary)', soft: 'var(--primary-soft)', expandableLabel: 'Ver saldo por conta', expandableContent: saldoBreakdown, masked: hideBalance, onToggleMask: () => setHideBalance((v) => !v) },
    { key: 'kpiReceitas', title: 'Receitas do período', value: formatBRL(kpis.receitas), description: 'Entradas no período selecionado', icon: TrendingUp, color: 'var(--income)', soft: 'var(--income-soft)', growth: kpis.growth.receitas },
    { key: 'kpiDespesas', title: 'Despesas do período', value: formatBRL(kpis.despesas), description: 'Saídas no período selecionado', icon: TrendingDown, color: 'var(--expense)', soft: 'var(--expense-soft)', growth: kpis.growth.despesas != null ? -kpis.growth.despesas : null },
    { key: 'kpiEconomia', title: 'Economia acumulada', value: formatBRL(kpis.economiaAcumulada), description: 'Em caixinhas + metas', icon: PiggyBank, color: 'var(--goals)', soft: 'var(--goals-soft)', expandableLabel: 'Ver caixinhas x metas', expandableContent: economiaBreakdown },
    { key: 'kpiMeta', title: 'Meta de economia', value: formatBRL(kpis.metaMensal.current), description: kpis.metaMensal.target > 0 ? `${metaPercent.toFixed(0)}% de ${formatBRL(kpis.metaMensal.target)} planejados` : 'defina uma meta em Configurações', icon: Target, color: 'var(--alert)', soft: 'var(--alert-soft)', progressPercent: metaPercent },
    { key: 'kpiPatrimonio', title: 'Patrimônio total', value: formatBRL(kpis.patrimonio), description: 'Contas + investimentos', icon: Landmark, color: 'var(--invest)', soft: 'var(--invest-soft)', growth: kpis.growth.patrimonio },
  ];
  const cards = allCards.filter((c) => isVisible(settings, 'dashboard', c.key));
  if (cards.length === 0) return null;
  return (
    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
      {cards.map((c) => <StatCard key={c.key} {...c} />)}
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
            <XAxis dataKey="month" tick={{ fontSize: 11, fill: 'var(--text-soft)' }} axisLine={{ stroke: 'var(--border)' }} tickLine={false} />
            <YAxis tick={{ fontSize: 11, fill: 'var(--text-soft)' }} axisLine={false} tickLine={false} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} />
            <Tooltip
              formatter={(value, name) => [formatBRL(value), name]}
              contentStyle={{ borderRadius: 12, border: '1px solid var(--border)', backgroundColor: 'var(--card)', color: 'var(--text)', fontSize: 12, fontFamily: 'Plus Jakarta Sans' }}
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
  const [activeIndex, setActiveIndex] = useState(null);
  if (entries.length === 0) {
    return (
      <Card className="animate-fade-up">
        <SectionTitle subtitle={subtitle}>{title}</SectionTitle>
        <EmptyState icon={PieChartIcon} title="Nenhum gasto registrado" description="Assim que você lançar despesas pagas neste período, elas aparecem aqui." />
      </Card>
    );
  }
  const active = activeIndex != null ? chartData[activeIndex] : null;
  return (
    <Card className="animate-fade-up">
      <SectionTitle subtitle={subtitle}>{title}</SectionTitle>
      <div className="flex flex-col sm:flex-row items-center gap-4">
        {/* Sem Tooltip flutuante de propósito — ela cobria o gráfico. O destaque aparece no
            centro do donut (que já é um espaço vazio) e na linha da legenda correspondente. */}
        <div style={{ width: 170, height: 170 }} className="shrink-0 relative">
          <ResponsiveContainer>
            <PieChart>
              <Pie
                data={chartData} dataKey="value" nameKey="name" innerRadius={52} outerRadius={78} paddingAngle={2} animationDuration={700}
                onClick={(_, i) => setActiveIndex((cur) => (cur === i ? null : i))}
              >
                {chartData.map((entry, i) => (
                  <Cell key={i} fill={entry.color} stroke="none" opacity={activeIndex == null || activeIndex === i ? 1 : 0.35} style={{ cursor: 'pointer', transition: 'opacity 0.15s ease' }} />
                ))}
              </Pie>
            </PieChart>
          </ResponsiveContainer>
          <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none px-2 text-center">
            <span className="text-[10px] leading-tight truncate max-w-[90px]" style={{ color: 'var(--text-soft)' }}>{active ? active.name : 'Total'}</span>
            <span className="font-display text-sm font-bold tabular-nums leading-tight" style={{ color: active ? active.color : 'var(--text)' }}>
              {formatBRL(active ? active.value : total)}
            </span>
          </div>
        </div>
        <div className="w-full space-y-1">
          {chartData.map((entry, i) => (
            <div
              key={i}
              onClick={() => setActiveIndex((cur) => (cur === i ? null : i))}
              className="flex items-center gap-2 text-sm rounded-lg px-2 py-1.5 cursor-pointer transition-colors"
              style={{ backgroundColor: activeIndex === i ? entry.color + '26' : 'transparent' }}
            >
              <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: entry.color }} />
              <span className="flex-1 truncate" style={{ color: 'var(--text)' }}>{entry.name}</span>
              <span className="tabular-nums font-medium" style={{ color: 'var(--text-soft)' }}>{total > 0 ? ((entry.value / total) * 100).toFixed(0) : 0}%</span>
              <span className="tabular-nums w-20 text-right shrink-0" style={{ color: 'var(--text)' }}>{formatBRL(entry.value)}</span>
            </div>
          ))}
        </div>
      </div>
    </Card>
  );
}

/* ---------- Transações recentes (resumo no dashboard) ---------- */

function RecentTransactions({ transactions, accounts, onEdit, onDelete, onSeeAll }) {
  // Só o que já aconteceu (ou vence hoje) — sem isso, as ocorrências futuras geradas
  // automaticamente pelas despesas recorrentes (até 12 meses à frente) dominariam a lista,
  // já que ela ordena por data mais recente primeiro.
  const todayStr = ymd(new Date());
  const recent = transactions.filter((t) => t.date <= todayStr).sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 7);
  const [confirmDeleteTx, setConfirmDeleteTx] = useState(null);
  return (
    <Card className="animate-fade-up" padding="p-5 sm:p-6">
      <SectionTitle action={<button onClick={onSeeAll} className="text-xs font-medium flex items-center gap-1" style={{ color: 'var(--primary)' }}>Ver todas <ArrowRight size={12} /></button>}>
        Transações recentes
      </SectionTitle>
      {recent.length === 0 ? (
        <EmptyState icon={ArrowLeftRight} title="Nenhum lançamento ainda" description="Seus lançamentos mais recentes aparecem aqui." />
      ) : (
        <>
          {/* Mobile: swipe pra esquerda revela editar/excluir, como nas outras listas de lançamentos. */}
          <div className="sm:hidden space-y-2">
            {recent.map((tx) => {
              const cat = CATEGORIES[tx.category] || CATEGORIES['Outros'];
              const inst = getInstallmentDisplay(tx);
              return (
                <SwipeableRow
                  key={tx.id} onEdit={() => onEdit(tx)} onDelete={() => setConfirmDeleteTx(tx)}
                  deleteConfirm={{ title: 'Excluir lançamento', description: `Tem certeza que deseja excluir "${tx.description}"? Essa ação não pode ser desfeita.` }}
                >
                  <div className="flex items-center gap-3 p-2">
                    <IconCircle icon={cat.icon} color={cat.color} soft={cat.soft} size={36} />
                    <div className="min-w-0 flex-1">
                      <div className="flex items-center gap-1.5 min-w-0">
                        <p className="text-sm truncate min-w-0" style={{ color: 'var(--text)' }}>{inst.desc}</p>
                        {inst.count && (
                          <span className="shrink-0 text-[10px] font-medium tabular-nums px-1.5 py-0.5 rounded-md" style={{ backgroundColor: 'var(--primary-soft)', color: 'var(--primary-dark)' }}>
                            {inst.index}/{inst.count}
                          </span>
                        )}
                      </div>
                      <p className="text-xs" style={{ color: 'var(--text-soft)' }}>{formatDateShortYear(tx.date)}</p>
                    </div>
                    <span className="tabular-nums text-sm font-medium shrink-0" style={{ color: tx.type === 'receita' ? 'var(--income)' : 'var(--expense)' }}>
                      {tx.type === 'receita' ? '+' : '-'} {formatBRL(tx.amount)}
                    </span>
                  </div>
                </SwipeableRow>
              );
            })}
          </div>

          {/* Desktop/tablet: linha única com ícones de ação, como antes. */}
          <div className="hidden sm:block space-y-1">
            {recent.map((tx) => {
              const cat = CATEGORIES[tx.category] || CATEGORIES['Outros'];
              const inst = getInstallmentDisplay(tx);
              return (
                <div key={tx.id} className="group flex items-center gap-3 py-2 px-1 rounded-xl hover:bg-black/[0.02]">
                  <IconCircle icon={cat.icon} color={cat.color} soft={cat.soft} size={36} />
                  <div className="min-w-0 flex-1">
                    <div className="flex items-center gap-1.5 min-w-0">
                      <p className="text-sm truncate min-w-0" style={{ color: 'var(--text)' }}>{inst.desc}</p>
                      {inst.count && (
                        <span className="shrink-0 text-[10px] font-medium tabular-nums px-1.5 py-0.5 rounded-md" style={{ backgroundColor: 'var(--primary-soft)', color: 'var(--primary-dark)' }}>
                          {inst.index}/{inst.count}
                        </span>
                      )}
                    </div>
                    <p className="text-xs" style={{ color: 'var(--text-soft)' }}>{formatDateShortYear(tx.date)}</p>
                  </div>
                  <span className="tabular-nums text-sm font-medium shrink-0" style={{ color: tx.type === 'receita' ? 'var(--income)' : 'var(--expense)' }}>
                    {tx.type === 'receita' ? '+' : '-'} {formatBRL(tx.amount)}
                  </span>
                  <div className="flex items-center gap-1 shrink-0">
                    <button onClick={() => onEdit(tx)} className="p-2 rounded-lg hover:bg-black/5" title="Editar"><Pencil size={14} color="var(--text-soft)" /></button>
                    <button onClick={() => setConfirmDeleteTx(tx)} className="p-2 rounded-lg hover:bg-black/5" title="Excluir"><Trash2 size={14} color="var(--expense)" /></button>
                  </div>
                </div>
              );
            })}
          </div>
        </>
      )}
      {confirmDeleteTx && (
        <ConfirmModal
          title="Excluir lançamento"
          description={`Tem certeza que deseja excluir "${confirmDeleteTx.description}"? Essa ação não pode ser desfeita.`}
          onConfirm={() => { onDelete(confirmDeleteTx); setConfirmDeleteTx(null); }}
          onClose={() => setConfirmDeleteTx(null)}
        />
      )}
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
  const metaPercent = kpis.metaMensal.percent;
  const rows = [
    { label: 'Maior receita', value: maiorReceita ? formatBRL(maiorReceita.amount) : '—', sub: maiorReceita?.description },
    { label: 'Maior despesa', value: maiorDespesa ? formatBRL(maiorDespesa.amount) : '—', sub: maiorDespesa?.description },
    { label: 'Média diária de gastos', value: formatBRL(mediaDiaria) },
    { label: 'Economia acumulada', value: formatBRL(kpis.economiaAcumulada) },
    { label: 'Saldo previsto (fechamento)', value: formatBRL(kpis.saldoPrevisto) },
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

function GoalCard({ goal, onAddFunds, onEdit, onDelete, onCompleted, compact }) {
  const Icon = GOAL_ICONS[goal.icon] || Target;
  const percent = Math.min(100, (goal.current / goal.target) * 100);
  const [adding, setAdding] = useState(false);
  const [amount, setAmount] = useState(0);
  const [showHistory, setShowHistory] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const cardRef = useRef(null);
  const deadlineLabel = new Date(goal.deadline + 'T00:00:00').toLocaleDateString('pt-BR', { month: 'short', year: 'numeric' });
  const estimate = useMemo(() => estimateGoalCompletion(goal), [goal.current, goal.target, goal.history]);
  const history = goal.history || [];
  function confirmAddFunds() {
    if (amount > 0) {
      const wasComplete = goal.current >= goal.target;
      const willBeComplete = goal.current + amount >= goal.target;
      onAddFunds(goal, amount);
      // Comemora só na transição pra concluída — não a cada aporte numa meta que já bateu 100%.
      if (!wasComplete && willBeComplete && onCompleted) {
        const rect = cardRef.current?.getBoundingClientRect();
        onCompleted(rect ? { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 } : null);
      }
    }
    setAdding(false); setAmount(0);
  }
  return (
    <Card ref={cardRef} className="animate-fade-up" padding="p-5">
      <div className="flex items-start gap-3 mb-4">
        <IconCircle icon={Icon} color="var(--primary)" soft="var(--primary-soft)" size={36} />
        <div className="min-w-0 flex-1">
          <p className="font-medium text-sm truncate" style={{ color: 'var(--text)' }}>{goal.name}</p>
          <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Previsão: {deadlineLabel}</p>
        </div>
        {!compact && (
          <div className="flex items-center gap-1 shrink-0">
            {onEdit && <button onClick={() => onEdit(goal)} className="p-2 rounded-lg hover:bg-black/5" title="Editar"><Pencil size={14} color="var(--text-soft)" /></button>}
            <button onClick={() => setConfirmDelete(true)} className="p-2 rounded-lg hover:bg-black/5" title="Excluir"><Trash2 size={14} color="var(--text-soft)" /></button>
          </div>
        )}
      </div>
      <div className="flex items-end justify-between mb-2">
        <span className="font-display text-xl font-bold tabular-nums" style={{ color: 'var(--text)' }}>{formatBRL(goal.current)}</span>
        <span className="text-xs tabular-nums" style={{ color: 'var(--text-soft)' }}>de {formatBRL(goal.target)}</span>
      </div>
      <ProgressBar percent={percent} color="var(--primary)" />
      <div className="flex items-center justify-between mt-2">
        <span className="text-xs font-semibold" style={{ color: 'var(--primary-dark)' }}>{percent.toFixed(0)}% concluído</span>
        {!compact && !adding && (
          <button onClick={() => setAdding(true)} className="text-xs font-medium flex items-center gap-1" style={{ color: 'var(--primary)' }}>
            <Plus size={12} /> Adicionar valor
          </button>
        )}
      </div>
      {adding && (
        <div className="flex items-center gap-2 mt-3">
          <CurrencyInput value={amount} onChange={setAmount} />
          <Button size="sm" onClick={confirmAddFunds}>OK</Button>
          <button onClick={() => { setAdding(false); setAmount(0); }} className="p-2" title="Cancelar"><X size={14} color="var(--text-soft)" /></button>
        </div>
      )}
      {!compact && (
        <div className="mt-3 pt-3" style={{ borderTop: '1px solid var(--border)' }}>
          {estimate.status === 'ok' && (
            <p className="text-xs flex items-start gap-1.5" style={{ color: 'var(--text-soft)' }}>
              <TrendingUp size={13} className="mt-0.5 shrink-0" color="var(--primary)" />
              <span>No ritmo atual (~{formatBRL(estimate.avgMonthly)}/mês), a meta deve ser concluída em <strong style={{ color: 'var(--text)' }}>{estimate.etaLabel}</strong>.</span>
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
                      <span className="tabular-nums font-medium" style={{ color: 'var(--primary-dark)' }}>+{formatBRL(h.amount)}</span>
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

function GoalForm({ initial, onSave, onClose }) {
  const [form, setForm] = useState(initial || { name: '', target: 0, current: 0, deadline: '', icon: 'Wallet' });
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
    <Modal title={initial ? 'Editar meta' : 'Nova meta'} onClose={onClose}>
      <div className="space-y-4">
        <div>
          <FieldLabel error={errors.name}>Nome da meta</FieldLabel>
          <input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Viagem para a praia" />
          {errors.name && <p className="text-xs mt-1" style={{ color: 'var(--expense)' }}>{errors.name}</p>}
        </div>
        <div className={initial ? '' : 'grid grid-cols-1 sm:grid-cols-2 gap-3'}>
          <div>
            <FieldLabel error={errors.target}>Valor objetivo</FieldLabel>
            <CurrencyInput value={form.target} onChange={(v) => setForm({ ...form, target: v })} />
            {errors.target && <p className="text-xs mt-1" style={{ color: 'var(--expense)' }}>{errors.target}</p>}
          </div>
          {/* No modo edição não dá pra mexer direto no valor já guardado — isso é feito por
              "Adicionar valor" no card, que também registra no histórico de aportes. Editar aqui
              livremente desalinharia esse histórico do valor atual. */}
          {!initial && (
            <div>
              <FieldLabel>Valor já guardado</FieldLabel>
              <CurrencyInput value={form.current} onChange={(v) => setForm({ ...form, current: v })} />
            </div>
          )}
        </div>
        <div>
          <FieldLabel error={errors.deadline}>Previsão de conclusão</FieldLabel>
          <input type="date" value={form.deadline} onChange={(e) => setForm({ ...form, deadline: e.target.value })} className={inputClass} style={inputStyle} />
          {errors.deadline && <p className="text-xs mt-1" style={{ color: 'var(--expense)' }}>{errors.deadline}</p>}
        </div>
        <div>
          <FieldLabel>Ícone</FieldLabel>
          <div className="flex flex-wrap gap-2">
            {Object.entries(GOAL_ICONS).map(([key, Icon]) => (
              <button key={key} onClick={() => setForm({ ...form, icon: key })} className="p-2.5 rounded-xl" style={{ backgroundColor: form.icon === key ? 'var(--primary-soft)' : 'transparent', border: '1px solid var(--border)' }}>
                <Icon size={16} color={form.icon === key ? 'var(--primary)' : 'var(--text-soft)'} />
              </button>
            ))}
          </div>
        </div>
        <div className="flex justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={() => { if (validate()) onSave(form); }}>{initial ? 'Salvar alterações' : 'Criar meta'}</Button>
        </div>
      </div>
    </Modal>
  );
}

function GoalsSection({ goals, onAddFunds, onDelete, onCompleted, onSeeAll, compact }) {
  return (
    <Card className="animate-fade-up" padding="p-5 sm:p-6">
      <SectionTitle action={onSeeAll && <button onClick={onSeeAll} className="text-xs font-medium flex items-center gap-1" style={{ color: 'var(--primary)' }}>Ver todas <ArrowRight size={12} /></button>}>
        Metas financeiras
      </SectionTitle>
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-4">
        {goals.map((g) => <GoalCard key={g.id} goal={g} onAddFunds={onAddFunds} onDelete={onDelete} onCompleted={onCompleted} compact={compact} />)}
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

  // Prioriza os gastos e lançamentos de despesas do mês — inclusive os já pagos — para o
  // calendário refletir onde o dinheiro realmente foi, e não só a rotina de fatura do cartão.
  // O fechamento de fatura deixou de ser sinalizado aqui: é um evento interno da operadora,
  // não um gasto do usuário; o vencimento (que exige ação) continua aparecendo, só que em
  // segundo plano.
  const events = [];
  transactions.filter((t) => isSameMonth(t.date, year, month)).forEach((t) => {
    const d = new Date(t.date + 'T00:00:00');
    events.push({ day: d.getDate(), label: t.description, type: t.type === 'receita' ? 'recebimento' : 'gasto', amount: t.amount });
  });
  cards.forEach((c) => {
    const invoice = computeCardInvoice(c, transactions);
    const adjustedDue = getAdjustedDueDate(year, month, c.dueDay);
    if (adjustedDue.getMonth() === month) {
      const wasAdjusted = adjustedDue.getDate() !== Math.min(c.dueDay, new Date(year, month + 1, 0).getDate());
      events.push({ day: adjustedDue.getDate(), label: `Vencimento fatura ${c.bank}${wasAdjusted ? ' (antecipado p/ dia útil)' : ''}`, type: 'vencimento', amount: invoice });
    }
  });
  const typePriority = { gasto: 3, recebimento: 2, vencimento: 1 };
  events.sort((a, b) => a.day - b.day || typePriority[b.type] - typePriority[a.type]);
  const eventColors = { gasto: 'var(--expense)', recebimento: 'var(--income)', vencimento: 'var(--alert)' };
  // Quando um dia tem mais de um evento, o ponto no calendário usa a cor do tipo mais
  // relevante (gasto > recebimento > vencimento) — assim dias com despesa sempre se destacam.
  const dayDominantType = {};
  events.forEach((e) => {
    const current = dayDominantType[e.day];
    if (!current || typePriority[e.type] > typePriority[current]) dayDominantType[e.day] = e.type;
  });

  const now = new Date();
  const monthOffset = (year - now.getFullYear()) * 12 + (month - now.getMonth());

  return (
    <Card className="animate-fade-up">
      <SectionTitle>Calendário financeiro</SectionTitle>
      <div className="flex justify-center mb-3">
        <MonthNavigator
          label={capitalizeFirst(viewDate.toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' }))}
          monthOffset={monthOffset}
          onPrev={() => setViewDate(new Date(year, month - 1, 1))}
          onNext={() => setViewDate(new Date(year, month + 1, 1))}
          onToday={() => setViewDate(new Date(now.getFullYear(), now.getMonth(), 1))}
        />
      </div>
      <div className="grid grid-cols-7 gap-1 mb-3">
        {weekDays.map((d, i) => <div key={i} className="text-center text-[11px] font-medium py-1" style={{ color: 'var(--text-soft)' }}>{d}</div>)}
        {cells.map((c) => (
          <div key={c.key} className="aspect-square flex items-center justify-center relative">
            {c.day && (
              <span className="text-xs w-7 h-7 flex items-center justify-center rounded-full" style={{ color: 'var(--text)' }}>
                {c.day}
                {dayDominantType[c.day] && <span className="absolute bottom-0.5 w-1.5 h-1.5 rounded-full" style={{ backgroundColor: eventColors[dayDominantType[c.day]] }} />}
              </span>
            )}
          </div>
        ))}
      </div>
      <div className="flex items-center gap-3 mb-3 flex-wrap">
        <span className="flex items-center gap-1.5 text-[11px]" style={{ color: 'var(--text-soft)' }}><span className="w-1.5 h-1.5 rounded-full" style={{ backgroundColor: eventColors.gasto }} /> Gasto</span>
        <span className="flex items-center gap-1.5 text-[11px]" style={{ color: 'var(--text-soft)' }}><span className="w-1.5 h-1.5 rounded-full" style={{ backgroundColor: eventColors.recebimento }} /> Recebimento</span>
        <span className="flex items-center gap-1.5 text-[11px]" style={{ color: 'var(--text-soft)' }}><span className="w-1.5 h-1.5 rounded-full" style={{ backgroundColor: eventColors.vencimento }} /> Vencimento de fatura</span>
      </div>
      <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
        {events.map((e, i) => (
          <div key={i} className="flex items-center justify-between text-xs px-3 py-2 rounded-lg" style={{ backgroundColor: 'var(--bg)' }}>
            <div className="flex items-center gap-2 min-w-0">
              <span className="font-display font-semibold w-6 text-center shrink-0" style={{ color: 'var(--primary)' }}>{e.day}</span>
              <span className="truncate" style={{ color: 'var(--text)' }}>{e.label}</span>
            </div>
            <span className="tabular-nums font-medium shrink-0 ml-2" style={{ color: eventColors[e.type] }}>{formatBRL(e.amount)}</span>
          </div>
        ))}
        {events.length === 0 && <p className="text-xs text-center py-4" style={{ color: 'var(--text-soft)' }}>Nenhum evento próximo neste mês.</p>}
      </div>
    </Card>
  );
}

/* ---------- Cartões (preview) ---------- */

function PayInvoiceModal({ card, amount, accounts, onConfirm, onClose }) {
  const [accountId, setAccountId] = useState(card.accountId || '');
  return (
    <Modal title={`Pagar fatura — ${card.bank}`} onClose={onClose}>
      <div className="space-y-4">
        <div className="rounded-xl p-4 text-center" style={{ backgroundColor: 'var(--bg)' }}>
          <p className="text-xs mb-1" style={{ color: 'var(--text-soft)' }}>Valor da fatura</p>
          <p className="font-display text-2xl font-bold tabular-nums" style={{ color: 'var(--text)' }}>{formatBRL(amount)}</p>
        </div>
        <div>
          <FieldLabel>Debitar de qual conta?</FieldLabel>
          <Select value={accountId} onChange={(e) => setAccountId(e.target.value)} className="w-full px-3.5 py-2.5 rounded-xl text-base sm:text-sm focus-ring" style={inputStyle}>
            <option value="">Nenhuma (só marcar como paga)</option>
            {accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
          </Select>
          <p className="text-xs mt-1.5" style={{ color: 'var(--text-soft)' }}>
            {accountId ? 'O saldo dessa conta é descontado automaticamente — sem precisar atualizar o valor na mão.' : 'O saldo de nenhuma conta será alterado, só a fatura é marcada como paga.'}
          </p>
        </div>
        <div className="flex justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={() => { onConfirm(accountId || null); onClose(); }}>Confirmar pagamento</Button>
        </div>
      </div>
    </Modal>
  );
}

// Versão compacta do cartão pra aba Fatura mensal — uma linha só, em vez do card grande (esse
// já existe por inteiro na aba "Meus cartões", incluindo limite, vencimento e antecipar
// parcelas; repetir tudo aqui só ocupava espaço à toa). Clicar na linha seleciona/filtra por
// esse cartão só; clicar de novo tira a seleção. "Pagar fatura" continua disponível, só que
// como um botão pequeno embutido na própria linha.
function CardInvoiceRow({ card, transactions, accounts, selected, onToggleSelect, year, month, subview, onPayInvoice, gradient }) {
  const [showPayModal, setShowPayModal] = useState(false);
  const { cycleInvoice, displayInvoice, displayCount, isPayable } = useMemo(() => {
    // Sempre pelo CICLO da fatura do mês sendo navegado — a mesma conta que a lista de
    // lançamentos usa na subview "Fatura". Antes, no mês corrente (offset 0), essa linha
    // calculava a fatura "em aberto agora" (por data de vencimento), um número diferente do
    // total do mês sendo exibido sempre que o cartão já tinha fechado o ciclo daquele mês — daí
    // o valor não bater com "0 lançamentos" logo abaixo, e não mudar ao trocar de mês.
    const cycleItems = transactions.filter((t) => t.type === 'despesa' && t.cardId === card.id
      && (() => { const c = getCardInvoiceCycle(card, t.date); return c.year === year && c.month === month; })());
    const cycleInvoice = cycleItems.reduce((s, t) => s + t.amount, 0);
    // Na subview "Todas as despesas do mês" a lista abaixo usa mês calendário, não ciclo de
    // fatura (de propósito, pra misturar débito e crédito) — então essa linha do topo precisa
    // mostrar o mesmo total por mês calendário quando um cartão está selecionado ali, senão os
    // dois números não batem.
    const monthItems = subview === 'todas'
      ? transactions.filter((t) => t.type === 'despesa' && t.cardId === card.id && isSameMonth(t.date, year, month))
      : cycleItems;
    // "Pagar fatura" só faz sentido no mês em que a fatura REALMENTE em aberto (a próxima a
    // vencer) cai — e isso depende do dia de fechamento de CADA cartão, não do mês atual do
    // calendário: um cartão que já fechou pode ter a fatura aberta caindo no mês seguinte,
    // enquanto outro cartão ainda está com a fatura aberta no mês corrente.
    const dueCycle = getCardInvoiceCycle(card, ymd(getNextCardDueDate(card)));
    const isPayable = dueCycle.year === year && dueCycle.month === month;
    return { cycleInvoice, displayInvoice: monthItems.reduce((s, t) => s + t.amount, 0), displayCount: monthItems.length, isPayable };
  }, [card, transactions, year, month, subview]);

  return (
    <div
      role="button" tabIndex={0}
      onClick={onToggleSelect}
      onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onToggleSelect(); } }}
      className="w-full flex items-center gap-3 p-3 rounded-xl text-left cursor-pointer transition-colors hover:bg-black/[0.03] focus-ring"
      style={{ backgroundColor: 'var(--card)', border: selected ? '2px solid var(--primary)' : '1px solid var(--border)' }}
    >
      <div className="rounded-xl flex items-center justify-center shrink-0" style={{ width: 36, height: 36, background: gradient }}>
        <CreditCard size={16} color="#fff" strokeWidth={2.2} />
      </div>
      <div className="flex-1 min-w-0">
        <p className="text-sm font-medium truncate" style={{ color: 'var(--text)' }}>{card.bank}</p>
        <p className="text-xs truncate" style={{ color: 'var(--text-soft)' }}>{card.brand} · {displayCount} lançamento{displayCount === 1 ? '' : 's'}</p>
      </div>
      <div className="text-right shrink-0">
        <p className="text-sm font-semibold tabular-nums" style={{ color: 'var(--text)' }}>{formatBRL(displayInvoice)}</p>
        {subview === 'fatura' && isPayable && (cycleInvoice > 0 ? (
          <button onClick={(e) => { e.stopPropagation(); setShowPayModal(true); }} className="text-xs font-medium px-2.5 py-1 rounded-lg transition-colors hover:opacity-80" style={{ backgroundColor: 'var(--primary-soft)', color: 'var(--primary-dark)' }}>Pagar fatura</button>
        ) : (
          <span className="text-xs font-medium px-2.5 py-1 rounded-lg flex items-center justify-end gap-1" style={{ backgroundColor: 'var(--income-soft)', color: 'var(--income)' }}><Check size={10} /> Em dia</span>
        ))}
      </div>
      {showPayModal && (
        <PayInvoiceModal card={card} amount={cycleInvoice} accounts={accounts} onConfirm={(accountId) => { onPayInvoice(card, cycleInvoice, accountId); setShowPayModal(false); }} onClose={() => setShowPayModal(false)} />
      )}
    </div>
  );
}

function CreditCardVisual({ card, transactions, accounts = [], gradient, onPayInvoice, onAdvanceInstallments, onEdit, onDelete, viewedCycle }) {
  // viewedCycle: opcional { year, month, label } — quando presente (navegação por mês na aba
  // Fatura), mostra o total da fatura daquele ciclo específico, não necessariamente a fatura
  // real em aberto agora. Nesse modo a visualização fica mais enxuta: pagar fatura, antecipar
  // parcelas e vencimento/limite são conceitos do "agora", não fazem sentido pra um mês
  // navegado (passado ou futuro), então ficam ocultos.
  const isHistoricalView = !!viewedCycle;
  const invoice = useMemo(() => {
    if (isHistoricalView) {
      return transactions
        .filter((t) => t.type === 'despesa' && t.cardId === card.id)
        .filter((t) => { const c = getCardInvoiceCycle(card, t.date); return c.year === viewedCycle.year && c.month === viewedCycle.month; })
        .reduce((s, t) => s + t.amount, 0);
    }
    return computeCardInvoice(card, transactions);
  }, [card, transactions, isHistoricalView, viewedCycle?.year, viewedCycle?.month]);
  const invoiceCount = useMemo(() => {
    if (!isHistoricalView) return 0;
    return transactions.filter((t) => t.type === 'despesa' && t.cardId === card.id && (() => { const c = getCardInvoiceCycle(card, t.date); return c.year === viewedCycle.year && c.month === viewedCycle.month; })()).length;
  }, [card, transactions, isHistoricalView, viewedCycle?.year, viewedCycle?.month]);
  const percentUsed = card.limit > 0 ? (invoice / card.limit) * 100 : 0;
  const barColor = percentUsed > 85 ? 'var(--expense)' : percentUsed > 65 ? 'var(--alert)' : 'var(--income)';
  const nextDue = useMemo(() => getNextCardDueDate(card), [card.dueDay]);
  const dueWasAdjusted = nextDue.getDate() !== card.dueDay;
  const nextDueLabel = nextDue.toLocaleDateString('pt-BR', { day: '2-digit', month: 'short' });
  const [showPayModal, setShowPayModal] = useState(false);
  const [confirmAdvance, setConfirmAdvance] = useState(false);
  const [showEditForm, setShowEditForm] = useState(false);
  const openClosing = useMemo(() => getOpenInvoiceClosingDate(card), [card.dueDay, card.closingDay]);
  const futureInstallments = useMemo(() => {
    const cutoff = ymd(openClosing);
    return transactions.filter((t) => t.cardId === card.id && t.installmentGroupId && t.date > cutoff);
  }, [card.id, transactions, openClosing]);
  const futureTotal = futureInstallments.reduce((s, t) => s + t.amount, 0);
  const pendingInvoiceCount = useMemo(() => transactions.filter((t) => t.cardId === card.id && t.status === 'Pendente').length, [transactions, card.id]);
  return (
    <div className="rounded-2xl p-5 text-white shadow-soft-lg" style={{ background: gradient }}>
      <div className="flex items-start justify-between mb-6">
        <div>
          <p className="text-sm font-medium opacity-90">{card.bank}</p>
          <p className="text-xs opacity-70">{card.brand}</p>
        </div>
        {onEdit && !isHistoricalView && (
          <button onClick={() => setShowEditForm(true)} className="p-2 rounded-lg hover:bg-white/10" title="Editar cartão">
            <Pencil size={16} className="opacity-80" />
          </button>
        )}
      </div>
      <div className="flex items-end justify-between mb-4">
        <div>
          <p className="text-xs opacity-70 mb-1">{isHistoricalView ? `Fatura de ${viewedCycle.label}` : 'Fatura atual'}</p>
          <p className="font-display text-2xl font-bold tabular-nums">{formatBRL(invoice)}</p>
        </div>
        {!isHistoricalView && (invoice > 0 ? (
          <button onClick={() => setShowPayModal(true)} className="text-xs font-medium px-2.5 py-1.5 rounded-lg bg-white/15 hover:bg-white/25 transition-colors shrink-0">
            Pagar fatura
          </button>
        ) : (
          <span className="text-xs font-medium px-2.5 py-1.5 rounded-lg bg-white/10 flex items-center gap-1 shrink-0"><Check size={12} /> Em dia</span>
        ))}
      </div>
      {isHistoricalView ? (
        <p className="text-xs opacity-70 pt-3" style={{ borderTop: '1px solid rgba(255,255,255,0.2)' }}>
          {invoiceCount} lançamento{invoiceCount === 1 ? '' : 's'} nessa fatura
        </p>
      ) : (
        <>
          <div className="mb-2">
            <div className="w-full h-1.5 rounded-full bg-white/20 overflow-hidden">
              <div className="h-full rounded-full" style={{ width: `${Math.min(100, percentUsed)}%`, backgroundColor: barColor }} />
            </div>
          </div>
          <div className="flex items-center justify-between text-xs opacity-80">
            <span>{formatBRL(invoice)} de {formatBRL(card.limit)}</span>
            <span>{percentUsed.toFixed(0)}% usado</span>
          </div>
          <div className="flex items-center justify-between text-xs opacity-70 mt-3 pt-3" style={{ borderTop: '1px solid rgba(255,255,255,0.2)' }}>
            <span>Fecha dia {card.closingDay}</span>
            <span className="text-right">
              Próx. vencimento: {nextDueLabel}
              {dueWasAdjusted && <span className="block opacity-70">(dia {card.dueDay}, antecipado p/ dia útil)</span>}
            </span>
          </div>
          {card.paidThroughDate && (
            <p className="text-[11px] opacity-60 mt-2">Fatura paga até {formatDate(card.paidThroughDate)}</p>
          )}
          {futureInstallments.length > 0 && onAdvanceInstallments && (
            <button onClick={() => setConfirmAdvance(true)} className="w-full mt-3 pt-3 text-xs text-left flex items-center justify-between gap-2" style={{ borderTop: '1px solid rgba(255,255,255,0.2)' }}>
              <span className="opacity-80">{futureInstallments.length} parcela(s) futura(s) — {formatBRL(futureTotal)}</span>
              <span className="font-medium underline shrink-0">Antecipar</span>
            </button>
          )}
        </>
      )}
      {showPayModal && (
        <PayInvoiceModal card={card} amount={invoice} accounts={accounts} onConfirm={(accountId) => onPayInvoice(card, invoice, accountId)} onClose={() => setShowPayModal(false)} />
      )}
      {showEditForm && (
        <CardForm initial={card} accounts={accounts} pendingInvoiceCount={pendingInvoiceCount} onSave={(f) => { onEdit(f); setShowEditForm(false); }} onDelete={onDelete} onClose={() => setShowEditForm(false)} />
      )}
      {confirmAdvance && (
        <ConfirmModal
          title="Antecipar parcelas futuras"
          description={`Isso traz ${futureInstallments.length} parcela(s) que ainda não venceram (totalizando ${formatBRL(futureTotal)}) para a fatura atual, quitando o restante da compra de uma vez.`}
          variant="primary" confirmLabel="Antecipar"
          onConfirm={() => onAdvanceInstallments(card)}
          onClose={() => setConfirmAdvance(false)}
        />
      )}
    </div>
  );
}

// Gera 5 tonalidades do cartão a partir da cor de destaque escolhida em Configurações — sempre
// misturando com quase-preto (mantém o visual "cartão premium escuro" dos gradientes fixos
// anteriores) em proporções crescentes, para que qualquer cor de destaque (incluindo o tema
// Monocromático, que na prática é um cinza bem escuro) produza uma família de tons reconhecível
// e nunca vire um cinza genérico igual pra todo mundo.
function buildCardGradients(colorThemeKey) {
  const theme = COLOR_THEMES[colorThemeKey] || COLOR_THEMES.default;
  const base = theme.light; // sempre a variante clara/mais saturada, independente do tema claro/escuro do app
  // Faixa mais larga que antes (0.14–0.86 em vez de 0.28–0.78) pra aumentar o contraste entre os
  // tons extremos. Os 5 tons são gerados em ordem clara→escura e depois reordenados (1,3,5,2,4)
  // pra que cartões vizinhos na grade nunca fiquem com tonalidades quase iguais lado a lado.
  const tones = [0.14, 0.32, 0.50, 0.68, 0.86].map((amt) => {
    const from = mixHex(base, '#0A0A0A', amt);
    const to = mixHex(base, '#0A0A0A', Math.min(0.95, amt + 0.24));
    return `linear-gradient(135deg, ${from} 0%, ${to} 100%)`;
  });
  return [tones[0], tones[2], tones[4], tones[1], tones[3]];
}

function CardsPreview({ cards, transactions, accounts, cardGradients, onPayInvoice, onAdvanceInstallments, onSeeAll }) {
  return (
    <Card className="animate-fade-up" padding="p-5 sm:p-6">
      <SectionTitle action={<button onClick={onSeeAll} className="text-xs font-medium flex items-center gap-1" style={{ color: 'var(--primary)' }}>Ver todos <ArrowRight size={12} /></button>}>
        Cartões de crédito
      </SectionTitle>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        {cards.map((c, i) => <CreditCardVisual key={c.id} card={c} transactions={transactions} accounts={accounts} onPayInvoice={onPayInvoice} onAdvanceInstallments={onAdvanceInstallments} gradient={cardGradients[i % cardGradients.length]} />)}
      </div>
    </Card>
  );
}

/* ---------- Assinaturas / Despesas recorrentes (preview) ---------- */

function RecurringPreview({ recurring, onSeeAll }) {
  const total = recurring.reduce((s, r) => s + r.value, 0);
  return (
    <Card className="animate-fade-up" padding="p-5 sm:p-6">
      <SectionTitle subtitle={`${formatBRL(total)}/mês`} action={<button onClick={onSeeAll} className="text-xs font-medium flex items-center gap-1" style={{ color: 'var(--primary)' }}>Ver todas <ArrowRight size={12} /></button>}>
        Despesas recorrentes
      </SectionTitle>
      <div className="space-y-3">
        {recurring.slice(0, 5).map((r) => {
          const cat = CATEGORIES[r.category] || CATEGORIES['Outros'];
          return (
            <div key={r.id} className="flex items-center gap-3">
              <IconCircle icon={cat.icon} color={cat.color} soft={cat.soft} size={36} />
              <div className="flex-1 min-w-0">
                <p className="text-sm font-medium truncate" style={{ color: 'var(--text)' }}>{r.name}</p>
                <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Renova dia {r.renewalDay}</p>
              </div>
              <span className="text-sm tabular-nums font-medium" style={{ color: 'var(--expense)' }}>- {formatBRL(r.value)}</span>
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
  const { settings } = data;
  const kpis = computeKPIs(data.period, data.customRange, data.monthlyHistory, data.transactions, data.accounts, data.cards, data.goals, data.caixinhas, data.settings.monthlySavingsTarget);
  const v = (item) => isVisible(settings, 'dashboard', item);
  const allBlocks = ['kpiSaldo', 'kpiReceitas', 'kpiDespesas', 'kpiEconomia', 'kpiMeta', 'kpiPatrimonio', 'evolutionChart', 'categoryDonut', 'recentTransactions', 'monthSummary', 'goalsSection', 'financialCalendar', 'cardsPreview', 'recurringPreview', 'investmentsPreview', 'insights'];
  if (allBlocks.every((b) => !v(b))) {
    return (
      <Card>
        <EmptyState icon={LayoutDashboard} title="Nenhum bloco selecionado" description='Todos os blocos do Dashboard estão ocultos. Vá em Configurações → "Personalização da interface" para reativar algum.' />
      </Card>
    );
  }
  const showKPIRow = ['kpiSaldo', 'kpiReceitas', 'kpiDespesas', 'kpiEconomia', 'kpiMeta', 'kpiPatrimonio'].some(v);
  const showChartsRow = v('evolutionChart') || v('categoryDonut');
  const showMiddleRow = v('recentTransactions') || v('monthSummary');
  const showCalendarCardsRow = v('financialCalendar') || v('cardsPreview');
  const showPreviewRow = v('recurringPreview') || v('investmentsPreview');
  return (
    <div className="space-y-6">
      {showKPIRow && <KPIRow kpis={kpis} settings={settings} accounts={data.accounts} caixinhas={data.caixinhas} goals={data.goals} />}
      {showChartsRow && (
        <div className="grid grid-cols-1 lg:grid-cols-5 gap-6">
          {v('evolutionChart') && <div className={v('categoryDonut') ? 'lg:col-span-3' : 'lg:col-span-5'}><EvolutionChart data={data.monthlyHistory} /></div>}
          {v('categoryDonut') && <div className={v('evolutionChart') ? 'lg:col-span-2' : 'lg:col-span-5'}><CategoryDonut data={data.categoryTotalsForPeriod} subtitle="no período selecionado" /></div>}
        </div>
      )}
      {showMiddleRow && (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {v('recentTransactions') && (
            <div className={v('monthSummary') ? 'lg:col-span-2' : 'lg:col-span-3'}>
              <RecentTransactions transactions={data.transactions} accounts={data.accounts} onEdit={actions.editTransaction} onDelete={actions.deleteTransaction} onSeeAll={() => actions.goTo('transacoes')} />
            </div>
          )}
          {v('monthSummary') && <div className={v('recentTransactions') ? '' : 'lg:col-span-3'}><MonthSummaryPanel transactions={data.transactions} kpis={kpis} /></div>}
        </div>
      )}
      {v('goalsSection') && <GoalsSection goals={data.goals} onAddFunds={actions.addGoalFunds} onDelete={actions.deleteGoal} onSeeAll={() => actions.goTo('metas')} compact />}
      {showCalendarCardsRow && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {v('financialCalendar') && <div className={v('cardsPreview') ? '' : 'lg:col-span-2'}><FinancialCalendar cards={data.cards} transactions={data.transactions} /></div>}
          {v('cardsPreview') && <div className={v('financialCalendar') ? '' : 'lg:col-span-2'}><CardsPreview cards={data.cards} transactions={data.transactions} accounts={data.accounts} cardGradients={data.cardGradients} onPayInvoice={actions.payCardInvoice} onAdvanceInstallments={actions.advanceAllFutureInstallments} onSeeAll={() => actions.goTo('cartoes')} /></div>}
        </div>
      )}
      {showPreviewRow && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {v('recurringPreview') && <div className={v('investmentsPreview') ? '' : 'lg:col-span-2'}><RecurringPreview recurring={data.recurring} onSeeAll={() => actions.goTo('recorrentes')} /></div>}
          {v('investmentsPreview') && <div className={v('recurringPreview') ? '' : 'lg:col-span-2'}><InvestmentsPreview investments={data.investments} total={data.investmentsTotal} onSeeAll={() => actions.goTo('investimentos')} /></div>}
        </div>
      )}
      {v('insights') && <InsightsSection insights={data.insights} />}
    </div>
  );
}

/* ============================================================
   FORMULÁRIO DE TRANSAÇÃO
   ============================================================ */

function TransactionForm({ initial, accounts, cards, benefits = [], transactions = [], onSave, onClose, onDelete }) {
  const [form, setForm] = useState(initial || {
    type: 'despesa', description: '', amount: 0, category: 'Mercado', account: accounts[0]?.id || '',
    paymentMethod: 'Pix', date: new Date().toISOString().slice(0, 10), status: 'Pago',
    isSalary: false, grossSalary: 0, dependents: 0, cardId: null, benefitId: null, benefitType: null,
  });
  const [errors, setErrors] = useState({});
  const [shakeError, setShakeError] = useState(false);
  const [flashField, setFlashField] = useState(null);
  const [duplicateWarning, setDuplicateWarning] = useState(null);
  const [autofillNotice, setAutofillNotice] = useState(null);
  const [sessionPatternNotice, setSessionPatternNotice] = useState(null);
  const [categoryTouched, setCategoryTouched] = useState(false);
  const [paymentTouched, setPaymentTouched] = useState(false);
  const descRef = useRef(null);
  const salaryRef = useRef(null);
  const amountRef = useRef(null);
  const dateRef = useRef(null);
  const accountRef = useRef(null);
  // Parcelamento é oferecido tanto ao criar um lançamento novo quanto ao editar um que ainda não
  // é parcela de um grupo (convertendo uma compra à vista em parcelada). Editar uma parcela que
  // já existe dentro de um grupo continua editando só aquela ocorrência — não faz sentido reabrir
  // o plano inteiro por aqui.
  const [installmentEnabled, setInstallmentEnabled] = useState(false);
  const [installmentCount, setInstallmentCount] = useState(2);
  const [installmentCountText, setInstallmentCountText] = useState('2');
  const [installmentTotal, setInstallmentTotal] = useState(0);

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

  // "Inércia de sessão": quando os últimos lançamentos criados foram todos na mesma forma de
  // pagamento (ex: lançando um extrato inteiro de um cartão, um por um), presume que o próximo
  // segue o mesmo padrão e já pré-seleciona assim que o formulário abre, sem precisar digitar
  // nada. Roda só uma vez, na criação de um lançamento novo — o autofill por nome parecido (mais
  // específico) ainda pode sobrescrever isso depois, já que essa presunção não marca o campo
  // como "tocado" pelo usuário.
  useEffect(() => {
    if (initial) return;
    function paymentSignature(t) {
      if (t.cardId) return `card:${t.cardId}`;
      if (t.benefitId) return `benefit:${t.benefitId}`;
      return `account:${t.account}:${t.paymentMethod}`;
    }
    const recent = transactions.filter((t) => t.type === form.type).slice(0, 3);
    if (recent.length < 3) return;
    const sig = paymentSignature(recent[0]);
    if (!recent.every((t) => paymentSignature(t) === sig)) return;
    const last = recent[0];
    setForm((f) => ({
      ...f,
      category: f.category,
      paymentMethod: last.paymentMethod,
      account: last.account || f.account,
      cardId: last.cardId ?? null,
      benefitId: last.benefitId ?? null,
      benefitType: last.benefitType ?? null,
    }));
    const card = cards.find((c) => c.id === last.cardId);
    const account = accounts.find((a) => a.id === last.account);
    const label = card ? card.bank : account ? account.bank : last.paymentMethod;
    setSessionPatternNotice(`Seus últimos lançamentos foram em ${label} — deixei pré-selecionado, é só trocar se for diferente desta vez.`);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Autofill: ao digitar a descrição, procura no histórico um lançamento com nome parecido e
  // preenche sozinho a categoria e a forma de pagamento usadas da última vez. Reaproveita o
  // mesmo textSimilarity já usado na detecção de duplicata da importação.
  useEffect(() => {
    if (initial || (categoryTouched && paymentTouched)) { setAutofillNotice(null); return; }
    const q = form.description.trim();
    if (q.length < 3) { setAutofillNotice(null); return; }
    const id = setTimeout(() => {
      let best = null; let bestScore = 0;
      transactions.forEach((t) => {
        if (t.type !== form.type || !t.description || t.description.trim().toLowerCase() === q.toLowerCase()) return;
        const score = textSimilarity(q, t.description);
        if (score > bestScore) { bestScore = score; best = t; }
      });
      if (bestScore < 0.5 || !best) { setAutofillNotice(null); return; }
      if (!paymentTouched) setSessionPatternNotice(null);
      setForm((f) => ({
        ...f,
        category: categoryTouched ? f.category : best.category,
        paymentMethod: paymentTouched ? f.paymentMethod : best.paymentMethod,
        account: paymentTouched ? f.account : (best.account || f.account),
        cardId: paymentTouched ? f.cardId : (best.cardId ?? null),
        benefitId: paymentTouched ? f.benefitId : (best.benefitId ?? null),
        benefitType: paymentTouched ? f.benefitType : (best.benefitType ?? null),
      }));
      const filledWhat = !categoryTouched && !paymentTouched ? 'categoria e forma de pagamento'
        : !categoryTouched ? 'categoria' : !paymentTouched ? 'forma de pagamento' : null;
      setAutofillNotice(filledWhat ? `${filledWhat} preenchida com base em "${best.description}"` : null);
    }, 400);
    return () => clearTimeout(id);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [form.description, form.type, categoryTouched, paymentTouched]);

  const selectedCard = cards.find((c) => c.id === form.cardId);
  const selectedBenefit = benefits.find((b) => b.id === form.benefitId);
  const perInstallment = installmentCount > 0 ? installmentTotal / installmentCount : 0;
  const isEditingInstallment = !!initial?.installmentGroupId;
  const canToggleInstallments = !isEditingInstallment;
  const currentInvoiceCycle = useMemo(
    () => (isEditingInstallment && selectedCard ? getCardInvoiceCycle(selectedCard, form.date) : null),
    [isEditingInstallment, selectedCard, form.date]
  );

  function validate() {
    const e = {};
    if (!form.description.trim()) e.description = 'Informe uma descrição';
    if (installmentEnabled) {
      if (!installmentTotal || installmentTotal <= 0) e.amount = 'Informe o valor total da compra';
    } else if (!form.amount || form.amount <= 0) e.amount = 'Informe um valor maior que zero';
    if (!form.date) e.date = 'Informe uma data';
    if (!form.account) e.account = 'Selecione uma conta';
    if (form.isSalary && (!form.grossSalary || form.grossSalary <= 0)) e.grossSalary = 'Informe o salário bruto';
    setErrors(e);
    return e;
  }

  function triggerShake() {
    setShakeError(true);
    setTimeout(() => setShakeError(false), 450);
  }

  function scrollToFirstError(e) {
    const order = [['description', descRef], ['grossSalary', salaryRef], ['amount', amountRef], ['date', dateRef], ['account', accountRef]];
    const first = order.find(([key]) => e[key]);
    if (first && first[1].current) {
      first[1].current.scrollIntoView({ behavior: 'smooth', block: 'center' });
      setFlashField(first[0]);
      setTimeout(() => setFlashField(null), 1100);
    }
  }

  function doSave() {
    if (installmentEnabled && canToggleInstallments) {
      const count = Math.max(2, Math.min(48, Number(installmentCountText) || 2));
      onSave({ ...form, amount: perInstallment, installmentPlan: { count, totalAmount: installmentTotal } });
    } else {
      onSave(form);
    }
  }

  function handleSubmit() {
    const e = validate();
    if (Object.keys(e).length > 0) { triggerShake(); scrollToFirstError(e); return; }
    // Só checa duplicata em lançamento novo — editar um já existente não é uma inclusão.
    if (!initial && !duplicateWarning) {
      const amountToCheck = installmentEnabled ? perInstallment : form.amount;
      const pool = transactions.map((t) => ({ date: t.date, description: t.description, amount: t.amount, used: false }));
      const match = findBestDuplicateMatch({ date: form.date, description: form.description, amount: amountToCheck }, pool);
      if (match) { setDuplicateWarning(match); return; }
    }
    doSave();
  }

  function selectPaymentSource(v) {
    setPaymentTouched(true);
    if (v.startsWith('account:')) {
      const accountId = v.slice(8);
      setForm({ ...form, account: accountId, cardId: null, benefitId: null, benefitType: null, status: form.status === 'Pendente' && form.cardId ? 'Pago' : form.status });
      setInstallmentEnabled(false);
    } else if (v.startsWith('card:')) {
      const cardId = v.slice(5);
      const card = cards.find((c) => c.id === cardId);
      setForm({ ...form, cardId, benefitId: null, benefitType: null, account: card ? card.accountId : form.account, paymentMethod: 'Cartão de crédito', status: 'Pendente' });
    } else {
      const [, benefitId, benefitType] = v.split(':');
      const benefit = benefits.find((b) => b.id === benefitId);
      setForm({ ...form, cardId: null, benefitId, benefitType, account: benefit ? form.account : form.account, paymentMethod: 'Vale-benefícios' });
      setInstallmentEnabled(false);
    }
  }

  return (
    <Modal title={initial ? 'Editar lançamento' : 'Novo lançamento'} onClose={onClose} wide>
      <div className="space-y-4">
        <div className="flex gap-2">
          {['despesa', 'receita'].map((t) => (
            <button
              key={t} onClick={() => setForm({ ...form, type: t, isSalary: t === 'receita' ? form.isSalary : false, cardId: t === 'receita' ? null : form.cardId, benefitId: t === 'receita' ? null : form.benefitId, benefitType: t === 'receita' ? null : form.benefitType })}
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
          <label className="flex items-center gap-2.5 text-sm cursor-pointer select-none py-1" style={{ color: 'var(--text)' }}>
            <input type="checkbox" checked={!!form.isSalary} onChange={(e) => setForm({ ...form, isSalary: e.target.checked, description: e.target.checked && !form.description ? 'Salário' : form.description })} className="w-5 h-5 shrink-0 focus-ring" />
            É salário (CLT)? Calculamos o líquido estimado com INSS e IRRF.
          </label>
        )}
        <div ref={descRef} className={flashField === 'description' ? 'rounded-xl animate-field-flash' : ''}>
          <FieldLabel error={errors.description}>Descrição</FieldLabel>
          <input value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Supermercado" />
          {errors.description && <p className="text-xs mt-1" style={{ color: 'var(--expense)' }}>{errors.description}</p>}
          {autofillNotice && <p className="text-xs mt-1.5" style={{ color: 'var(--primary)' }}>{autofillNotice}</p>}
        </div>

        {form.isSalary && (
          <div className="rounded-xl p-4 space-y-3" style={{ backgroundColor: 'var(--income-soft)' }}>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div ref={salaryRef} className={flashField === 'grossSalary' ? 'rounded-xl animate-field-flash' : ''}>
                <FieldLabel error={errors.grossSalary}>Salário bruto</FieldLabel>
                <CurrencyInput value={form.grossSalary} onChange={(v) => setForm({ ...form, grossSalary: v })} />
              </div>
              <div>
                <FieldLabel>Dependentes (IR)</FieldLabel>
                <input type="number" inputMode="numeric" min={0} value={form.dependents} onChange={(e) => setForm({ ...form, dependents: Number(e.target.value) })} className={inputClass} style={inputStyle} />
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

        {isEditingInstallment && (
          <div className="rounded-xl p-3 text-xs flex items-center gap-2" style={{ backgroundColor: 'var(--primary-soft)', color: 'var(--primary-dark)' }}>
            <CreditCard size={14} className="shrink-0" />
            Parcela {initial.installmentIndex} de {initial.installmentCount} desta compra parcelada. Editar aqui altera só esta parcela — pra antecipar as parcelas restantes, use o cartão em Cartões.
          </div>
        )}

        {isEditingInstallment && selectedCard && currentInvoiceCycle && (
          <div className="rounded-xl p-3 text-xs space-y-2" style={{ backgroundColor: 'var(--bg)', border: '1px solid var(--border)' }}>
            <p style={{ color: 'var(--text-soft)' }}>
              Está na fatura de <strong style={{ color: 'var(--text)' }}>{capitalizeFirst(new Date(currentInvoiceCycle.year, currentInvoiceCycle.month, 1).toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' }))}</strong>. Se o banco colocou essa compra num mês diferente do esperado (comum perto do fechamento), dá pra mover a parcela:
            </p>
            <div className="flex gap-2">
              <button type="button" onClick={() => setForm({ ...form, date: shiftToAdjacentInvoiceCycle(selectedCard, form.date, -1) })} className="flex-1 text-xs font-medium px-2.5 py-1.5 rounded-lg flex items-center justify-center gap-1" style={{ color: 'var(--primary)', backgroundColor: 'var(--primary-soft)' }}>
                <ChevronLeft size={12} /> Fatura anterior
              </button>
              <button type="button" onClick={() => setForm({ ...form, date: shiftToAdjacentInvoiceCycle(selectedCard, form.date, 1) })} className="flex-1 text-xs font-medium px-2.5 py-1.5 rounded-lg flex items-center justify-center gap-1" style={{ color: 'var(--primary)', backgroundColor: 'var(--primary-soft)' }}>
                Próxima fatura <ChevronRight size={12} />
              </button>
            </div>
          </div>
        )}

        {installmentEnabled && canToggleInstallments ? (
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div ref={amountRef} className={flashField === 'amount' ? 'rounded-xl animate-field-flash' : ''}>
              <FieldLabel error={errors.amount}>Valor total da compra</FieldLabel>
              <CurrencyInput value={installmentTotal} onChange={setInstallmentTotal} />
              {errors.amount && <p className="text-xs mt-1" style={{ color: 'var(--expense)' }}>{errors.amount}</p>}
            </div>
            <div ref={dateRef} className={flashField === 'date' ? 'rounded-xl animate-field-flash' : ''}>
              <FieldLabel error={errors.date}>Data da compra</FieldLabel>
              <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} className={inputClass} style={inputStyle} />
              {errors.date && <p className="text-xs mt-1" style={{ color: 'var(--expense)' }}>{errors.date}</p>}
            </div>
          </div>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div ref={amountRef} className={flashField === 'amount' ? 'rounded-xl animate-field-flash' : ''}>
              <FieldLabel error={errors.amount}>Valor</FieldLabel>
              {form.isSalary ? (
                <div className={inputClass} style={{ ...inputStyle, display: 'flex', alignItems: 'center', color: 'var(--text-soft)' }}>{formatBRL(form.amount)} <span className="ml-1 text-xs">(líquido calculado)</span></div>
              ) : (
                <CurrencyInput value={form.amount} onChange={(v) => setForm({ ...form, amount: v })} />
              )}
              {errors.amount && <p className="text-xs mt-1" style={{ color: 'var(--expense)' }}>{errors.amount}</p>}
            </div>
            <div ref={dateRef} className={flashField === 'date' ? 'rounded-xl animate-field-flash' : ''}>
              <FieldLabel error={errors.date}>Data</FieldLabel>
              <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} className={inputClass} style={inputStyle} />
              {errors.date && <p className="text-xs mt-1" style={{ color: 'var(--expense)' }}>{errors.date}</p>}
            </div>
          </div>
        )}
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <FieldLabel>Categoria</FieldLabel>
            <Select value={form.category} onChange={(e) => { setForm({ ...form, category: e.target.value }); setCategoryTouched(true); }} className={inputClass} style={inputStyle}>
              {CATEGORY_NAMES.map((c) => <option key={c} value={c}>{c}</option>)}
            </Select>
          </div>
          {form.type === 'receita' ? (
            <div ref={accountRef} className={flashField === 'account' ? 'rounded-xl animate-field-flash' : ''}>
              <FieldLabel error={errors.account}>Conta</FieldLabel>
              <Select value={form.account} onChange={(e) => { setForm({ ...form, account: e.target.value }); setPaymentTouched(true); }} className={inputClass} style={inputStyle}>
                {accounts.map((a) => <option key={a.id} value={a.id}>{a.bank} ({a.type})</option>)}
              </Select>
            </div>
          ) : (
            <div ref={accountRef} className={flashField === 'account' ? 'rounded-xl animate-field-flash' : ''}>
              <FieldLabel error={errors.account}>Como foi pago</FieldLabel>
              <Select
                value={form.cardId ? `card:${form.cardId}` : form.benefitId ? `benefit:${form.benefitId}:${form.benefitType}` : `account:${form.account}`}
                onChange={(e) => selectPaymentSource(e.target.value)}
                className={inputClass} style={inputStyle}
              >
                <optgroup label="Contas (débito)">
                  {accounts.map((a) => <option key={a.id} value={`account:${a.id}`}>{a.bank} ({a.type})</option>)}
                </optgroup>
                {cards.length > 0 && (
                  <optgroup label="Cartões de crédito">
                    {cards.map((c) => <option key={c.id} value={`card:${c.id}`}>{c.bank} — {c.brand}</option>)}
                  </optgroup>
                )}
                {benefits.length > 0 && (
                  <optgroup label="Vale-benefícios">
                    {benefits.map((b) => [
                      <option key={`${b.id}-food`} value={`benefit:${b.id}:foodBalance`}>{b.provider} — Alimentação (saldo {formatBRL(b.foodBalance)})</option>,
                      <option key={`${b.id}-mob`} value={`benefit:${b.id}:mobilityBalance`}>{b.provider} — Transporte (saldo {formatBRL(b.mobilityBalance)})</option>,
                    ])}
                  </optgroup>
                )}
              </Select>
              {sessionPatternNotice && !paymentTouched && <p className="text-xs mt-1.5" style={{ color: 'var(--primary)' }}>{sessionPatternNotice}</p>}
            </div>
          )}
        </div>
        {selectedCard && <p className="text-xs -mt-2" style={{ color: 'var(--text-soft)' }}>Essa despesa entra na fatura do cartão e só sai da conta quando você pagar a fatura.</p>}
        {selectedBenefit && <p className="text-xs -mt-2" style={{ color: 'var(--text-soft)' }}>O valor é descontado automaticamente do saldo desse vale-benefícios ao salvar — sem precisar atualizar na mão.</p>}
        {form.type === 'despesa' && selectedCard && canToggleInstallments && (
          <div className="rounded-xl p-4 space-y-3" style={{ backgroundColor: 'var(--bg)', border: '1px solid var(--border)' }}>
            <label className="flex items-center gap-2.5 text-sm cursor-pointer select-none py-1" style={{ color: 'var(--text)' }}>
              <input type="checkbox" checked={installmentEnabled} onChange={(e) => { setInstallmentEnabled(e.target.checked); if (e.target.checked) setInstallmentTotal(form.amount || 0); }} className="w-5 h-5 shrink-0 focus-ring" />
              {initial ? 'Converter em compra parcelada' : 'Compra parcelada'}
            </label>
            {installmentEnabled && (
              <>
                <div>
                  <FieldLabel>Número de parcelas</FieldLabel>
                  <input
                    type="number" inputMode="numeric" min={2} max={48}
                    value={installmentCountText}
                    onFocus={(e) => e.target.select()}
                    onChange={(e) => {
                      const raw = e.target.value;
                      setInstallmentCountText(raw);
                      const n = Number(raw);
                      if (raw !== '' && !Number.isNaN(n)) setInstallmentCount(n);
                    }}
                    onBlur={() => {
                      const clamped = Math.max(2, Math.min(48, Number(installmentCountText) || 2));
                      setInstallmentCount(clamped);
                      setInstallmentCountText(String(clamped));
                    }}
                    className={inputClass} style={inputStyle}
                  />
                </div>
                <div className="rounded-lg px-3 py-2 text-sm font-medium tabular-nums" style={{ backgroundColor: 'var(--primary-soft)', color: 'var(--primary-dark)' }}>
                  {installmentCount}x de {formatBRL(perInstallment)}
                </div>
                <p className="text-xs" style={{ color: 'var(--text-soft)' }}>
                  {initial
                    ? 'Ao salvar, esse lançamento à vista vira essas parcelas novas, respeitando o fechamento do cartão. Não dá pra desfazer sozinho, então confira o valor total antes de salvar.'
                    : 'Cada parcela vira um lançamento no mês certo, respeitando o fechamento do cartão (dia ' + selectedCard.closingDay + ') — não conta como despesa recorrente. Dá pra antecipar as parcelas que faltam a qualquer momento na página do cartão.'}
                </p>
              </>
            )}
          </div>
        )}
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <FieldLabel>Forma de pagamento</FieldLabel>
            <Select value={form.paymentMethod} onChange={(e) => setForm({ ...form, paymentMethod: e.target.value })} className={inputClass} style={inputStyle}>
              {PAYMENT_METHODS.map((p) => <option key={p} value={p}>{p}</option>)}
            </Select>
          </div>
          <div>
            <FieldLabel>Status</FieldLabel>
            <Select value={form.status} disabled={(installmentEnabled && canToggleInstallments) || !!form.cardId} onChange={(e) => setForm({ ...form, status: e.target.value })} className={inputClass} style={{ ...inputStyle, opacity: (installmentEnabled && canToggleInstallments) || !!form.cardId ? 0.6 : 1 }}>
              {STATUS_OPTIONS.map((s) => <option key={s} value={s}>{statusLabel(s, form.type)}</option>)}
            </Select>
            {form.cardId && <p className="text-xs mt-1" style={{ color: 'var(--text-soft)' }}>Fica Pendente até você pagar a fatura desse cartão, na aba Cartões.</p>}
          </div>
        </div>
        {duplicateWarning && (
          <div className="rounded-xl p-3 text-xs space-y-2" style={{ backgroundColor: 'var(--expense-soft)', color: 'var(--expense)' }}>
            <p>Isso parece parecido com um lançamento que já existe: <strong>{duplicateWarning.description}</strong>, {formatDateShortYear(duplicateWarning.date)}, {formatBRL(duplicateWarning.amount)}. Quer salvar mesmo assim?</p>
            <div className="flex gap-2">
              <button type="button" onClick={() => setDuplicateWarning(null)} className="flex-1 text-xs font-medium px-2.5 py-1.5 rounded-lg" style={{ backgroundColor: 'var(--card)', color: 'var(--text-soft)' }}>Rever lançamento</button>
              <button type="button" onClick={doSave} className="flex-1 text-xs font-medium px-2.5 py-1.5 rounded-lg" style={{ backgroundColor: 'var(--expense)', color: '#fff' }}>Salvar mesmo assim</button>
            </div>
          </div>
        )}
        <div className="flex items-center justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={handleSubmit} className={shakeError ? 'animate-shake' : ''}>{initial ? 'Salvar alterações' : 'Adicionar lançamento'}</Button>
        </div>
      </div>
    </Modal>
  );
}

function ImportReviewModal({ parsed, accounts, cards, onConfirm, onClose }) {
  const nubankCard = cards.find((c) => c.bank.toLowerCase().includes('nubank')) || cards[0];
  const nubankAccount = accounts.find((a) => a.bank.toLowerCase().includes('nubank')) || accounts[0];
  const useCard = cards.length > 0;
  const [targetCard, setTargetCard] = useState(nubankCard?.id || '');
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
    const card = useCard ? cards.find((c) => c.id === targetCard) : null;
    const toImport = rows
      .filter((r) => r.include)
      .map((r) => ({
        id: uid(),
        date: r.date,
        description: r.description,
        category: r.category,
        type: 'despesa',
        account: card ? card.accountId : targetAccount,
        cardId: card ? card.id : null,
        paymentMethod: 'Cartão de crédito',
        amount: r.amount,
        status: card ? 'Pendente' : 'Pago',
        // Antes a parcela virava texto solto tipo "(parcela 3/3)" dentro da descrição — a tag
        // na 2ª linha do card (data · cartão · forma de pagamento · parcela) depende desses três
        // campos estruturados, os mesmos usados numa compra parcelada criada manualmente.
        installmentGroupId: r.installment ? uid() : null,
        installmentIndex: r.installment ? r.installment.current : null,
        installmentCount: r.installment ? r.installment.total : null,
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
          <div className="rounded-xl p-3" style={{ backgroundColor: 'var(--bg)' }}>
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
          <FieldLabel>{useCard ? 'Lançar no cartão' : 'Lançar na conta'}</FieldLabel>
          {useCard ? (
            <Select value={targetCard} onChange={(e) => setTargetCard(e.target.value)} className={inputClass} style={inputStyle}>
              {cards.map((c) => <option key={c.id} value={c.id}>{c.bank} — {c.brand}</option>)}
            </Select>
          ) : (
            <>
              <Select value={targetAccount} onChange={(e) => setTargetAccount(e.target.value)} className={inputClass} style={inputStyle}>
                {accounts.map((a) => <option key={a.id} value={a.id}>{a.bank} ({a.type})</option>)}
              </Select>
              <p className="text-xs mt-1" style={{ color: 'var(--text-soft)' }}>Você ainda não cadastrou um cartão — os lançamentos vão debitar essa conta diretamente. Cadastre o cartão em "Cartões" para a fatura ser calculada automaticamente.</p>
            </>
          )}
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
                      <input type="checkbox" checked={r.include} onChange={() => toggleRow(r.rowId)} className="w-5 h-5 focus-ring cursor-pointer" />
                    </td>
                    <td className="py-2 pr-3 whitespace-nowrap" style={{ color: 'var(--text-soft)' }}>{formatDate(r.date)}</td>
                    <td className="py-2 pr-3">
                      <span style={{ color: 'var(--text)' }}>{r.description}</span>
                      {r.installment && <span className="ml-2 text-xs" style={{ color: 'var(--text-soft)' }}>parcela {r.installment.current}/{r.installment.total}</span>}
                      {r.isDuplicate && <span className="ml-2 text-xs font-medium" style={{ color: 'var(--alert)' }}>{r.isExactMatch ? 'duplicata' : 'possível duplicata (nome parecido)'}</span>}
                    </td>
                    <td className="py-2 pr-3">
                      <Select value={r.category} onChange={(e) => setRowCategory(r.rowId, e.target.value)} className="px-2 py-1.5 rounded-lg text-base sm:text-xs focus-ring" style={inputStyle}>
                        {CATEGORY_NAMES.map((c) => <option key={c} value={c}>{c}</option>)}
                      </Select>
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
          <Button onClick={handleConfirm} disabled={selectedCount === 0 || (useCard ? !targetCard : !targetAccount)}>Importar {selectedCount > 0 ? `(${selectedCount})` : ''}</Button>
        </div>
      </div>
    </Modal>
  );
}

/* ============================================================
   PÁGINA: TRANSAÇÕES / RECEITAS / DESPESAS (componente genérico)
   ============================================================ */

function TransactionsPage({ transactions, accounts, cards, benefits = [], settings, onAdd, onEdit, onDelete, onImport, onMarkPaid, onGoToFatura }) {
  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState('all'); // 'all' | 'despesa' | 'receita'
  const [categoryFilter, setCategoryFilter] = useState('all');
  const [accountFilter, setAccountFilter] = useState('all');
  const [statusFilter, setStatusFilter] = useState('all');
  const [sortConfig, setSortConfig] = useState({ key: 'date', direction: 'desc' });
  const [page, setPage] = useState(1);
  const [editing, setEditing] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [confirmMarkPaid, setConfirmMarkPaid] = useState(null);
  const [confirmDeleteRow, setConfirmDeleteRow] = useState(null);
  const [showFilters, setShowFilters] = useState(false);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const filtersRef = useRef(null);
  const exportMenuRef = useRef(null);
  const activeFilterCount = [typeFilter, categoryFilter, accountFilter, statusFilter].filter((f) => f !== 'all').length;
  const fileInputRef = useRef(null);
  const [importPreview, setImportPreview] = useState(null);
  const pageSize = 8;

  const filtered = useMemo(() => transactions.filter((t) => {
    const matchesType = typeFilter === 'all' || t.type === typeFilter;
    const matchesSearch = t.description.toLowerCase().includes(search.toLowerCase());
    const matchesCategory = categoryFilter === 'all' || t.category === categoryFilter;
    const matchesAccount = accountFilter === 'all' || t.account === accountFilter;
    const matchesStatus = statusFilter === 'all' || t.status === statusFilter;
    return matchesType && matchesSearch && matchesCategory && matchesAccount && matchesStatus;
  }), [transactions, typeFilter, search, categoryFilter, accountFilter, statusFilter]);

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

  useEffect(() => { setPage(1); }, [search, typeFilter, categoryFilter, accountFilter, statusFilter]);

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
    link.href = url; link.download = 'transacoes.csv'; link.click();
    URL.revokeObjectURL(url);
  }
  function exportExcel() {
    const rows = sorted.map((t) => ({ Data: formatDate(t.date), Descrição: t.description, Categoria: t.category, Tipo: t.type, Conta: accounts.find((a) => a.id === t.account)?.bank || 'Conta removida', 'Forma de pagamento': t.paymentMethod, Valor: t.amount, Status: t.status }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Transações');
    XLSX.writeFile(wb, 'transacoes.xlsx');
  }

  const sortableColumns = [
    { key: 'date', label: 'Data' }, { key: 'description', label: 'Descrição' }, { key: 'category', label: 'Categoria', center: true },
    { key: 'amount', label: 'Valor' }, { key: 'status', label: 'Status', center: true },
  ];

  return (
    <div className="space-y-6 print-area">
      {onGoToFatura && isVisible(settings, 'transacoes', 'faturaBanner') && (
        <div className="flex items-center justify-between gap-3 rounded-2xl px-4 py-3 no-print" style={{ backgroundColor: 'var(--primary-soft)' }}>
          <p className="text-sm" style={{ color: 'var(--primary-dark)' }}>
            Esta lista mostra tudo, incluindo parcelas de meses futuros. Pra ver só o que cai neste mês (tipo uma fatura), use o acompanhamento mensal.
          </p>
          <Button size="sm" variant="secondary" icon={CalendarIcon} onClick={onGoToFatura}>Fatura mensal</Button>
        </div>
      )}
      {isVisible(settings, 'transacoes', 'statCards') && (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          {typeFilter !== 'receita' && <StatCard title="Total de despesas" value={formatBRL(totalDespesas)} description="No filtro atual" icon={TrendingDown} color="var(--expense)" soft="var(--expense-soft)" />}
          {typeFilter !== 'despesa' && <StatCard title="Total de receitas" value={formatBRL(totalReceitas)} description="No filtro atual" icon={TrendingUp} color="var(--income)" soft="var(--income-soft)" />}
        </div>
      )}

      <Card>
        <div className="space-y-2 no-print">
          {/* Linha 1: busca + filtros. flex-nowrap de propósito — com flex-wrap, ao aparecer o
              badge de contagem no botão "Filtros" ele deixava de caber ao lado da busca (que
              tinha uma largura mínima de 180px) e quebrava sozinho pra uma 2ª linha, enquanto a
              busca esticava pra ocupar a linha inteira. Com min-w-0 no mobile, a busca encolhe
              o quanto for preciso pra sempre sobrar espaço pro botão de filtros do lado. */}
          <div className="flex flex-nowrap items-center gap-2">
            <div className="relative flex-1 min-w-0 sm:min-w-[180px]">
              <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2" color="var(--text-soft)" />
              <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Pesquisar" className="w-full pl-8 pr-3 py-2 rounded-xl text-base sm:text-sm focus-ring" style={inputStyle} />
            </div>

            {/* Telas maiores: filtros lado a lado. No mobile viram um único botão "Filtros" com popover — 3 selects lado a lado apertava demais a linha de controles. */}
            <div className="hidden sm:flex items-center gap-2">
              <Select value={typeFilter} onChange={(e) => setTypeFilter(e.target.value)} className="px-3 py-2 rounded-xl text-sm focus-ring" style={inputStyle}>
                <option value="all">Despesas e receitas</option>
                <option value="despesa">Só despesas</option>
                <option value="receita">Só receitas</option>
              </Select>
              <Select value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)} className="px-3 py-2 rounded-xl text-sm focus-ring" style={inputStyle}>
                <option value="all">Todas categorias</option>
                {CATEGORY_NAMES.map((c) => <option key={c} value={c}>{c}</option>)}
              </Select>
              <Select value={accountFilter} onChange={(e) => setAccountFilter(e.target.value)} className="px-3 py-2 rounded-xl text-sm focus-ring" style={inputStyle}>
                <option value="all">Todas contas</option>
                {accounts.map((a) => <option key={a.id} value={a.id}>{a.bank}</option>)}
              </Select>
              <Select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} className="px-3 py-2 rounded-xl text-sm focus-ring" style={inputStyle}>
                <option value="all">Todos status</option>
                {STATUS_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
              </Select>
            </div>
            <div className="sm:hidden shrink-0">
              <button
                ref={filtersRef}
                onClick={() => setShowFilters((v) => !v)}
                className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-sm font-medium"
                style={inputStyle}
              >
                <Filter size={14} color="var(--text-soft)" /> Filtros
                {activeFilterCount > 0 && (
                  <span className="w-4 h-4 rounded-full text-[10px] font-semibold flex items-center justify-center" style={{ backgroundColor: 'var(--primary)', color: '#fff' }}>{activeFilterCount}</span>
                )}
              </button>
              <Popover open={showFilters} onClose={() => setShowFilters(false)} triggerRef={filtersRef} width={256} align="center" className="p-3 space-y-2.5">
                <Select value={typeFilter} onChange={(e) => { setTypeFilter(e.target.value); setShowFilters(false); }} className="w-full px-3 py-2.5 rounded-xl text-base focus-ring" style={inputStyle}>
                  <option value="all">Despesas e receitas</option>
                  <option value="despesa">Só despesas</option>
                  <option value="receita">Só receitas</option>
                </Select>
                <Select value={categoryFilter} onChange={(e) => { setCategoryFilter(e.target.value); setShowFilters(false); }} className="w-full px-3 py-2.5 rounded-xl text-base focus-ring" style={inputStyle}>
                  <option value="all">Todas categorias</option>
                  {CATEGORY_NAMES.map((c) => <option key={c} value={c}>{c}</option>)}
                </Select>
                <Select value={accountFilter} onChange={(e) => { setAccountFilter(e.target.value); setShowFilters(false); }} className="w-full px-3 py-2.5 rounded-xl text-base focus-ring" style={inputStyle}>
                  <option value="all">Todas contas</option>
                  {accounts.map((a) => <option key={a.id} value={a.id}>{a.bank}</option>)}
                </Select>
                <Select value={statusFilter} onChange={(e) => { setStatusFilter(e.target.value); setShowFilters(false); }} className="w-full px-3 py-2.5 rounded-xl text-base focus-ring" style={inputStyle}>
                  <option value="all">Todos status</option>
                  {STATUS_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
                </Select>
                {activeFilterCount > 0 && (
                  <button onClick={() => { setTypeFilter('all'); setCategoryFilter('all'); setAccountFilter('all'); setStatusFilter('all'); setShowFilters(false); }} className="text-xs font-medium" style={{ color: 'var(--primary)' }}>Limpar filtros</button>
                )}
              </Popover>
            </div>
          </div>

          {/* Linha 2: importar, exportar/imprimir e novo lançamento — sempre nessa linha,
              nunca estoura pra uma 3ª (tamanhos compactos no mobile garantem isso). */}
          <div className="flex items-center gap-2">
            <input type="file" accept=".csv" ref={fileInputRef} onChange={handleImportFile} className="hidden" />
            <Button variant="secondary" size="toolbar" icon={Upload} onClick={() => fileInputRef.current.click()}>Importar fatura</Button>

            {/* Telas maiores: botões de exportação individuais. No mobile viram um menu "⋮" — 3
                botões de texto a mais nessa linha eram demais pra uma tela estreita. */}
            <div className="hidden sm:flex items-center gap-2">
              <Button variant="secondary" size="toolbar" icon={Download} onClick={exportCSV}>CSV</Button>
              <Button variant="secondary" size="toolbar" icon={Download} onClick={exportExcel}>Excel</Button>
              <Button variant="secondary" size="toolbar" icon={FileText} onClick={() => window.print()}>Imprimir</Button>
            </div>

            <div className="flex-1" />

            <Button size="toolbar" icon={Plus} onClick={() => setShowForm(true)}>Novo</Button>

            <div className="sm:hidden">
              <button ref={exportMenuRef} onClick={() => setShowExportMenu((v) => !v)} className="flex items-center justify-center w-8 h-8 rounded-xl shrink-0" style={inputStyle} title="Exportar ou imprimir">
                <MoreVertical size={14} color="var(--text-soft)" />
              </button>
              <Popover open={showExportMenu} onClose={() => setShowExportMenu(false)} triggerRef={exportMenuRef} width={192} align="center" className="p-1.5">
                <button onClick={() => { exportCSV(); setShowExportMenu(false); }} className="w-full text-left text-sm px-3 py-2.5 rounded-lg hover:bg-black/5 flex items-center gap-2.5" style={{ color: 'var(--text)' }}><Download size={14} color="var(--text-soft)" /> CSV</button>
                <button onClick={() => { exportExcel(); setShowExportMenu(false); }} className="w-full text-left text-sm px-3 py-2.5 rounded-lg hover:bg-black/5 flex items-center gap-2.5" style={{ color: 'var(--text)' }}><Download size={14} color="var(--text-soft)" /> Excel</button>
                <button onClick={() => { window.print(); setShowExportMenu(false); }} className="w-full text-left text-sm px-3 py-2.5 rounded-lg hover:bg-black/5 flex items-center gap-2.5" style={{ color: 'var(--text)' }}><FileText size={14} color="var(--text-soft)" /> Imprimir</button>
              </Popover>
            </div>
          </div>
        </div>
        {activeFilterCount > 0 && (
          <div className="flex sm:hidden items-center flex-wrap gap-1.5 mt-2">
            {categoryFilter !== 'all' && <Badge color="var(--primary)" soft="var(--primary-soft)">{categoryFilter}</Badge>}
            {accountFilter !== 'all' && <Badge color="var(--primary)" soft="var(--primary-soft)">{accounts.find((a) => a.id === accountFilter)?.bank}</Badge>}
            {statusFilter !== 'all' && <Badge color="var(--primary)" soft="var(--primary-soft)">{statusFilter}</Badge>}
          </div>
        )}

        {pageData.length === 0 ? (
          <EmptyState icon={Search} title="Nenhuma transação encontrada" description="Ajuste os filtros ou adicione um novo lançamento." />
        ) : (
          <>
            {/* Mobile: cards com swipe (arraste pra esquerda) revelando editar/excluir — uma
                tabela de 8 colunas não cabe numa tela de celular sem cortar informação. */}
            <div className="sm:hidden mt-4 space-y-2">
              {pageData.map((tx) => {
                const cat = CATEGORIES[tx.category] || CATEGORIES['Outros'];
                const card = tx.cardId ? cards.find((c) => c.id === tx.cardId) : null;
                const accName = accounts.find((a) => a.id === tx.account)?.bank || 'Conta removida';
                const inst = getInstallmentDisplay(tx);
                return (
                  <SwipeableRow
                    key={tx.id} onEdit={() => setEditing(tx)} onDelete={() => onDelete(tx)}
                    deleteConfirm={{ title: 'Excluir lançamento', description: `Tem certeza que deseja excluir "${tx.description}"? Essa ação não pode ser desfeita.` }}
                  >
                    <div className="flex items-center gap-3 p-3">
                      <IconCircle icon={cat.icon} color={cat.color} soft={cat.soft} size={36} />
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center justify-between gap-2">
                          <div className="flex items-center gap-1.5 min-w-0">
                            <p className="text-sm font-medium truncate min-w-0" style={{ color: 'var(--text)' }}>{inst.desc}</p>
                            {inst.count && <span className="shrink-0 text-[10px] font-medium tabular-nums px-1.5 py-0.5 rounded-md" style={{ backgroundColor: 'var(--primary-soft)', color: 'var(--primary-dark)' }}>{inst.index}/{inst.count}</span>}
                          </div>
                          <span className="text-sm font-medium tabular-nums shrink-0" style={{ color: tx.type === 'receita' ? 'var(--income)' : 'var(--expense)' }}>{tx.type === 'receita' ? '+' : '-'} {formatBRL(tx.amount)}</span>
                        </div>
                        <div className="flex items-center gap-1.5 mt-0.5 text-xs" style={{ color: 'var(--text-soft)' }}>
                          <span className="shrink-0">{formatDateShortYear(tx.date)}</span>
                          <span className="shrink-0">·</span>
                          <span className="truncate">{card ? `Crédito · ${card.bank}` : `Débito · ${accName}`}</span>
                        </div>
                        {onMarkPaid && tx.status === 'Pendente' && !tx.cardId && (
                          <button onClick={() => setConfirmMarkPaid(tx)} className="mt-1.5 text-xs font-medium flex items-center gap-1" style={{ color: 'var(--income)' }}>
                            <Check size={12} /> Marcar como pago
                          </button>
                        )}
                      </div>
                    </div>
                  </SwipeableRow>
                );
              })}
              <p className="text-[11px] text-center pt-1" style={{ color: 'var(--text-soft)' }}>Arraste um lançamento pra esquerda para editar ou excluir</p>
            </div>

            {/* Desktop/tablet: tabela completa */}
            <div className="hidden sm:block overflow-x-auto mt-4 -mx-1">
              <table className="w-full min-w-[720px]">
                <thead>
                  <tr className="text-left text-xs" style={{ color: 'var(--text-soft)' }}>
                    {sortableColumns.map((c) => (
                      <th key={c.key} className={`pb-2 pr-3 font-medium cursor-pointer select-none${c.center ? ' text-center' : ''}`} onClick={() => handleSort(c.key)}>
                        <span className={`flex items-center gap-1${c.center ? ' justify-center' : ''}`}>{c.label} {sortConfig.key === c.key && (sortConfig.direction === 'asc' ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}</span>
                      </th>
                    ))}
                    <th className="pb-2 pr-3 font-medium">Conta</th>
                    <th className="pb-2 pr-3 font-medium text-center">Pagamento</th>
                    <th className="pb-2 pr-1 font-medium no-print"></th>
                  </tr>
                </thead>
                <tbody style={{ borderTop: '1px solid var(--border)' }}>
                  {pageData.map((tx) => {
                    const inst = getInstallmentDisplay(tx);
                    return (
                    <tr key={tx.id} className="text-sm hover:bg-black/[0.02]">
                      <td className="py-3 pr-3 whitespace-nowrap" style={{ color: 'var(--text-soft)' }}>{formatDate(tx.date)}</td>
                      <td className="py-3 pr-3" style={{ color: 'var(--text)' }}>
                        {inst.desc}
                        {inst.count && (
                          <span className="ml-2 text-[11px] font-medium tabular-nums px-1.5 py-0.5 rounded-md" style={{ backgroundColor: 'var(--primary-soft)', color: 'var(--primary-dark)' }}>
                            {inst.index}/{inst.count}
                          </span>
                        )}
                      </td>
                      <td className="py-3 pr-3 text-center"><CategoryBadge category={tx.category} /></td>
                      <td className="py-3 pr-3 tabular-nums font-medium whitespace-nowrap" style={{ color: tx.type === 'receita' ? 'var(--income)' : 'var(--expense)' }}>{tx.type === 'receita' ? '+' : '-'} {formatBRL(tx.amount)}</td>
                      <td className="py-3 pr-3 text-center"><StatusBadge status={tx.status} type={tx.type} /></td>
                      <td className="py-3 pr-3 whitespace-nowrap" style={{ color: 'var(--text-soft)' }}>{accounts.find((a) => a.id === tx.account)?.bank || 'Conta removida'}</td>
                      <td className="py-3 pr-3 whitespace-nowrap text-center" style={{ color: 'var(--text-soft)' }}>{tx.paymentMethod}</td>
                      <td className="py-3 pr-1 no-print">
                        <div className="flex items-center gap-1">
                          {onMarkPaid && tx.status === 'Pendente' && !tx.cardId && (
                            <button onClick={() => setConfirmMarkPaid(tx)} title="Marcar como pago" className="p-2 rounded-lg hover:bg-black/5"><Check size={14} color="var(--income)" /></button>
                          )}
                          <button onClick={() => setEditing(tx)} className="p-2 rounded-lg hover:bg-black/5" title="Editar"><Pencil size={14} color="var(--text-soft)" /></button>
                          <button onClick={() => setConfirmDeleteRow(tx)} className="p-2 rounded-lg hover:bg-black/5" title="Excluir"><Trash2 size={14} color="var(--expense)" /></button>
                        </div>
                      </td>
                    </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </>
        )}

        {totalPages > 1 && (
          <div className="flex items-center justify-between mt-4 no-print">
            <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Página {page} de {totalPages} · {sorted.length} lançamentos</p>
            <div className="flex items-center gap-1">
              <button disabled={page === 1} onClick={() => setPage((p) => p - 1)} className="p-2 rounded-lg hover:bg-black/5 disabled:opacity-40" style={{ color: 'var(--text-soft)' }} aria-label="Página anterior" title="Página anterior"><ChevronLeft size={16} /></button>
              <button disabled={page === totalPages} onClick={() => setPage((p) => p + 1)} className="p-2 rounded-lg hover:bg-black/5 disabled:opacity-40" style={{ color: 'var(--text-soft)' }} aria-label="Próxima página" title="Próxima página"><ChevronRight size={16} /></button>
            </div>
          </div>
        )}
      </Card>

      {showForm && <TransactionForm accounts={accounts} cards={cards} benefits={benefits} transactions={transactions} onSave={(f) => { onAdd(f); setShowForm(false); }} onClose={() => setShowForm(false)} />}
      {editing && <TransactionForm initial={editing} accounts={accounts} cards={cards} benefits={benefits} transactions={transactions} onSave={(f) => { onEdit(f); setEditing(null); }} onClose={() => setEditing(null)} onDelete={(tx) => { onDelete(tx); setEditing(null); }} />}
      {importPreview && <ImportReviewModal parsed={importPreview} accounts={accounts} cards={cards} onConfirm={handleConfirmImport} onClose={() => setImportPreview(null)} />}
      {confirmMarkPaid && (
        <ConfirmModal
          title="Marcar como pago"
          description={`Isso desconta ${formatBRL(confirmMarkPaid.amount)} do saldo de ${accounts.find((a) => a.id === confirmMarkPaid.account)?.bank || 'sua conta'} agora. Confirma?`}
          variant="primary" confirmLabel="Marcar como pago"
          onConfirm={() => { onMarkPaid(confirmMarkPaid); setConfirmMarkPaid(null); }}
          onClose={() => setConfirmMarkPaid(null)}
        />
      )}
      {confirmDeleteRow && (
        <ConfirmModal
          title="Excluir lançamento"
          description={`Tem certeza que deseja excluir "${confirmDeleteRow.description}"? Essa ação não pode ser desfeita.`}
          onConfirm={() => { onDelete(confirmDeleteRow); setConfirmDeleteRow(null); }}
          onClose={() => setConfirmDeleteRow(null)}
        />
      )}
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
          <Select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })} className={inputClass} style={inputStyle}>
            <option>Conta Corrente</option>
            <option>Poupança</option>
          </Select>
        </div>
        <div>
          <FieldLabel>Saldo inicial</FieldLabel>
          <SignedCurrencyInput value={form.balance} onChange={(v) => setForm({ ...form, balance: v })} />
          <p className="text-xs mt-1" style={{ color: 'var(--text-soft)' }}>A conta pode ficar negativa (cheque especial, conta estourada etc) — use o botão + / − pra alternar o sinal.</p>
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

/* ---------- Padrão reutilizável: "atualizar valor" com histórico e justificativa ---------- */
/* Usado em caixinhas e vale-benefícios: um campo único atualiza o total. Se o novo valor for
   menor que o atual, pede uma observação curta antes de confirmar (dinheiro foi usado/retirado). */

function UpdateValueControl({ currentValue, onUpdate, label = 'Atualizar valor' }) {
  const [editing, setEditing] = useState(false);
  const [value, setValue] = useState(currentValue);
  const [confirming, setConfirming] = useState(false);
  const [note, setNote] = useState('');

  function handleSubmit() {
    if (value === currentValue) { setEditing(false); return; }
    if (value < currentValue) setConfirming(true);
    else { onUpdate(value, null); setEditing(false); }
  }
  function handleConfirmDecrease() {
    onUpdate(value, note.trim() || null);
    setConfirming(false); setEditing(false); setNote('');
  }
  function cancel() { setConfirming(false); setEditing(false); setNote(''); }

  if (!editing) {
    return (
      <button onClick={() => { setValue(currentValue); setEditing(true); }} className="text-xs font-medium flex items-center gap-1" style={{ color: 'var(--primary)' }}>
        <Pencil size={11} /> {label}
      </button>
    );
  }
  if (confirming) {
    return (
      <div className="space-y-2 mt-1.5">
        <p className="text-xs flex items-start gap-1.5" style={{ color: 'var(--expense)' }}>
          <AlertCircle size={13} className="mt-0.5 shrink-0" /> O novo valor é menor que o atual — o que aconteceu?
        </p>
        <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="Ex: usei parte pra um imprevisto" className={inputClass} style={inputStyle} />
        <div className="flex items-center gap-2">
          <Button size="sm" onClick={handleConfirmDecrease}>Confirmar</Button>
          <button onClick={cancel} className="p-2" title="Cancelar"><X size={14} color="var(--text-soft)" /></button>
        </div>
      </div>
    );
  }
  return (
    <div className="flex items-center gap-2 mt-1.5">
      <CurrencyInput value={value} onChange={setValue} />
      <Button size="sm" onClick={handleSubmit}>OK</Button>
      <button onClick={cancel} className="p-2" title="Cancelar"><X size={14} color="var(--text-soft)" /></button>
    </div>
  );
}

function BalanceHistoryLog({ history }) {
  const [openNote, setOpenNote] = useState(null);
  if (!history || history.length === 0) return null;
  const items = [...history].reverse();
  return (
    <div className="space-y-1 mt-2 max-h-32 overflow-y-auto pr-1">
      {items.map((h, i) => (
        <div key={i} className="text-xs px-2 py-1.5 rounded-lg" style={{ backgroundColor: 'var(--bg)' }}>
          <div className="flex items-center justify-between gap-2">
            <span style={{ color: 'var(--text-soft)' }}>{formatMonthYear(h.date)}</span>
            <div className="flex items-center gap-1.5">
              <span className="tabular-nums font-medium" style={{ color: 'var(--text)' }}>{formatBRL(h.value)}</span>
              {h.note && (
                <button onClick={() => setOpenNote(openNote === i ? null : i)} className="p-1 rounded hover:bg-black/5 shrink-0">
                  <MessageCircle size={12} color="var(--primary)" />
                </button>
              )}
            </div>
          </div>
          {openNote === i && h.note && <p className="mt-1 italic" style={{ color: 'var(--text-soft)' }}>"{h.note}"</p>}
        </div>
      ))}
    </div>
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
          <Select value={form.accountId} onChange={(e) => setForm({ ...form, accountId: e.target.value })} className={inputClass} style={inputStyle}>
            {accounts.map((a) => <option key={a.id} value={a.id}>{a.bank} ({a.type})</option>)}
          </Select>
        </div>
        <div className="flex justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={() => { if (!form.name.trim()) { setError('Informe um nome'); return; } onSave(form); }}>Criar caixinha</Button>
        </div>
      </div>
    </Modal>
  );
}

function CaixinhaCard({ caixinha, accountName, onUpdateValue, onDelete }) {
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const history = caixinha.history || [];
  return (
    <Card padding="p-5">
      <div className="flex items-start gap-3 mb-3">
        <IconCircle icon={PiggyBank} color="var(--goals)" soft="var(--goals-soft)" size={36} />
        <div className="min-w-0 flex-1">
          <p className="text-sm font-medium truncate" style={{ color: 'var(--text)' }}>{caixinha.name}</p>
          <p className="text-xs" style={{ color: 'var(--text-soft)' }}>{accountName}</p>
        </div>
        <button onClick={() => setConfirmDelete(true)} className="p-2 rounded-lg hover:bg-black/5" title="Excluir"><Trash2 size={14} color="var(--text-soft)" /></button>
      </div>
      <p className="font-display text-lg font-bold tabular-nums mb-2" style={{ color: 'var(--text)' }}>{formatBRL(caixinha.balance)}</p>
      <UpdateValueControl currentValue={caixinha.balance} onUpdate={(value, note) => onUpdateValue(caixinha, value, note)} />
      {history.length > 0 && (
        <>
          <button onClick={() => setShowHistory((s) => !s)} className="text-xs font-medium mt-2.5 flex items-center gap-1" style={{ color: 'var(--text-soft)' }}>
            {showHistory ? <ChevronUp size={12} /> : <ChevronDown size={12} />} Histórico ({history.length})
          </button>
          {showHistory && <BalanceHistoryLog history={history} />}
        </>
      )}
      {confirmDelete && (
        <ConfirmModal
          title="Excluir caixinha"
          description={`Tem certeza que deseja excluir a caixinha "${caixinha.name}"? O saldo guardado nela (${formatBRL(caixinha.balance)}) volta para a conta "${accountName}".`}
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
  const isNegative = acc.balance < 0;
  const atRisk = (acc.alertThreshold > 0 && acc.balance <= acc.alertThreshold) || isNegative;
  return (
    <Card className="animate-fade-up" padding="p-5" style={atRisk ? { border: '1px solid var(--expense)' } : undefined}>
      <div className="flex items-center gap-3 mb-3">
        <IconCircle icon={Icon} color="var(--invest)" soft="var(--invest-soft)" size={36} />
        <div className="min-w-0 flex-1">
          <p className="text-sm font-medium truncate" style={{ color: 'var(--text)' }}>{acc.bank}</p>
          <p className="text-xs" style={{ color: 'var(--text-soft)' }}>{acc.type}</p>
        </div>
        <button onClick={() => setConfirmDelete(true)} className="p-2 rounded-lg hover:bg-black/5 shrink-0" title="Excluir"><Trash2 size={14} color="var(--text-soft)" /></button>
      </div>
      {!editingBalance ? (
        <div className="flex items-center gap-1.5 mt-3 min-w-0">
          <p className="font-display text-xl font-bold tabular-nums truncate" style={{ color: atRisk ? 'var(--expense)' : 'var(--text)' }}>{formatBRL(acc.balance)}</p>
          <button onClick={() => { setBalanceValue(acc.balance); setEditingBalance(true); }} className="p-2 rounded-lg hover:bg-black/5 shrink-0" title="Editar saldo"><Pencil size={12} color="var(--text-soft)" /></button>
        </div>
      ) : (
        <div className="flex items-center gap-2 mt-3 flex-wrap">
          <SignedCurrencyInput value={balanceValue} onChange={setBalanceValue} />
          <Button size="sm" onClick={() => { onSetBalance(acc, balanceValue); setEditingBalance(false); }}>OK</Button>
          <button onClick={() => setEditingBalance(false)} className="p-2" title="Cancelar"><X size={14} color="var(--text-soft)" /></button>
        </div>
      )}
      {isNegative ? (
        <p className="text-xs flex items-start gap-1 mt-1" style={{ color: 'var(--expense)' }}>
          <AlertCircle size={12} className="mt-0.5 shrink-0" /> Conta negativa
        </p>
      ) : atRisk && (
        <p className="text-xs flex items-start gap-1 mt-1" style={{ color: 'var(--expense)' }}>
          <AlertCircle size={12} className="mt-0.5 shrink-0" /> Saldo em zona de risco (abaixo de {formatBRL(acc.alertThreshold)})
        </p>
      )}
      {!editingThreshold ? (
        <button onClick={() => setEditingThreshold(true)} className="text-xs font-medium mt-3 flex items-start gap-1 text-left" style={{ color: 'var(--primary)' }}>
          <Bell size={12} className="mt-0.5 shrink-0" /> <span>{acc.alertThreshold > 0 ? `Alerta em ${formatBRL(acc.alertThreshold)}` : 'Definir alerta de saldo baixo'}</span>
        </button>
      ) : (
        <div className="flex items-center gap-2 mt-3 flex-wrap">
          <CurrencyInput value={thresholdValue} onChange={setThresholdValue} />
          <Button size="sm" onClick={() => { onSetThreshold(acc, thresholdValue); setEditingThreshold(false); }}>OK</Button>
          <button onClick={() => setEditingThreshold(false)} className="p-2" title="Cancelar"><X size={14} color="var(--text-soft)" /></button>
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

function AccountsPage({ accounts, caixinhas, transactions, settings, onAddAccount, onDeleteAccount, onSetAccountThreshold, onSetAccountBalance, onAddCaixinha, onDeleteCaixinha, onUpdateCaixinhaValue }) {
  const [showAccountForm, setShowAccountForm] = useState(false);
  const [showCaixinhaForm, setShowCaixinhaForm] = useState(false);
  const total = accounts.reduce((s, a) => s + a.balance, 0);
  return (
    <div className="space-y-6">
      {isVisible(settings, 'contas', 'saldoTotalCard') && (
        <Card>
          <div className="flex items-center justify-between mb-1">
            <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Saldo total em contas</p>
          </div>
          <p className="font-display text-3xl font-bold tabular-nums" style={{ color: 'var(--text)' }}>{formatBRL(total)}</p>
        </Card>
      )}

      <div>
        <SectionTitle action={<Button size="sm" icon={Plus} onClick={() => setShowAccountForm(true)}>Nova conta</Button>}>Suas contas</SectionTitle>
        <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">
          {accounts.map((acc) => {
            const usageCount = transactions.filter((t) => t.account === acc.id).length + caixinhas.filter((c) => c.accountId === acc.id).length;
            return <AccountCard key={acc.id} acc={acc} onDelete={onDeleteAccount} onSetThreshold={onSetAccountThreshold} onSetBalance={onSetAccountBalance} usageCount={usageCount} />;
          })}
        </div>
      </div>

      {isVisible(settings, 'contas', 'caixinhas') && (
        <div>
          <SectionTitle action={<Button size="sm" icon={Plus} onClick={() => setShowCaixinhaForm(true)}>Nova caixinha</Button>}>
            Caixinhas <span className="text-xs font-normal" style={{ color: 'var(--text-soft)' }}>reservas flexíveis dentro das suas contas</span>
          </SectionTitle>
          {caixinhas.length === 0 ? (
            <Card><EmptyState icon={PiggyBank} title="Nenhuma caixinha criada" description="Crie caixinhas para separar dinheiro para objetivos do dia a dia." /></Card>
          ) : (
            <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">
              {caixinhas.map((cx) => (
                <CaixinhaCard key={cx.id} caixinha={cx} accountName={accounts.find((a) => a.id === cx.accountId)?.bank || ''} onUpdateValue={onUpdateCaixinhaValue} onDelete={onDeleteCaixinha} />
              ))}
            </div>
          )}
        </div>
      )}

      {showAccountForm && <AccountForm onSave={(f) => { onAddAccount(f); setShowAccountForm(false); }} onClose={() => setShowAccountForm(false)} />}
      {showCaixinhaForm && <CaixinhaForm accounts={accounts} onSave={(f) => { onAddCaixinha(f); setShowCaixinhaForm(false); }} onClose={() => setShowCaixinhaForm(false)} />}
    </div>
  );
}

/* ============================================================
   PÁGINA: CARTÕES
   ============================================================ */

function CardForm({ initial, accounts, pendingInvoiceCount = 0, onSave, onDelete, onClose }) {
  const [form, setForm] = useState(initial || { bank: '', brand: '', accountId: accounts[0]?.id || '', limit: 0, closingDay: 1, dueDay: 10, paidThroughDate: null });
  const [errors, setErrors] = useState({});
  const [confirmDelete, setConfirmDelete] = useState(false);
  const clampDay = (v) => Math.min(31, Math.max(1, Math.round(v) || 1));
  function validate() {
    const e = {};
    if (!form.bank.trim()) e.bank = 'Informe o banco';
    if (!form.limit || form.limit <= 0) e.limit = 'Informe um limite maior que zero';
    if (!form.accountId) e.accountId = 'Selecione a conta que paga a fatura';
    setErrors(e);
    return Object.keys(e).length === 0;
  }
  return (
    <Modal title={initial ? 'Editar cartão' : 'Novo cartão'} onClose={onClose}>
      <div className="space-y-4">
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <FieldLabel error={errors.bank}>Banco</FieldLabel>
            <input value={form.bank} onChange={(e) => setForm({ ...form, bank: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Nubank" />
          </div>
          <div>
            <FieldLabel>Bandeira</FieldLabel>
            <input value={form.brand} onChange={(e) => setForm({ ...form, brand: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Mastercard" />
          </div>
        </div>
        <div>
          <FieldLabel error={errors.accountId}>Conta que paga a fatura</FieldLabel>
          <Select value={form.accountId} onChange={(e) => setForm({ ...form, accountId: e.target.value })} className={inputClass} style={inputStyle}>
            {accounts.length === 0 && <option value="">Nenhuma conta cadastrada</option>}
            {accounts.map((a) => <option key={a.id} value={a.id}>{a.bank} ({a.type})</option>)}
          </Select>
        </div>
        <div>
          <FieldLabel error={errors.limit}>Limite total</FieldLabel>
          <CurrencyInput value={form.limit} onChange={(v) => setForm({ ...form, limit: v })} />
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div><FieldLabel>Dia do fechamento</FieldLabel><input type="number" inputMode="numeric" min={1} max={31} value={form.closingDay} onChange={(e) => setForm({ ...form, closingDay: Number(e.target.value) })} onBlur={(e) => setForm({ ...form, closingDay: clampDay(Number(e.target.value)) })} className={inputClass} style={inputStyle} /></div>
          <div><FieldLabel>Dia do vencimento</FieldLabel><input type="number" inputMode="numeric" min={1} max={31} value={form.dueDay} onChange={(e) => setForm({ ...form, dueDay: Number(e.target.value) })} onBlur={(e) => setForm({ ...form, dueDay: clampDay(Number(e.target.value)) })} className={inputClass} style={inputStyle} /></div>
        </div>
        <p className="text-xs" style={{ color: 'var(--text-soft)' }}>A fatura é calculada automaticamente a partir dos lançamentos feitos neste cartão — não precisa informar um valor inicial.</p>
        <div className="flex items-center justify-between gap-3 pt-2">
          {initial && onDelete ? (
            <button onClick={() => setConfirmDelete(true)} className="text-xs font-medium flex items-center gap-1 shrink-0" style={{ color: 'var(--expense)' }}>
              <Trash2 size={13} /> Excluir cartão
            </button>
          ) : <span />}
          <div className="flex flex-wrap justify-end gap-3">
            <Button variant="secondary" onClick={onClose}>Cancelar</Button>
            <Button onClick={() => { if (!validate()) return; onSave({ ...form, closingDay: clampDay(form.closingDay), dueDay: clampDay(form.dueDay) }); }}>{initial ? 'Salvar' : 'Adicionar cartão'}</Button>
          </div>
        </div>
      </div>
      {confirmDelete && (pendingInvoiceCount > 0 ? (
        <ConfirmModal
          title="Não é possível excluir"
          description={`Este cartão ainda tem ${pendingInvoiceCount} lançamento(s) com fatura em aberto (pendente). Pague a fatura ou edite/exclua esses lançamentos antes de excluir o cartão.`}
          confirmLabel="Entendi" variant="primary"
          onConfirm={() => setConfirmDelete(false)}
          onClose={() => setConfirmDelete(false)}
        />
      ) : (
        <ConfirmModal
          title="Excluir cartão"
          description={`Tem certeza que deseja excluir o cartão "${form.bank}"? Essa ação não pode ser desfeita.`}
          onConfirm={() => { onDelete(initial); onClose(); }}
          onClose={() => setConfirmDelete(false)}
        />
      ))}
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
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
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

function BenefitBalance({ label, icon: Icon, value, history, onUpdate }) {
  const [showHistory, setShowHistory] = useState(false);
  return (
    <div className="rounded-xl p-3" style={{ backgroundColor: 'var(--bg)' }}>
      <div className="flex items-center gap-1.5 mb-1 text-xs" style={{ color: 'var(--text-soft)' }}><Icon size={13} /> {label}</div>
      <p className="font-display text-base font-bold tabular-nums" style={{ color: 'var(--text)' }}>{formatBRL(value)}</p>
      <UpdateValueControl currentValue={value} onUpdate={onUpdate} label="Atualizar saldo" />
      {history && history.length > 0 && (
        <>
          <button onClick={() => setShowHistory((s) => !s)} className="text-[11px] font-medium mt-2 flex items-center gap-1" style={{ color: 'var(--text-soft)' }}>
            {showHistory ? <ChevronUp size={11} /> : <ChevronDown size={11} />} Histórico ({history.length})
          </button>
          {showHistory && <BalanceHistoryLog history={history} />}
        </>
      )}
    </div>
  );
}

function BenefitCard({ benefit, onUpdate, onDelete }) {
  return (
    <Card padding="p-5">
      <div className="flex items-start gap-3 mb-3">
        <IconCircle icon={CreditCard} color="var(--goals)" soft="var(--goals-soft)" size={36} />
        <div className="min-w-0 flex-1">
          <p className="text-sm font-medium truncate" style={{ color: 'var(--text)' }}>{benefit.provider}</p>
          <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Vale-benefícios</p>
        </div>
        <button onClick={() => onDelete(benefit)} className="p-2 rounded-lg hover:bg-black/5" title="Excluir"><Trash2 size={14} color="var(--text-soft)" /></button>
      </div>
      <div className="grid grid-cols-1 gap-3">
        <BenefitBalance label="Alimentação" icon={Utensils} value={benefit.foodBalance} history={benefit.foodHistory} onUpdate={(value, note) => onUpdate(benefit, 'foodBalance', value, note)} />
        <BenefitBalance label="Transporte" icon={Car} value={benefit.mobilityBalance} history={benefit.mobilityHistory} onUpdate={(value, note) => onUpdate(benefit, 'mobilityBalance', value, note)} />
      </div>
    </Card>
  );
}

/* ---------- Fatura mensal (acompanhamento por mês, separando crédito de débito) ---------- */

function MonthNavigator({ label, monthOffset, onPrev, onNext, onToday }) {
  return (
    <div className="flex flex-col items-center gap-1">
      <div className="flex items-center gap-1">
        <button onClick={onPrev} className="p-2 rounded-lg hover:bg-black/5" style={{ color: 'var(--text-soft)' }} title="Mês anterior"><ChevronLeft size={16} /></button>
        <p className="text-sm font-medium min-w-[150px] text-center" style={{ color: 'var(--text)' }}>{label}</p>
        <button onClick={onNext} className="p-2 rounded-lg hover:bg-black/5" style={{ color: 'var(--text-soft)' }} title="Próximo mês"><ChevronRight size={16} /></button>
      </div>
      {monthOffset !== 0 && (
        <button onClick={onToday} className="text-xs font-medium" style={{ color: 'var(--primary)' }}>hoje</button>
      )}
    </div>
  );
}

function MonthlyInvoicePage({ cards, transactions, accounts, benefits = [], cardGradients, onPayInvoice, onAdvanceInstallments, onMarkPaid, onEditTransaction, onDeleteTransaction }) {
  const [monthOffset, setMonthOffset] = useState(0);
  const [cardFilter, setCardFilter] = useState('all'); // 'all' | <cardId> — selecionado clicando na linha do cartão
  const [search, setSearch] = useState('');
  const [subview, setSubview] = useState('fatura'); // 'fatura' | 'todas'
  const [confirmMarkPaid, setConfirmMarkPaid] = useState(null);
  const [confirmDeleteRow, setConfirmDeleteRow] = useState(null);
  const [editingTx, setEditingTx] = useState(null);

  const refDate = useMemo(() => {
    const d = new Date();
    d.setDate(1);
    d.setMonth(d.getMonth() + monthOffset);
    return d;
  }, [monthOffset]);
  const year = refDate.getFullYear();
  const month = refDate.getMonth();
  const monthLabel = refDate.toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
  // Clicar numa linha de cartão seleciona só ele (e some com as outras linhas da lista); clicar
  // de novo tira a seleção. Não faria sentido "clicar fora" pra desmarcar, então o próprio card
  // selecionado é o único jeito de reverter.
  const toggleCard = (cardId) => setCardFilter((cur) => (cur === cardId ? 'all' : cardId));
  const visibleCards = cardFilter === 'all' ? cards : cards.filter((c) => c.id === cardFilter);

  // "Fatura" mostra só crédito, porque é a única que faz sentido marcar como paga (a compra em
  // débito já saiu da conta na hora ou é agendada, não existe "fatura" pra ela). Agrupa por
  // CICLO da fatura (respeitando o dia de fechamento de cada cartão), não pelo mês-calendário
  // bruto da compra — por isso cada cartão é avaliado com seu próprio getCardInvoiceCycle.
  const faturaTx = useMemo(() => transactions
    .filter((t) => {
      if (t.type !== 'despesa' || !t.cardId) return false;
      if (cardFilter !== 'all' && cardFilter !== t.cardId) return false;
      const card = cards.find((c) => c.id === t.cardId);
      if (!card) return false;
      const cycle = getCardInvoiceCycle(card, t.date);
      if (cycle.year !== year || cycle.month !== month) return false;
      if (search.trim()) {
        const q = search.trim().toLowerCase();
        if (!t.description.toLowerCase().includes(q) && !t.category.toLowerCase().includes(q)) return false;
      }
      return true;
    })
    .sort((a, b) => b.date.localeCompare(a.date)), [transactions, cardFilter, year, month, cards, search]);

  // "Todas as despesas do mês" mistura débito e crédito de propósito, só pra visualizar o mês
  // inteiro — sem botão de pagar fatura, porque débito não tem esse conceito.
  const todasTx = useMemo(() => transactions
    .filter((t) => t.type === 'despesa' && isSameMonth(t.date, year, month)
      && (cardFilter === 'all' || t.cardId === cardFilter)
      && (!search.trim() || t.description.toLowerCase().includes(search.trim().toLowerCase()) || t.category.toLowerCase().includes(search.trim().toLowerCase())))
    .sort((a, b) => b.date.localeCompare(a.date)), [transactions, cardFilter, year, month, search]);

  const list = subview === 'fatura' ? faturaTx : todasTx;
  const total = list.reduce((s, t) => s + t.amount, 0);
  // Uma única tag resumindo a fatura inteira, em vez de repetir Pago/Pendente em cada linha (que
  // normalmente é tudo igual, já que o pagamento acontece de uma vez pra fatura toda).
  const faturaStatus = useMemo(() => {
    if (subview !== 'fatura' || faturaTx.length === 0) return null;
    const paidCount = faturaTx.filter((t) => t.status === 'Pago').length;
    if (paidCount === faturaTx.length) return { label: 'Fatura paga', color: 'var(--income)', soft: 'var(--income-soft)' };
    if (paidCount === 0) return { label: 'Fatura pendente', color: 'var(--alert)', soft: 'var(--alert-soft)' };
    return { label: 'Fatura paga parcialmente', color: 'var(--alert)', soft: 'var(--alert-soft)' };
  }, [subview, faturaTx]);

  return (
    <div className="space-y-6">
      <div className="flex justify-center">
        <MonthNavigator label={capitalizeFirst(monthLabel)} monthOffset={monthOffset} onPrev={() => setMonthOffset((m) => m - 1)} onNext={() => setMonthOffset((m) => m + 1)} onToday={() => setMonthOffset(0)} />
      </div>

      {cards.length > 0 && (
        <div className="space-y-2 sm:max-w-xl">
          {visibleCards.map((c) => (
            <CardInvoiceRow
              key={c.id} card={c} transactions={transactions} accounts={accounts}
              selected={cardFilter === c.id} onToggleSelect={() => toggleCard(c.id)}
              year={year} month={month} subview={subview}
              onPayInvoice={onPayInvoice}
              gradient={cardGradients[cards.indexOf(c) % cardGradients.length]}
            />
          ))}
        </div>
      )}

      <div className="relative sm:max-w-xs">
        <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2" color="var(--text-soft)" />
        <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Pesquisar" className="w-full pl-8 pr-3 py-2 rounded-xl text-base sm:text-sm focus-ring" style={inputStyle} />
      </div>

      <Card>
        <div className="flex flex-wrap items-center gap-1 mb-4 p-1 rounded-xl w-fit max-w-full" style={{ backgroundColor: 'var(--bg)' }}>
          <button
            onClick={() => setSubview('fatura')}
            className="text-xs font-medium px-3 py-1.5 rounded-lg transition-colors focus-ring"
            style={subview === 'fatura' ? { backgroundColor: 'var(--surface)', color: 'var(--text)', boxShadow: '0 1px 2px rgba(0,0,0,0.08)' } : { color: 'var(--text-soft)' }}
          >
            Fatura (crédito)
          </button>
          <button
            onClick={() => setSubview('todas')}
            className="text-xs font-medium px-3 py-1.5 rounded-lg transition-colors focus-ring"
            style={subview === 'todas' ? { backgroundColor: 'var(--surface)', color: 'var(--text)', boxShadow: '0 1px 2px rgba(0,0,0,0.08)' } : { color: 'var(--text-soft)' }}
          >
            Todas as despesas do mês
          </button>
        </div>
        <div className="flex items-center justify-between mb-3 gap-3 flex-wrap">
          <SectionTitle subtitle={subview === 'fatura' ? 'Só lançamentos no crédito deste mês' : 'Débito + crédito deste mês'}>
            {list.length} lançamento(s)
          </SectionTitle>
          <div className="flex items-center gap-2">
            {faturaStatus && <Badge color={faturaStatus.color} soft={faturaStatus.soft}>{faturaStatus.label}</Badge>}
            <span className="text-sm font-semibold tabular-nums" style={{ color: 'var(--text)' }}>{formatBRL(total)}</span>
          </div>
        </div>
        {list.length === 0 ? (
          <EmptyState icon={CalendarIcon} title="Nada por aqui" description="Nenhum lançamento encontrado para este mês com esse filtro." />
        ) : (
          <>
            {/* Mobile: mesmo card com swipe (arraste pra esquerda) usado em Transações, revelando
                editar/excluir por trás — antes essa aba só tinha a linha larga de desktop. */}
            <div className="sm:hidden space-y-2">
              {list.map((t) => {
                const cat = CATEGORIES[t.category] || CATEGORIES['Outros'];
                const card = t.cardId ? cards.find((c) => c.id === t.cardId) : null;
                const acc = !card ? accounts.find((a) => a.id === t.account) : null;
                const inst = getInstallmentDisplay(t);
                return (
                  <SwipeableRow
                    key={t.id} onEdit={onEditTransaction ? () => setEditingTx(t) : undefined} onDelete={onDeleteTransaction ? () => onDeleteTransaction(t) : undefined}
                    deleteConfirm={{ title: 'Excluir lançamento', description: `Tem certeza que deseja excluir "${t.description}"? Essa ação não pode ser desfeita.` }}
                  >
                    <div className="flex items-center gap-3 p-3">
                      <IconCircle icon={cat.icon} color={cat.color} soft={cat.soft} size={36} />
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center justify-between gap-2">
                          <div className="flex items-center gap-1.5 min-w-0">
                            <p className="text-sm font-medium truncate min-w-0" style={{ color: 'var(--text)' }}>{inst.desc}</p>
                            {inst.count && <span className="shrink-0 text-[10px] font-medium tabular-nums px-1.5 py-0.5 rounded-md" style={{ backgroundColor: 'var(--primary-soft)', color: 'var(--primary-dark)' }}>{inst.index}/{inst.count}</span>}
                          </div>
                          <span className="text-sm font-medium tabular-nums shrink-0" style={{ color: 'var(--expense)' }}>{formatBRL(t.amount)}</span>
                        </div>
                        <div className="flex items-center gap-1.5 mt-0.5 text-xs">
                          <span className="truncate" style={{ color: 'var(--text-soft)' }}>
                            {formatDateShortYear(t.date)} · {card ? card.bank : (acc ? acc.bank : 'Conta removida')} · {card ? 'Crédito' : 'Débito'}
                          </span>
                          {subview !== 'fatura' && <span className="shrink-0"><StatusBadge status={t.status} type={t.type} /></span>}
                        </div>
                        {onMarkPaid && t.status === 'Pendente' && !t.cardId && (
                          <button onClick={() => setConfirmMarkPaid(t)} className="mt-1.5 text-xs font-medium flex items-center gap-1" style={{ color: 'var(--income)' }}>
                            <Check size={12} /> Marcar como pago
                          </button>
                        )}
                      </div>
                    </div>
                  </SwipeableRow>
                );
              })}
              <p className="text-[11px] text-center pt-1" style={{ color: 'var(--text-soft)' }}>Arraste um lançamento pra esquerda para editar ou excluir</p>
            </div>

            {/* Desktop/tablet: linha única, como antes. */}
            <div className="hidden sm:block space-y-1">
              {list.map((t) => {
                const cat = CATEGORIES[t.category] || CATEGORIES['Outros'];
                const card = t.cardId ? cards.find((c) => c.id === t.cardId) : null;
                const acc = !card ? accounts.find((a) => a.id === t.account) : null;
                const inst = getInstallmentDisplay(t);
                return (
                  <div key={t.id} className="flex items-center gap-3 py-2.5 px-1 hover:bg-black/[0.02] rounded-lg">
                    <IconCircle icon={cat.icon} color={cat.color} soft={cat.soft} size={36} />
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-1.5 min-w-0">
                        <p className="text-sm font-medium truncate min-w-0" style={{ color: 'var(--text)' }}>{inst.desc}</p>
                        {inst.count && <span className="shrink-0 text-[10px] font-medium tabular-nums px-1.5 py-0.5 rounded-md" style={{ backgroundColor: 'var(--primary-soft)', color: 'var(--primary-dark)' }}>{inst.index}/{inst.count}</span>}
                      </div>
                      <p className="text-xs truncate" style={{ color: 'var(--text-soft)' }}>
                        {formatDateShortYear(t.date)} · {card ? card.bank : (acc ? acc.bank : 'Conta removida')} · {card ? 'Crédito' : 'Débito'}
                      </p>
                    </div>
                    {subview !== 'fatura' && <StatusBadge status={t.status} type={t.type} />}
                    {onMarkPaid && t.status === 'Pendente' && !t.cardId && (
                      <button onClick={() => setConfirmMarkPaid(t)} title="Marcar como pago" className="p-2 rounded-lg hover:bg-black/5"><Check size={14} color="var(--income)" /></button>
                    )}
                    <span className="text-sm font-medium tabular-nums shrink-0" style={{ color: 'var(--expense)' }}>{formatBRL(t.amount)}</span>
                    {onEditTransaction && (
                      <button onClick={() => setEditingTx(t)} className="p-2 rounded-lg hover:bg-black/5 shrink-0" title="Editar lançamento"><Pencil size={14} color="var(--text-soft)" /></button>
                    )}
                    {onDeleteTransaction && (
                      <button onClick={() => setConfirmDeleteRow(t)} className="p-2 rounded-lg hover:bg-black/5 shrink-0" title="Excluir lançamento"><Trash2 size={14} color="var(--expense)" /></button>
                    )}
                  </div>
                );
              })}
            </div>
          </>
        )}
      </Card>
      {confirmMarkPaid && (
        <ConfirmModal
          title="Marcar como pago"
          description={`Isso desconta ${formatBRL(confirmMarkPaid.amount)} do saldo da conta agora. Confirma?`}
          variant="primary" confirmLabel="Marcar como pago"
          onConfirm={() => { onMarkPaid(confirmMarkPaid); setConfirmMarkPaid(null); }}
          onClose={() => setConfirmMarkPaid(null)}
        />
      )}
      {confirmDeleteRow && (
        <ConfirmModal
          title="Excluir lançamento"
          description={`Tem certeza que deseja excluir "${confirmDeleteRow.description}"? Essa ação não pode ser desfeita.`}
          onConfirm={() => { onDeleteTransaction(confirmDeleteRow); setConfirmDeleteRow(null); }}
          onClose={() => setConfirmDeleteRow(null)}
        />
      )}
      {editingTx && (
        <TransactionForm
          initial={editingTx} accounts={accounts} cards={cards} benefits={benefits} transactions={transactions}
          onSave={(f) => { onEditTransaction(f); setEditingTx(null); }}
          onClose={() => setEditingTx(null)}
          onDelete={(tx) => { onDeleteTransaction(tx); setEditingTx(null); }}
        />
      )}
    </div>
  );
}

function CardsPage({ cards, transactions, accounts, recurring, settings, cardGradients, onAdd, onEdit, onDelete, onPayInvoice, onAdvanceInstallments, benefits, onAddBenefit, onDeleteBenefit, onUpdateBenefit, view = 'cartoes', onChangeView, onMarkPaid, onEditTransaction, onDeleteTransaction }) {
  const [showForm, setShowForm] = useState(false);
  const [showBenefitForm, setShowBenefitForm] = useState(false);
  const [confirmDeleteBenefit, setConfirmDeleteBenefit] = useState(null);
  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-center gap-1 p-1 rounded-xl w-fit max-w-full" style={{ backgroundColor: 'var(--bg)' }}>
        <button
          onClick={() => onChangeView && onChangeView('cartoes')}
          className="text-xs font-medium px-3 py-1.5 rounded-lg transition-colors focus-ring"
          style={view === 'cartoes' ? { backgroundColor: 'var(--surface)', color: 'var(--text)', boxShadow: '0 1px 2px rgba(0,0,0,0.08)' } : { color: 'var(--text-soft)' }}
        >
          Meus cartões
        </button>
        <button
          onClick={() => onChangeView && onChangeView('fatura')}
          className="text-xs font-medium px-3 py-1.5 rounded-lg transition-colors focus-ring"
          style={view === 'fatura' ? { backgroundColor: 'var(--surface)', color: 'var(--text)', boxShadow: '0 1px 2px rgba(0,0,0,0.08)' } : { color: 'var(--text-soft)' }}
        >
          Fatura mensal
        </button>
      </div>

      {view === 'fatura' ? (
        <MonthlyInvoicePage cards={cards} transactions={transactions} accounts={accounts} benefits={benefits} cardGradients={cardGradients} onPayInvoice={onPayInvoice} onAdvanceInstallments={onAdvanceInstallments} onMarkPaid={onMarkPaid} onEditTransaction={onEditTransaction} onDeleteTransaction={onDeleteTransaction} />
      ) : (
        <>
          <SectionTitle action={<Button size="sm" icon={Plus} onClick={() => setShowForm(true)}>Novo cartão</Button>}>Seus cartões</SectionTitle>
          <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-5">
            {cards.map((c, i) => (
              <CreditCardVisual key={c.id} card={c} transactions={transactions} accounts={accounts} onPayInvoice={onPayInvoice} onAdvanceInstallments={onAdvanceInstallments} onEdit={onEdit} onDelete={onDelete} gradient={cardGradients[i % cardGradients.length]} />
            ))}
          </div>

          {isVisible(settings, 'cartoes', 'beneficios') && (
            <>
              <SectionTitle subtitle="Alimentação e transporte" action={<Button size="sm" icon={Plus} onClick={() => setShowBenefitForm(true)}>Novo vale-benefícios</Button>}>
                Vale-benefícios
              </SectionTitle>
              {benefits.length === 0 ? (
                <Card><EmptyState icon={Utensils} title="Nenhum vale-benefícios cadastrado" description="Adicione seu cartão Flash (ou similar) para acompanhar alimentação e transporte separadamente." /></Card>
              ) : (
                <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">
                  {benefits.map((b) => (
                    <BenefitCard key={b.id} benefit={b} onUpdate={onUpdateBenefit} onDelete={() => setConfirmDeleteBenefit(b)} />
                  ))}
                </div>
              )}
            </>
          )}

          {showForm && <CardForm accounts={accounts} onSave={(f) => { onAdd(f); setShowForm(false); }} onClose={() => setShowForm(false)} />}
          {showBenefitForm && <BenefitForm onSave={(f) => { onAddBenefit(f); setShowBenefitForm(false); }} onClose={() => setShowBenefitForm(false)} />}
          {confirmDeleteBenefit && (
            <ConfirmModal
              title="Excluir vale-benefícios"
              description={`Tem certeza que deseja excluir "${confirmDeleteBenefit.provider}"? Essa ação não pode ser desfeita.`}
              onConfirm={() => { onDeleteBenefit(confirmDeleteBenefit); setConfirmDeleteBenefit(null); }}
              onClose={() => setConfirmDeleteBenefit(null)}
            />
          )}
        </>
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
  const [currentTouched, setCurrentTouched] = useState(!!initial);
  return (
    <Modal title={initial ? 'Editar investimento' : 'Novo investimento'} onClose={onClose}>
      <div className="space-y-4">
        <div>
          <FieldLabel error={error}>Nome</FieldLabel>
          <input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Tesouro Selic" />
        </div>
        <div>
          <FieldLabel>Categoria</FieldLabel>
          <Select value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} className={inputClass} style={inputStyle}>
            {INVESTMENT_CATEGORY_NAMES.map((c) => <option key={c} value={c}>{c}</option>)}
          </Select>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <FieldLabel>Valor investido</FieldLabel>
            <CurrencyInput value={form.invested} onChange={(v) => setForm({ ...form, invested: v, currentValue: currentTouched ? form.currentValue : v })} />
          </div>
          <div>
            <FieldLabel>Valor atual</FieldLabel>
            <CurrencyInput value={form.currentValue} onChange={(v) => { setCurrentTouched(true); setForm({ ...form, currentValue: v }); }} />
          </div>
        </div>
        <p className="text-xs" style={{ color: 'var(--text-soft)' }}>
          {currentTouched ? 'A rentabilidade é calculada automaticamente a partir da diferença entre esses dois valores.' : 'Valor atual acompanha o investido até você ajustá-lo — informe o valor de mercado de hoje se for diferente.'}
        </p>
        <div className="flex justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={() => { if (!form.name.trim()) { setError('Informe o nome'); return; } onSave(form); }}>{initial ? 'Salvar' : 'Adicionar'}</Button>
        </div>
      </div>
    </Modal>
  );
}

function InvestmentsPage({ investments, settings, onAdd, onEdit, onDelete }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [activeAllocation, setActiveAllocation] = useState(null);

  const totalInvested = investments.reduce((s, i) => s + i.invested, 0);
  const totalCurrent = investments.reduce((s, i) => s + i.currentValue, 0);
  const gain = totalCurrent - totalInvested;
  const gainPercent = totalInvested > 0 ? (gain / totalInvested) * 100 : 0;

  const byCategory = {};
  investments.forEach((i) => { byCategory[i.category] = (byCategory[i.category] || 0) + i.currentValue; });
  const allocation = Object.entries(byCategory).map(([name, value]) => ({ name, value, color: INVESTMENT_CATEGORIES[name] || '#A8A398' }));

  const kpiCards = [
    { key: 'kpiInvestido', title: 'Valor investido', value: formatBRL(totalInvested), description: 'Total aportado', icon: Wallet, color: 'var(--primary)', soft: 'var(--primary-soft)' },
    { key: 'kpiAtual', title: 'Valor atual', value: formatBRL(totalCurrent), description: 'Valor de mercado hoje', icon: PiggyBank, color: 'var(--invest)', soft: 'var(--invest-soft)' },
    { key: 'kpiRentabilidade', title: 'Rentabilidade', value: `${gain >= 0 ? '+' : ''}${gainPercent.toFixed(1)}%`, description: formatBRL(gain), icon: gain >= 0 ? TrendingUp : TrendingDown, color: gain >= 0 ? 'var(--income)' : 'var(--expense)', soft: gain >= 0 ? 'var(--income-soft)' : 'var(--expense-soft)' },
  ].filter((c) => isVisible(settings, 'investimentos', c.key));

  return (
    <div className="space-y-6">
      {kpiCards.length > 0 && (
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
          {kpiCards.map((c) => <StatCard key={c.key} {...c} />)}
        </div>
      )}

      {investments.length > 0 && isVisible(settings, 'investimentos', 'allocationChart') && (
        <Card className="max-w-xs">
          <SectionTitle>Distribuição por categoria</SectionTitle>
          <div style={{ width: '100%', height: 180 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie
                  data={allocation} dataKey="value" nameKey="name" innerRadius={48} outerRadius={72} paddingAngle={2}
                  onClick={(_, i) => setActiveAllocation((cur) => (cur === i ? null : i))}
                >
                  {allocation.map((a, i) => (
                    <Cell key={i} fill={a.color} stroke="none" opacity={activeAllocation == null || activeAllocation === i ? 1 : 0.35} style={{ cursor: 'pointer', transition: 'opacity 0.15s ease' }} />
                  ))}
                </Pie>
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div className="space-y-1 mt-2">
            {allocation.map((a, i) => (
              <div
                key={i}
                onClick={() => setActiveAllocation((cur) => (cur === i ? null : i))}
                className="flex items-center gap-2 text-sm rounded-lg px-2 py-1.5 cursor-pointer transition-colors"
                style={{ backgroundColor: activeAllocation === i ? a.color + '26' : 'transparent' }}
              >
                <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: a.color }} />
                <span className="flex-1 truncate" style={{ color: 'var(--text)' }}>{a.name}</span>
                <span className="tabular-nums" style={{ color: 'var(--text-soft)' }}>{totalCurrent > 0 ? ((a.value / totalCurrent) * 100).toFixed(0) : 0}%</span>
                <span className="tabular-nums w-20 text-right shrink-0" style={{ color: 'var(--text)' }}>{formatBRL(a.value)}</span>
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
          <>
            {/* Mobile: mesmo card com swipe (arraste pra esquerda) usado em Transações e
                Recorrentes, revelando editar/excluir por trás — antes só existia a tabela larga,
                que obrigava rolar na horizontal num celular. */}
            <div className="sm:hidden space-y-2">
              {investments.map((inv) => {
                const g = inv.currentValue - inv.invested;
                const gp = inv.invested > 0 ? (g / inv.invested) * 100 : 0;
                const color = INVESTMENT_CATEGORIES[inv.category] || '#A8A398';
                return (
                  <SwipeableRow
                    key={inv.id} onEdit={() => setEditing(inv)} onDelete={() => setConfirmDelete(inv)}
                    deleteConfirm={{ title: 'Excluir investimento', description: `Tem certeza que deseja excluir "${inv.name}"? Essa ação não pode ser desfeita.` }}
                  >
                    <div className="flex items-center gap-3 p-3">
                      <div className="rounded-xl flex items-center justify-center shrink-0" style={{ width: 36, height: 36, backgroundColor: color + '26' }}>
                        <span className="w-2.5 h-2.5 rounded-full" style={{ backgroundColor: color }} />
                      </div>
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center justify-between gap-2">
                          <p className="text-sm font-medium truncate" style={{ color: 'var(--text)' }}>{inv.name}</p>
                          <span className="text-sm font-medium tabular-nums shrink-0" style={{ color: g >= 0 ? 'var(--income)' : 'var(--expense)' }}>{g >= 0 ? '+' : ''}{gp.toFixed(1)}%</span>
                        </div>
                        <p className="text-xs truncate" style={{ color: 'var(--text-soft)' }}>Investido {formatBRL(inv.invested)} · Atual {formatBRL(inv.currentValue)}</p>
                      </div>
                    </div>
                  </SwipeableRow>
                );
              })}
            </div>

            {/* Desktop/tablet: tabela, como antes. */}
            <div className="hidden sm:block overflow-x-auto">
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
                          <button onClick={() => setEditing(inv)} className="p-2 rounded-lg hover:bg-black/5" title="Editar"><Pencil size={14} color="var(--text-soft)" /></button>
                          <button onClick={() => setConfirmDelete(inv)} className="p-2 rounded-lg hover:bg-black/5" title="Excluir"><Trash2 size={14} color="var(--expense)" /></button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          </>
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

function GoalsPage({ goals, onAdd, onEdit, onAddFunds, onDelete, onCompleted }) {
  const [showForm, setShowForm] = useState(false);
  const [editingGoal, setEditingGoal] = useState(null);
  return (
    <div className="space-y-6">
      <SectionTitle action={<Button size="sm" icon={Plus} onClick={() => setShowForm(true)}>Nova meta</Button>}>Suas metas financeiras</SectionTitle>
      {goals.length === 0 ? (
        <Card><EmptyState icon={Target} title="Nenhuma meta criada" description="Crie metas para acompanhar seus objetivos financeiros." /></Card>
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">
          {goals.map((g) => <GoalCard key={g.id} goal={g} onAddFunds={onAddFunds} onEdit={setEditingGoal} onDelete={onDelete} onCompleted={onCompleted} />)}
        </div>
      )}
      {showForm && <GoalForm onSave={(f) => { onAdd(f); setShowForm(false); }} onClose={() => setShowForm(false)} />}
      {editingGoal && <GoalForm initial={editingGoal} onSave={(f) => { onEdit(f); setEditingGoal(null); }} onClose={() => setEditingGoal(null)} />}
    </div>
  );
}

/* ============================================================
   PÁGINA: DESPESAS RECORRENTES
   ============================================================ */

function RecurringForm({ accounts = [], cards = [], initial, onSave, onClose, onLaunchNow, onDelete }) {
  const [form, setForm] = useState(() => initial
    ? { ...initial }
    : { name: '', category: 'Lazer', value: 0, renewalDay: 1, cardId: null, accountId: accounts[0]?.id || '' });
  const [error, setError] = useState('');
  const isEditing = !!initial;

  function selectPaymentSource(v) {
    if (v.startsWith('card:')) setForm({ ...form, cardId: v.slice(5) });
    else setForm({ ...form, cardId: null, accountId: v.slice(8) });
  }

  return (
    <Modal title={isEditing ? 'Editar despesa recorrente' : 'Nova despesa recorrente'} onClose={onClose}>
      <div className="space-y-4">
        <div>
          <FieldLabel error={error}>Nome</FieldLabel>
          <input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} className={inputClass} style={inputStyle} placeholder="Ex: Academia" />
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <FieldLabel>Categoria</FieldLabel>
            <Select value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} className={inputClass} style={inputStyle}>
              {CATEGORY_NAMES.map((c) => <option key={c} value={c}>{c}</option>)}
            </Select>
          </div>
          <div>
            <FieldLabel>Dia da renovação</FieldLabel>
            <input type="number" inputMode="numeric" min={1} max={31} value={form.renewalDay} onChange={(e) => setForm({ ...form, renewalDay: Number(e.target.value) })} className={inputClass} style={inputStyle} />
          </div>
        </div>
        <div>
          <FieldLabel>Valor mensal</FieldLabel>
          <CurrencyInput value={form.value} onChange={(v) => setForm({ ...form, value: v })} />
        </div>
        <div>
          <FieldLabel>Débito ou crédito</FieldLabel>
          <Select value={form.cardId ? `card:${form.cardId}` : `account:${form.accountId}`} onChange={(e) => selectPaymentSource(e.target.value)} className={inputClass} style={inputStyle}>
            {accounts.length > 0 && (
              <optgroup label="Débito (sai direto da conta)">
                {accounts.map((a) => <option key={a.id} value={`account:${a.id}`}>{a.bank} ({a.type})</option>)}
              </optgroup>
            )}
            {cards.length > 0 && (
              <optgroup label="Cartões de crédito">
                {cards.map((c) => <option key={c.id} value={`card:${c.id}`}>{c.bank} — {c.brand}</option>)}
              </optgroup>
            )}
          </Select>
          <p className="text-xs mt-1" style={{ color: 'var(--text-soft)' }}>
            {form.cardId ? 'Entra automaticamente nas próximas faturas desse cartão.' : 'Entra automaticamente nos próximos meses como um lançamento pendente na conta.'}
          </p>
        </div>
        {isEditing && onLaunchNow && (
          <div className="rounded-xl p-3 flex items-center justify-between gap-3 flex-wrap" style={{ backgroundColor: 'var(--bg)', border: '1px solid var(--border)' }}>
            <div className="min-w-0">
              <p className="text-xs font-medium" style={{ color: 'var(--text)' }}>Lançamento manual de hoje</p>
              <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Útil se algum mês ficou de fora da geração automática.</p>
            </div>
            <button type="button" onClick={() => onLaunchNow(form)} className="text-xs font-medium shrink-0 px-2.5 py-1.5 rounded-lg" style={{ color: 'var(--primary)', backgroundColor: 'var(--primary-soft)' }}>
              Lançar agora
            </button>
          </div>
        )}
        <div className="flex items-center justify-end gap-3 pt-2">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={() => { if (!form.name.trim()) { setError('Informe um nome'); return; } onSave(form); }}>{isEditing ? 'Salvar' : 'Adicionar'}</Button>
        </div>
      </div>
    </Modal>
  );
}

function RecurringExpensesPage({ recurring, accounts, cards, settings, onAdd, onEdit, onDelete, onLaunchNow }) {
  const [showForm, setShowForm] = useState(false);
  const [editingItem, setEditingItem] = useState(null);
  const [confirmDeleteItem, setConfirmDeleteItem] = useState(null);
  const [activeBarIndex, setActiveBarIndex] = useState(null);
  const total = recurring.reduce((s, r) => s + r.value, 0);
  const byCategory = {};
  recurring.forEach((r) => { byCategory[r.category] = (byCategory[r.category] || 0) + r.value; });
  const chartData = Object.entries(byCategory)
    .map(([name, value]) => ({ name, value, color: (CATEGORIES[name] || CATEGORIES['Outros']).color }))
    .sort((a, b) => b.value - a.value);
  const chartTotal = chartData.reduce((s, c) => s + c.value, 0);
  const showChart = isVisible(settings, 'recorrentes', 'categoryChart');

  const kpiCards = [
    { key: 'kpiTotalMensal', title: 'Total recorrente mensal', value: formatBRL(total), description: `${recurring.length} assinaturas e despesas fixas`, icon: RefreshCw, color: 'var(--primary)', soft: 'var(--primary-soft)' },
    { key: 'kpiTotalAnual', title: 'Total anual estimado', value: formatBRL(total * 12), description: 'Projeção com base no valor atual', icon: CalendarIcon, color: 'var(--invest)', soft: 'var(--invest-soft)' },
  ].filter((c) => isVisible(settings, 'recorrentes', c.key));

  return (
    <div className="space-y-6">
      {kpiCards.length > 0 && (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          {kpiCards.map((c) => <StatCard key={c.key} {...c} />)}
        </div>
      )}
      <p className="text-xs -mt-2" style={{ color: 'var(--text-soft)' }}>
        Cada despesa recorrente já entra sozinha como lançamento pendente nos próximos meses — não precisa lançar na mão todo mês. Toque numa despesa pra editar ou lançar manualmente um mês específico.
      </p>

      <div className="grid grid-cols-1 lg:grid-cols-5 gap-6">
        {showChart && (
          <Card className="lg:col-span-2">
            <SectionTitle subtitle="onde você mais gasta de forma fixa">Por categoria</SectionTitle>
            {/* Sem rótulos no eixo — com muitas categorias, texto espremido do lado das barras
                fica ilegível (cortado). A leitura fica por conta da legenda com cores abaixo,
                igual ao gráfico de pizza. */}
            <div style={{ width: '100%', height: Math.max(96, chartData.length * 38) }}>
              <ResponsiveContainer>
                <BarChart data={chartData} layout="vertical" margin={{ left: 4, right: 4, top: 4, bottom: 4 }} barCategoryGap={14} onMouseLeave={() => setActiveBarIndex(null)}>
                  <XAxis type="number" hide />
                  <YAxis type="category" dataKey="name" hide />
                  <Bar dataKey="value" radius={[0, 6, 6, 0]} barSize={22}
                    onClick={(_, i) => setActiveBarIndex((cur) => (cur === i ? null : i))}
                  >
                    {chartData.map((entry, i) => (
                      <Cell
                        key={i} fill={entry.color}
                        opacity={activeBarIndex == null || activeBarIndex === i ? 1 : 0.35}
                        style={{ cursor: 'pointer', transition: 'opacity 0.15s ease' }}
                      />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div className="space-y-1 mt-3 pt-3" style={{ borderTop: '1px solid var(--border)' }}>
              {chartData.map((entry, i) => (
                <button
                  key={i}
                  onClick={() => setActiveBarIndex((cur) => (cur === i ? null : i))}
                  className="w-full flex items-center gap-2 text-xs rounded-lg px-1.5 py-1.5 transition-colors"
                  style={{ backgroundColor: activeBarIndex === i ? entry.color + '1f' : 'transparent' }}
                >
                  <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: entry.color }} />
                  <span className="flex-1 truncate text-left" style={{ color: 'var(--text)' }}>{entry.name}</span>
                  <span className="tabular-nums shrink-0" style={{ color: 'var(--text-soft)' }}>{chartTotal > 0 ? ((entry.value / chartTotal) * 100).toFixed(0) : 0}%</span>
                  <span className="tabular-nums font-medium shrink-0 w-[76px] text-right" style={{ color: 'var(--text)' }}>{formatBRL(entry.value)}</span>
                </button>
              ))}
            </div>
          </Card>
        )}
        <Card className={showChart ? 'lg:col-span-3' : 'lg:col-span-5'}>
          <SectionTitle action={<Button size="sm" icon={Plus} onClick={() => setShowForm(true)}>Nova despesa</Button>}>Assinaturas e despesas fixas</SectionTitle>
          {recurring.length === 0 ? (
            <EmptyState icon={RefreshCw} title="Nenhuma despesa recorrente" description="Adicione assinaturas e contas fixas para acompanhar o total mensal." />
          ) : (
            <>
              {/* Mobile: mesmo card com swipe (arraste pra esquerda) usado em Transações,
                  revelando editar/excluir por trás. */}
              <div className="sm:hidden space-y-2">
                {recurring.map((r) => {
                  const cat = CATEGORIES[r.category] || CATEGORIES['Outros'];
                  const card = r.cardId ? cards.find((c) => c.id === r.cardId) : null;
                  const account = r.accountId ? accounts.find((a) => a.id === r.accountId) : null;
                  const sourceLabel = card ? `Crédito · ${card.bank}` : account ? `Débito · ${account.bank}` : null;
                  return (
                    <SwipeableRow
                      key={r.id} onEdit={() => setEditingItem(r)} onDelete={() => onDelete(r)}
                      deleteConfirm={{ title: 'Excluir despesa recorrente', description: `Tem certeza que deseja excluir "${r.name}" da lista? Os lançamentos futuros dela também serão removidos — os que já venceram continuam no seu histórico. Essa ação não pode ser desfeita.` }}
                    >
                      <div className="flex items-start gap-3 p-3">
                        <IconCircle icon={cat.icon} color={cat.color} soft={cat.soft} size={36} />
                        <div className="flex-1 min-w-0">
                          <div className="flex items-start justify-between gap-2">
                            <p className="text-sm font-medium break-words min-w-0" style={{ color: 'var(--text)' }}>{r.name}</p>
                            <span className="text-sm tabular-nums font-medium shrink-0" style={{ color: 'var(--expense)' }}>- {formatBRL(r.value)}</span>
                          </div>
                          <div className="flex items-center gap-x-2 gap-y-0.5 mt-0.5 flex-wrap">
                            <span className="text-xs" style={{ color: 'var(--text-soft)' }}>Renova dia {r.renewalDay}</span>
                            {sourceLabel && <span className="text-xs" style={{ color: 'var(--text-soft)' }}>· {sourceLabel}</span>}
                          </div>
                        </div>
                      </div>
                    </SwipeableRow>
                  );
                })}
                <p className="text-[11px] text-center pt-1" style={{ color: 'var(--text-soft)' }}>Arraste uma despesa pra esquerda para editar ou excluir</p>
              </div>

              {/* Desktop/tablet: linha única com ícones de ação. */}
              <div className="hidden sm:block">
                {recurring.map((r, i) => {
                  const cat = CATEGORIES[r.category] || CATEGORIES['Outros'];
                  const card = r.cardId ? cards.find((c) => c.id === r.cardId) : null;
                  const account = r.accountId ? accounts.find((a) => a.id === r.accountId) : null;
                  const sourceLabel = card ? `Crédito · ${card.bank}` : account ? `Débito · ${account.bank}` : null;
                  return (
                    <div
                      key={r.id}
                      className="flex items-center gap-3 py-3 px-1 hover:bg-black/[0.02] rounded-lg"
                      style={i > 0 ? { borderTop: '1px solid var(--border)' } : undefined}
                    >
                      <IconCircle icon={cat.icon} color={cat.color} soft={cat.soft} size={36} />
                      <div className="flex-1 min-w-0">
                        <p className="text-sm font-medium break-words min-w-0" style={{ color: 'var(--text)' }}>{r.name}</p>
                        <div className="flex items-center gap-x-2 gap-y-0.5 mt-0.5 flex-wrap">
                          <span className="text-xs" style={{ color: 'var(--text-soft)' }}>Renova dia {r.renewalDay}</span>
                          {sourceLabel && <span className="text-xs" style={{ color: 'var(--text-soft)' }}>· {sourceLabel}</span>}
                        </div>
                      </div>
                      <span className="text-sm tabular-nums font-medium shrink-0" style={{ color: 'var(--expense)' }}>- {formatBRL(r.value)}</span>
                      <button onClick={() => setEditingItem(r)} className="p-2 rounded-lg hover:bg-black/5 shrink-0" title="Editar"><Pencil size={14} color="var(--text-soft)" /></button>
                      <button onClick={() => setConfirmDeleteItem(r)} className="p-2 rounded-lg hover:bg-black/5 shrink-0" title="Excluir"><Trash2 size={14} color="var(--expense)" /></button>
                    </div>
                  );
                })}
              </div>
            </>
          )}
        </Card>
      </div>
      {showForm && <RecurringForm accounts={accounts} cards={cards} onSave={(f) => { onAdd(f); setShowForm(false); }} onClose={() => setShowForm(false)} />}
      {editingItem && (
        <RecurringForm
          accounts={accounts} cards={cards} initial={editingItem}
          onSave={(f) => { onEdit(f); setEditingItem(null); }}
          onClose={() => setEditingItem(null)}
          onLaunchNow={(f) => { onLaunchNow(f); setEditingItem(null); }}
        />
      )}
      {confirmDeleteItem && (
        <ConfirmModal
          title="Excluir despesa recorrente"
          description={`Tem certeza que deseja excluir "${confirmDeleteItem.name}" da lista? Os lançamentos futuros dela também serão removidos — os que já venceram continuam no seu histórico. Essa ação não pode ser desfeita.`}
          onConfirm={() => { onDelete(confirmDeleteItem); setConfirmDeleteItem(null); }}
          onClose={() => setConfirmDeleteItem(null)}
        />
      )}
    </div>
  );
}

/* ============================================================
   PÁGINA: RELATÓRIOS
   ============================================================ */

function ReportsPage({ monthlyHistory, categoryComparison, transactions, settings }) {
  // A data de atualização é fixada na montagem da página (não muda a cada re-render enquanto o
  // usuário navega dentro dela) — os números em si já são sempre calculados na hora, isso é só
  // uma referência de "quando eu abri esse relatório".
  const updatedAt = useMemo(() => new Date(), []);
  function exportSummary() {
    const rows = monthlyHistory.map((m) => ({ Mês: m.month, Receitas: m.receitas, Despesas: m.despesas, Saldo: m.receitas - m.despesas, Patrimônio: m.patrimonio }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Resumo mensal');
    XLSX.writeFile(wb, 'relatorio-financeiro.xlsx');
  }
  const showComparison = isVisible(settings, 'relatorios', 'categoryComparison');
  const lastMonth = monthlyHistory[monthlyHistory.length - 1];
  const savingsRate = lastMonth && lastMonth.receitas > 0 ? ((lastMonth.receitas - lastMonth.despesas) / lastMonth.receitas) * 100 : null;
  const topExpenses = useMemo(() => {
    const now = new Date();
    const startStr = ymd(new Date(now.getFullYear(), now.getMonth(), 1));
    const endStr = ymd(new Date(now.getFullYear(), now.getMonth() + 1, 0));
    return transactions
      .filter((t) => t.type === 'despesa' && isRealized(t) && t.date >= startStr && t.date <= endStr)
      .sort((a, b) => b.amount - a.amount)
      .slice(0, 5);
  }, [transactions]);
  return (
    <div className="space-y-6 print-area">
      <div className="flex items-center justify-between gap-3 flex-wrap no-print">
        <p className="text-xs" style={{ color: 'var(--text-soft)' }}>
          Atualizado em {updatedAt.toLocaleDateString('pt-BR')} às {updatedAt.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}
        </p>
        <div className="flex gap-2">
          <Button variant="secondary" size="sm" icon={Download} onClick={exportSummary}>Exportar Excel</Button>
          <Button variant="secondary" size="sm" icon={FileText} onClick={() => window.print()}>Imprimir / PDF</Button>
        </div>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-6">
        <Card>
          <SectionTitle>Taxa de economia</SectionTitle>
          {savingsRate != null ? (
            <>
              <p className="text-2xl font-display font-semibold" style={{ color: savingsRate >= 0 ? 'var(--income)' : 'var(--expense)' }}>{savingsRate.toFixed(0)}%</p>
              <p className="text-xs mt-1" style={{ color: 'var(--text-soft)' }}>Do que entrou este mês, ficou guardado. Receitas − despesas ÷ receitas.</p>
            </>
          ) : (
            <p className="text-sm" style={{ color: 'var(--text-soft)' }}>Sem receitas registradas este mês pra calcular.</p>
          )}
        </Card>
        <Card>
          <SectionTitle>Maiores despesas do mês</SectionTitle>
          {topExpenses.length > 0 ? (
            <div className="space-y-2.5">
              {topExpenses.map((t) => {
                const cat = CATEGORIES[t.category] || CATEGORIES['Outros'];
                return (
                  <div key={t.id} className="flex items-center justify-between gap-2 text-sm">
                    <span className="truncate min-w-0 flex-1" style={{ color: 'var(--text)' }}>{t.description || cat.label}</span>
                    <span className="tabular-nums font-medium shrink-0" style={{ color: 'var(--expense)' }}>{formatBRL(t.amount)}</span>
                  </div>
                );
              })}
            </div>
          ) : (
            <p className="text-sm" style={{ color: 'var(--text-soft)' }}>Nenhuma despesa registrada este mês.</p>
          )}
        </Card>
      </div>

      {showComparison ? (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <CategoryDonut data={categoryComparison.current} title="Gastos por categoria" subtitle="mês atual" />
          <CategoryDonut data={categoryComparison.previous} title="Gastos por categoria" subtitle="mês anterior" />
        </div>
      ) : (
        <Card>
          <EmptyState icon={BarChart3} title="Nada configurado pra mostrar" description='Ative o comparativo por categoria em Configurações → Personalização da interface → Relatórios.' />
        </Card>
      )}
    </div>
  );
}

/* ============================================================
   PÁGINA: CONFIGURAÇÕES
   ============================================================ */

function ToggleSwitch({ checked, onChange, label }) {
  return (
    <button onClick={() => onChange(!checked)} role="switch" aria-checked={checked} aria-label={label} title={label} className="w-11 h-6 rounded-full p-0.5 transition-colors focus-ring shrink-0" style={{ backgroundColor: checked ? 'var(--primary)' : 'var(--border)' }}>
      <div className="w-5 h-5 rounded-full bg-white shadow-sm transition-transform" style={{ transform: checked ? 'translateX(20px)' : 'translateX(0)' }} />
    </button>
  );
}

function VisibilitySettingsSection({ settings, onChangeSettings }) {
  const [activeTab, setActiveTab] = useState(VISIBILITY_SCHEMA[0]?.key);
  function toggleItem(pageKey, itemKey, value) {
    const visibility = { ...(settings.visibility || {}) };
    visibility[pageKey] = { ...(visibility[pageKey] || {}), [itemKey]: value };
    onChangeSettings({ ...settings, visibility });
  }
  const activePage = VISIBILITY_SCHEMA.find((p) => p.key === activeTab) || VISIBILITY_SCHEMA[0];
  return (
    <Card>
      <SectionTitle subtitle='Escolha quais cards, gráficos e blocos aparecem em cada página. Navegação, listas principais e formulários são sempre exibidos.'>
        Personalização da interface
      </SectionTitle>
      {/* Abas em vez de accordions empilhados — com 7 páginas, uma lista vertical de accordions
          significava rolar bastante pra alternar de categoria. Uma barra de abas (com quebra de
          linha) deixa qualquer categoria a um toque, sem rolagem extra. */}
      <div className="flex flex-wrap gap-1.5 mb-4">
        {VISIBILITY_SCHEMA.map((page) => (
          <button
            key={page.key}
            onClick={() => setActiveTab(page.key)}
            className="text-xs font-medium px-3 py-2 rounded-lg transition-colors focus-ring"
            style={activeTab === page.key ? { backgroundColor: 'var(--primary)', color: '#fff' } : { backgroundColor: 'var(--bg)', color: 'var(--text-soft)' }}
          >
            {page.label}
          </button>
        ))}
      </div>
      {activePage && (
        <div className="space-y-3.5">
          <div className="flex items-center justify-between gap-3 pb-3" style={{ borderBottom: '1px solid var(--border)' }}>
            <p className="text-sm min-w-0 flex-1" style={{ color: 'var(--text-soft)' }}>{activePage.essentialLabel}</p>
            <Badge color="var(--text-soft)" soft="var(--bg)">Obrigatório</Badge>
          </div>
          {activePage.items.map((item) => (
            <div key={item.key} className="flex items-center justify-between gap-3">
              <p className="text-sm min-w-0 flex-1" style={{ color: 'var(--text)' }}>{item.label}</p>
              <ToggleSwitch checked={isVisible(settings, activePage.key, item.key)} onChange={(v) => toggleItem(activePage.key, item.key, v)} label={item.label} />
            </div>
          ))}
        </div>
      )}
    </Card>
  );
}

function SettingsPage({ settings, onChangeSettings, onReset, onClearData, dropboxConnected, dropboxBusy, dropboxLastBackup, dropboxSyncError, onConnectDropbox, onDisconnectDropbox, onBackupNow, onRestoreFromDropbox, onExportBackup, onImportBackup }) {
  const [showConfirmReset, setShowConfirmReset] = useState(false);
  const [showConfirmClear, setShowConfirmClear] = useState(false);
  const [showConfirmRestore, setShowConfirmRestore] = useState(false);
  const [pendingImportFile, setPendingImportFile] = useState(null);
  const importInputRef = useRef(null);
  const dropboxReady = isDropboxConfigured();
  return (
    <div className="space-y-6 max-w-2xl">
      <Card>
        <SectionTitle>Aparência</SectionTitle>
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            {settings.theme === 'dark' ? <Moon size={18} color="var(--text-soft)" /> : <Sun size={18} color="var(--text-soft)" />}
            <div>
              <p className="text-sm font-medium" style={{ color: 'var(--text)' }}>Modo escuro</p>
              <p className="text-xs" style={{ color: 'var(--text-soft)' }}>{settings.theme === 'dark' ? 'Ativado' : 'Desativado'}</p>
            </div>
          </div>
          <ToggleSwitch checked={settings.theme === 'dark'} onChange={(v) => onChangeSettings({ ...settings, theme: v ? 'dark' : 'light' })} label="Modo escuro" />
        </div>
        <div className="mt-5 pt-5" style={{ borderTop: '1px solid var(--border)' }}>
          <p className="text-sm font-medium mb-1" style={{ color: 'var(--text)' }}>Cor de destaque</p>
          <p className="text-xs mb-3" style={{ color: 'var(--text-soft)' }}>Vale tanto para o modo claro quanto para o escuro.</p>
          <div className="flex flex-wrap gap-2.5">
            {Object.entries(COLOR_THEMES).map(([key, t]) => {
              const selected = (settings.colorTheme || 'default') === key;
              const swatch = settings.theme === 'dark' ? t.dark : t.light;
              return (
                <button
                  key={key}
                  type="button"
                  onClick={() => onChangeSettings({ ...settings, colorTheme: key })}
                  title={t.label}
                  aria-label={t.label}
                  className="w-9 h-9 rounded-full flex items-center justify-center focus-ring transition-transform hover:scale-105"
                  style={{ backgroundColor: swatch, boxShadow: selected ? `0 0 0 2px var(--card), 0 0 0 4px ${swatch}` : '0 0 0 2px var(--card)' }}
                >
                  {selected && <Check size={15} color="#fff" strokeWidth={3} />}
                </button>
              );
            })}
          </div>
        </div>
      </Card>

      <Card>
        <SectionTitle>Navegação</SectionTitle>
        <div className="flex items-center justify-between gap-3">
          <div className="min-w-0 flex-1">
            <p className="text-sm font-medium" style={{ color: 'var(--text)' }}>Botão flutuante de novo lançamento</p>
            <p className="text-xs" style={{ color: 'var(--text-soft)' }}>Atalho fixo no canto da tela pra criar um lançamento de qualquer aba, sem precisar abrir o menu. Some em telas grandes, onde o menu lateral já fica sempre visível.</p>
          </div>
          <ToggleSwitch checked={settings.fabEnabled !== false} onChange={(v) => onChangeSettings({ ...settings, fabEnabled: v })} label="Botão flutuante de novo lançamento" />
        </div>
      </Card>

      <VisibilitySettingsSection settings={settings} onChangeSettings={onChangeSettings} />

      <Card>
        <SectionTitle>Meta de economia</SectionTitle>
        <div>
          <FieldLabel>Quanto você quer economizar por mês</FieldLabel>
          <CurrencyInput value={settings.monthlySavingsTarget || 0} onChange={(v) => onChangeSettings({ ...settings, monthlySavingsTarget: v })} />
          <p className="text-xs mt-1" style={{ color: 'var(--text-soft)' }}>Usado no card "Meta de economia" do dashboard (receitas − despesas do mês). Deixe 0 para ocultar a meta. Isso é diferente da aba Metas, que são objetivos como uma viagem ou reserva de emergência.</p>
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
              <ToggleSwitch checked={settings[item.key] !== false} onChange={(v) => onChangeSettings({ ...settings, [item.key]: v })} label={item.label} />
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
        <p className="text-sm mb-4" style={{ color: 'var(--text-soft)' }}>Seus dados ficam salvos automaticamente neste aplicativo. O backup em arquivo inclui tudo, inclusive suas configurações.</p>
        <div className="flex flex-wrap gap-3">
          <Button variant="secondary" icon={Download} onClick={onExportBackup}>Baixar backup (.json)</Button>
          <Button variant="secondary" icon={Upload} onClick={() => importInputRef.current?.click()}>Importar backup (.json)</Button>
          <input ref={importInputRef} type="file" accept="application/json" className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) setPendingImportFile(f); e.target.value = ''; }} />
        </div>
        <div className="flex flex-wrap gap-3 mt-3">
          <Button variant="secondary" onClick={() => setShowConfirmReset(true)}>Restaurar dados de exemplo</Button>
          <Button variant="secondary" onClick={() => setShowConfirmClear(true)}>Limpar dados</Button>
        </div>
      </Card>

      {pendingImportFile && (
        <ConfirmModal
          title="Importar backup"
          description={`Isso substituirá todos os dados e configurações deste navegador pelo conteúdo de "${pendingImportFile.name}". Essa ação não pode ser desfeita.`}
          onConfirm={() => { onImportBackup(pendingImportFile); setPendingImportFile(null); }}
          onClose={() => setPendingImportFile(null)}
        />
      )}

      {showConfirmReset && (
        <ConfirmModal
          title="Restaurar dados de exemplo"
          description="Isso substituirá suas transações, contas, cartões, metas, caixinhas e despesas recorrentes atuais pelos dados de demonstração. Essa ação não pode ser desfeita."
          onConfirm={onReset}
          onClose={() => setShowConfirmReset(false)}
        />
      )}
      {showConfirmClear && (
        <ConfirmModal
          title="Limpar todos os dados"
          description="Isso apaga permanentemente todas as suas transações, contas, cartões, metas, caixinhas, vale-benefícios, investimentos e despesas recorrentes. Não é possível desfazer — se tiver backup no Dropbox, ele continua lá."
          onConfirm={() => { onClearData(); setShowConfirmClear(false); }}
          onClose={() => setShowConfirmClear(false)}
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
  const mainRef = useRef(null);
  const [cardsView, setCardsView] = useState('cartoes');
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [period, setPeriod] = useState('mes');
  const [customRange, setCustomRange] = useState({ start: '', end: '' });
  const [search, setSearch] = useState('');
  const [userName, setUserName] = useState('Thiago Coura');
  const [settings, setSettings] = useState(() => {
    // Lê o tema direto do localStorage de forma síncrona, pra não piscar claro→escuro na abertura.
    const defaults = { notifyDueDates: true, notifyGoals: true, notifyInsights: true, monthlySavingsTarget: 2500, theme: 'light', colorTheme: 'default', fabEnabled: true };
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      const parsed = raw ? JSON.parse(raw) : null;
      return parsed?.settings ? { ...defaults, ...parsed.settings } : defaults;
    } catch { return defaults; }
  });

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
  const [celebration, setCelebration] = useState(null); // { id, origin: {x,y} | null } — explosão de emojis ao concluir uma meta
  function celebrateGoalCompletion(origin) {
    setCelebration({ id: Date.now(), origin });
  }

  // Encolhe o botão + (FAB) enquanto o conteúdo principal está sendo rolado, e devolve o
  // tamanho normal depois de 0,75s parado — cada novo scroll reinicia essa contagem.
  const [fabScrolling, setFabScrolling] = useState(false);
  const fabScrollTimeoutRef = useRef(null);
  function handleContentScroll() {
    setFabScrolling(true);
    if (fabScrollTimeoutRef.current) clearTimeout(fabScrollTimeoutRef.current);
    fabScrollTimeoutRef.current = setTimeout(() => setFabScrolling(false), 750);
  }
  useEffect(() => () => { if (fabScrollTimeoutRef.current) clearTimeout(fabScrollTimeoutRef.current); }, []);

  const [dropboxConnected, setDropboxConnected] = useState(isDropboxConnected());
  const [dropboxBusy, setDropboxBusy] = useState(false);
  const [dropboxLastBackup, setDropboxLastBackup] = useState(null);
  const [dropboxSyncError, setDropboxSyncError] = useState(false);
  const [lastModified, setLastModified] = useState(null);
  const [syncConflict, setSyncConflict] = useState(null);
  const dismissedSyncRef = useRef(null);
  const lastModifiedRef = useRef(null);
  useEffect(() => { lastModifiedRef.current = lastModified; }, [lastModified]);

  function applyLoadedData(loaded) {
    setTransactions(loaded.transactions || buildInitialTransactions());
    setAccounts(loaded.accounts || initialAccounts);
    // "Encaixa" silenciosamente a referência de fatura paga de cada cartão no fechamento de
    // ciclo mais próximo, caso tenha ficado desalinhada por pagamentos feitos antes desta
    // correção (quando era salva como "hoje", não o fechamento real). Roda toda vez que os
    // dados carregam — não precisa persistir a correção nem pedir nada ao usuário, porque o
    // valor corrigido é sempre recalculado do mesmo jeito a partir do que já está salvo.
    const loadedCards = loaded.cards || initialCards;
    setCards(loadedCards.map((c) => (
      c.paidThroughDate ? { ...c, paidThroughDate: snapToClosingBoundary(c, c.paidThroughDate) } : c
    )));
    setGoals(loaded.goals || initialGoals);
    setCaixinhas(loaded.caixinhas || initialCaixinhas);
    setRecurring(loaded.recurring || initialRecurring);
    setBenefits(loaded.benefits || initialBenefits);
    setInvestments(loaded.investments || initialInvestments);
    setUserName(loaded.userName || 'Thiago Coura');
    setSettings(loaded.settings || { notifyDueDates: true, notifyGoals: true, notifyInsights: true, monthlySavingsTarget: 2500, theme: 'light', colorTheme: 'default', fabEnabled: true });
    setLastModified(loaded.lastModified || null);
  }

  // Confere se existe um backup mais novo no Dropbox do que os dados deste navegador — cobre o
  // caso de ter mexido em outro aparelho e esquecido de esperar sincronizar. Não substitui nada
  // sozinho: só avisa e deixa você escolher.
  async function checkForNewerBackup() {
    if (!isDropboxConnected()) return;
    try {
      const remote = await downloadLatestBackup();
      if (!remote || !remote.data.lastModified) return;
      if (remote.filename === dismissedSyncRef.current) return;
      const localTime = lastModifiedRef.current;
      if (!localTime || remote.data.lastModified > localTime) {
        setSyncConflict({ data: remote.data, filename: remote.filename, date: new Date(remote.data.lastModified) });
      }
    } catch (e) { /* checagem em segundo plano — falha silenciosa, tenta de novo depois */ }
  }
  function resolveSyncUseRemote() {
    applyLoadedData(syncConflict.data);
    saveAppData(syncConflict.data);
    addToast('Dados atualizados com a versão mais recente do Dropbox.');
    setSyncConflict(null);
  }
  function resolveSyncKeepLocal() {
    dismissedSyncRef.current = syncConflict.filename;
    // Reenvia os dados daqui para o Dropbox, já que ficaram mais atuais que os de lá.
    uploadBackup(JSON.stringify({ transactions, accounts, cards, goals, caixinhas, recurring, benefits, investments, userName, settings, lastModified }))
      .then(() => setDropboxLastBackup(new Date())).catch(() => {});
    setSyncConflict(null);
  }

  useEffect(() => {
    (async () => {
      // Se voltamos de um login do Dropbox, finaliza a troca do código por token.
      const redirectResult = await handleDropboxRedirect().catch((e) => ({ status: 'error', message: e?.message }));
      if (redirectResult.status === 'connected') {
        setDropboxConnected(true);
        setDropboxSyncError(false);
        addToastSafe('Dropbox conectado com sucesso.');
      } else if (redirectResult.status === 'error') {
        addToastSafe(`Não foi possível conectar ao Dropbox: ${redirectResult.message || 'tente novamente.'}`, 'error');
      }

      const loaded = await loadAppData();
      if (loaded) {
        applyLoadedData(loaded);
        checkForNewerBackup();
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

  // Ao terminar de carregar os dados (localStorage, backup do Dropbox ou dados de exemplo),
  // garante que toda despesa recorrente tenha as próximas ocorrências já lançadas — cobre tanto
  // dados salvos antes dessa função existir quanto o caso de abrir o app depois de um tempo
  // parado (o "buffer" de meses gerados precisa ser reposto). generateRecurringOccurrences já
  // evita duplicar quem já tem lançamento naquele mês, então rodar de novo é sempre seguro.
  useEffect(() => {
    if (isLoading) return;
    let nextTx = transactions;
    let changed = false;
    recurring.forEach((item) => {
      const fresh = generateRecurringOccurrences(item, cards, nextTx);
      if (fresh.length > 0) { nextTx = [...nextTx, ...fresh]; changed = true; }
    });
    if (changed) {
      setTransactions(nextTx);
      persist({ transactions: nextTx });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isLoading]);

  // Reconfere periodicamente (e sempre que a aba volta a ficar em foco) se apareceu um backup
  // mais novo no Dropbox — cobre o caso de ter mexido em outro aparelho enquanto este ficou aberto.
  useEffect(() => {
    const interval = setInterval(() => checkForNewerBackup(), 3 * 60 * 1000);
    function onVisible() { if (document.visibilityState === 'visible') checkForNewerBackup(); }
    document.addEventListener('visibilitychange', onVisible);
    return () => { clearInterval(interval); document.removeEventListener('visibilitychange', onVisible); };
  }, []);

  // Mantém a cor da barra de status do celular (PWA) combinando com o tema atual.
  useEffect(() => {
    const meta = document.querySelector('meta[name="theme-color"]');
    const themeInfo = COLOR_THEMES[settings.colorTheme] || COLOR_THEMES.default;
    if (meta) meta.setAttribute('content', settings.theme === 'dark' ? '#0D1117' : themeInfo.light);
  }, [settings.theme, settings.colorTheme]);

  function addToast(message, type = 'success') {
    const id = uid();
    setToasts((prev) => [...prev, { id, message, type }]);
    setTimeout(() => setToasts((prev) => prev.filter((t) => t.id !== id)), 3500);
  }

  function persist(partial) {
    const now = new Date().toISOString();
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
      lastModified: now,
    };
    setLastModified(now);
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
    setDropboxSyncError(false);
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
  // Backup local em arquivo .json — inclui tudo, inclusive as configurações (aparência, metas
  // gerais, notificações e a personalização da interface), independente do Dropbox estar conectado.
  function exportBackupFile() {
    const payload = { transactions, accounts, cards, goals, caixinhas, recurring, benefits, investments, userName, settings, lastModified: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `cerne-backup-${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    addToast('Backup baixado.');
  }
  function importBackupFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = JSON.parse(e.target.result);
        applyLoadedData(data);
        saveAppData(data);
        addToast('Backup importado com sucesso.');
      } catch (err) {
        addToast('Não foi possível ler esse arquivo — confira se é um backup exportado pelo Cerne.', 'error');
      }
    };
    reader.readAsText(file);
  }

  function goTo(page) { setActivePage(page); if (mainRef.current) mainRef.current.scrollTop = 0; }
  function goToFatura() { setCardsView('fatura'); goTo('cartoes'); }

  /* ---- transações ---- */
  // Uma compra parcelada vira N lançamentos de uma vez (um por parcela), cada um datado pra
  // cair na fatura certa (respeitando o fechamento do cartão) — assim o histórico de gastos
  // por mês fica correto sem precisar de nenhuma lógica especial nos relatórios. Reaproveitada
  // tanto por uma compra parcelada nova quanto pela conversão de um lançamento à vista já
  // existente em parcelado.
  function generateInstallmentTransactions(form, installmentPlan) {
    const { count, totalAmount } = installmentPlan;
    const card = cards.find((c) => c.id === form.cardId);
    const groupId = uid();
    const perBase = Math.round((totalAmount / count) * 100) / 100;
    const lastAmount = Math.round((totalAmount - perBase * (count - 1)) * 100) / 100;
    const [py, pm, pd] = form.date.split('-').map(Number);
    const purchaseDate = new Date(py, pm - 1, pd);
    const rest = { ...form };
    delete rest.installmentPlan;
    const newTxs = [];
    for (let i = 0; i < count; i++) {
      const amount = i === count - 1 ? lastAmount : perBase;
      const dueDate = card ? getInstallmentDueDate(card, purchaseDate, i) : new Date(py, pm - 1 + i, pd);
      newTxs.push({
        ...rest, id: uid(), amount, date: ymd(dueDate), status: rest.cardId ? 'Pendente' : 'Pago',
        installmentGroupId: groupId, installmentIndex: i + 1, installmentCount: count,
      });
    }
    return newTxs;
  }
  function addTransaction(form) {
    if (form.installmentPlan) {
      const newTxs = generateInstallmentTransactions(form, form.installmentPlan);
      const updated = [...newTxs, ...transactions];
      let updatedAccounts = accounts;
      newTxs.forEach((t) => { updatedAccounts = reapplyAccountEffect(updatedAccounts, null, t); });
      setTransactions(updated); setAccounts(updatedAccounts);
      persist({ transactions: updated, accounts: updatedAccounts });
      addToast(`Compra parcelada em ${newTxs.length}x de ${formatBRL(newTxs[0].amount)} lançada.`);
      return;
    }
    const newTx = { ...form, id: uid() };
    const updated = [newTx, ...transactions];
    const updatedAccounts = reapplyAccountEffect(accounts, null, newTx);
    const updatedBenefits = reapplyBenefitEffect(benefits, null, newTx);
    setTransactions(updated); setAccounts(updatedAccounts); setBenefits(updatedBenefits);
    persist({ transactions: updated, accounts: updatedAccounts, benefits: updatedBenefits });
    addToast('Lançamento adicionado com sucesso.');
  }
  function editTransaction(form, toastMessage) {
    const oldTx = transactions.find((t) => t.id === form.id);
    if (form.installmentPlan) {
      const installmentPlan = form.installmentPlan;
      const rest = { ...form };
      const newTxs = generateInstallmentTransactions(rest, installmentPlan);
      const updated = [...newTxs, ...transactions.filter((t) => t.id !== form.id)];
      let updatedAccounts = oldTx ? reapplyAccountEffect(accounts, oldTx, null) : accounts;
      newTxs.forEach((t) => { updatedAccounts = reapplyAccountEffect(updatedAccounts, null, t); });
      const updatedBenefits = oldTx ? reapplyBenefitEffect(benefits, oldTx, null) : benefits;
      setTransactions(updated); setAccounts(updatedAccounts); setBenefits(updatedBenefits);
      persist({ transactions: updated, accounts: updatedAccounts, benefits: updatedBenefits });
      addToast(`Convertido em ${newTxs.length}x de ${formatBRL(newTxs[0].amount)}.`);
      return;
    }
    const updated = transactions.map((t) => (t.id === form.id ? form : t));
    const updatedAccounts = reapplyAccountEffect(accounts, oldTx, form);
    const updatedBenefits = reapplyBenefitEffect(benefits, oldTx, form);
    setTransactions(updated); setAccounts(updatedAccounts); setBenefits(updatedBenefits);
    persist({ transactions: updated, accounts: updatedAccounts, benefits: updatedBenefits });
    // Quando a forma de pagamento muda de um jeito que passa a afetar (ou deixa de afetar) o
    // saldo de uma conta — por exemplo, de cartão de crédito (sem efeito imediato) pra Pix/débito
    // (desconta na hora), ou o contrário — avisamos especificamente, já que é um efeito que o
    // usuário pode não estar esperando ao só trocar a forma de pagamento.
    let message = toastMessage;
    if (!message && oldTx) {
      const oldEffect = transactionBalanceEffect(oldTx);
      const newEffect = transactionBalanceEffect(form);
      if (oldEffect === 0 && newEffect !== 0) {
        const acc = accounts.find((a) => a.id === form.account);
        message = `Lançamento atualizado. ${formatBRL(Math.abs(newEffect))} agora desconta do saldo${acc ? ` de ${acc.bank}` : ''}.`;
      } else if (oldEffect !== 0 && newEffect === 0) {
        const acc = accounts.find((a) => a.id === oldTx.account);
        message = `Lançamento atualizado. ${formatBRL(Math.abs(oldEffect))} foi devolvido ao saldo${acc ? ` de ${acc.bank}` : ''}.`;
      } else if (oldTx.account !== form.account && (oldEffect !== 0 || newEffect !== 0)) {
        message = 'Lançamento atualizado. O saldo das contas envolvidas foi ajustado.';
      }
    }
    addToast(message || 'Lançamento atualizado.');
  }
  // Marca um lançamento pendente (débito agendado, por exemplo) como pago — reaproveita a mesma
  // lógica de saldo do editTransaction, então o valor desconta da conta na hora certa.
  function markTransactionPaid(tx) {
    editTransaction({ ...tx, status: 'Pago' }, 'Lançamento marcado como pago.');
  }
  function deleteTransaction(tx) {
    const updated = transactions.filter((t) => t.id !== tx.id);
    const updatedAccounts = reapplyAccountEffect(accounts, tx, null);
    const updatedBenefits = reapplyBenefitEffect(benefits, tx, null);
    setTransactions(updated); setAccounts(updatedAccounts); setBenefits(updatedBenefits);
    persist({ transactions: updated, accounts: updatedAccounts, benefits: updatedBenefits });
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
    const today = new Date().toISOString().slice(0, 10);
    const history = form.balance > 0 ? [{ date: today, value: form.balance, note: null }] : [];
    const updated = [...caixinhas, { ...form, id: uid(), history }];
    const updatedAccounts = applyBalanceDelta(accounts, form.accountId, -(form.balance || 0));
    setCaixinhas(updated); setAccounts(updatedAccounts); persist({ caixinhas: updated, accounts: updatedAccounts });
    addToast('Caixinha criada.');
  }
  function deleteCaixinha(cx) {
    const updated = caixinhas.filter((c) => c.id !== cx.id);
    // O saldo guardado volta a ficar disponível na conta de origem.
    const updatedAccounts = applyBalanceDelta(accounts, cx.accountId, cx.balance);
    setCaixinhas(updated); setAccounts(updatedAccounts); persist({ caixinhas: updated, accounts: updatedAccounts });
    addToast('Caixinha removida. O saldo voltou para a conta de origem.');
  }
  // Atualiza o valor total da caixinha (não só soma). A diferença é movida de/para a conta
  // vinculada, e cada atualização fica registrada no histórico (com a observação, se houve uma).
  function updateCaixinhaValue(cx, newValue, note) {
    const delta = newValue - cx.balance;
    const today = new Date().toISOString().slice(0, 10);
    const updated = caixinhas.map((c) => (c.id === cx.id
      ? { ...c, balance: newValue, history: [...(c.history || []), { date: today, value: newValue, note }] }
      : c));
    const updatedAccounts = applyBalanceDelta(accounts, cx.accountId, -delta);
    setCaixinhas(updated); setAccounts(updatedAccounts); persist({ caixinhas: updated, accounts: updatedAccounts });
    addToast(delta >= 0 ? `Caixinha "${cx.name}" atualizada: +${formatBRL(delta)}.` : `Caixinha "${cx.name}" atualizada: ${formatBRL(delta)}.`);
  }

  /* ---- cartões ---- */
  function addCard(form) {
    const updated = [...cards, { ...form, id: uid() }];
    setCards(updated); persist({ cards: updated });
    addToast('Cartão adicionado.');
  }
  function editCard(form) {
    const updated = cards.map((c) => (c.id === form.id ? form : c));
    setCards(updated); persist({ cards: updated });
    addToast('Cartão atualizado.');
  }
  function deleteCard(card) {
    const updated = cards.filter((c) => c.id !== card.id);
    setCards(updated); persist({ cards: updated });
    addToast('Cartão removido.');
  }
  // Pagar a fatura é um lançamento de verdade (despesa na conta escolhida) — é isso que faz o
  // saldo da conta cair e "zera" a fatura calculada do cartão a partir de hoje. Se o usuário
  // escolher "nenhuma conta", o pagamento só marca a fatura como quitada, sem mexer em saldo
  // (útil pra quem paga por fora do que é acompanhado no dashboard).
  function payCardInvoice(card, amount, accountId) {
    const today = new Date().toISOString().slice(0, 10);
    const cutoff = ymd(getOpenInvoiceClosingDate(card));
    // Todo lançamento no crédito coberto por este pagamento (dentro da janela que estava em
    // aberto) passa de Pendente para Pago — é isso que faz o status refletir "fatura paga" nas
    // abas de Transações e Fatura mensal, sem mudar o cálculo da fatura em si (que já usava
    // paidThroughDate/cutoff, não o status).
    const coveredTransactions = transactions.map((t) => (
      t.cardId === card.id && t.type === 'despesa'
        && (!card.paidThroughDate || t.date > card.paidThroughDate) && t.date <= cutoff
        ? { ...t, status: 'Pago' }
        : t
    ));
    const newTx = {
      id: uid(), description: `Fatura ${card.bank}`, amount, category: 'Outros', type: 'despesa',
      account: accountId || null, paymentMethod: 'Transferência', date: today, status: 'Pago',
      isInvoicePayment: true,
    };
    const updatedCards = cards.map((c) => (c.id === card.id ? { ...c, paidThroughDate: cutoff } : c));
    setCards(updatedCards);
    if (accountId) {
      const updatedTransactions = [newTx, ...coveredTransactions];
      const updatedAccounts = reapplyAccountEffect(accounts, null, newTx);
      setTransactions(updatedTransactions); setAccounts(updatedAccounts);
      persist({ transactions: updatedTransactions, accounts: updatedAccounts, cards: updatedCards });
    } else {
      setTransactions(coveredTransactions);
      persist({ transactions: coveredTransactions, cards: updatedCards });
    }
    addToast(`Fatura do ${card.bank} paga (${formatBRL(amount)}).`);
  }
  // Traz TODAS as parcelas futuras (ainda não vencidas) de um cartão pra fatura atual, somadas
  // numa parcela só por compra parcelada — é a forma de "antecipar a compra" e quitar o que
  // falta de uma vez, em vez de esperar os próximos meses.
  function advanceAllFutureInstallments(card) {
    const cutoff = ymd(getOpenInvoiceClosingDate(card));
    const future = transactions.filter((t) => t.cardId === card.id && t.installmentGroupId && t.date > cutoff);
    if (future.length === 0) return;
    const groups = {};
    future.forEach((t) => { (groups[t.installmentGroupId] = groups[t.installmentGroupId] || []).push(t); });
    const merged = Object.values(groups).map((group) => {
      const total = group.reduce((s, t) => s + t.amount, 0);
      const base = group[0];
      const cleanDesc = base.description.replace(/\s*\(\d+\/\d+\)\s*$/, '');
      return { ...base, id: uid(), amount: total, date: cutoff, description: `${cleanDesc} (parcelas antecipadas)`, installmentGroupId: null, installmentIndex: null, installmentCount: null };
    });
    const futureIds = new Set(future.map((t) => t.id));
    const kept = transactions.filter((t) => !futureIds.has(t.id));
    const updated = [...merged, ...kept];
    setTransactions(updated); persist({ transactions: updated });
    const total = future.reduce((s, t) => s + t.amount, 0);
    addToast(`${future.length} parcela(s) futura(s) antecipada(s) para a fatura atual (${formatBRL(total)}).`);
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
  function updateBenefit(b, field, newValue, note) {
    const delta = newValue - b[field];
    const historyField = field === 'foodBalance' ? 'foodHistory' : 'mobilityHistory';
    const today = new Date().toISOString().slice(0, 10);
    const updated = benefits.map((x) => (x.id === b.id
      ? { ...x, [field]: newValue, [historyField]: [...(x[historyField] || []), { date: today, value: newValue, note }] }
      : x));
    setBenefits(updated); persist({ benefits: updated });
    addToast(delta >= 0 ? `Saldo atualizado: +${formatBRL(delta)}.` : `Saldo atualizado: ${formatBRL(delta)}.`);
  }

  /* ---- metas ---- */
  function addGoal(form) {
    const today = new Date().toISOString().slice(0, 10);
    const history = form.current > 0 ? [{ date: today, amount: form.current }] : [];
    const updated = [...goals, { ...form, id: uid(), history }];
    setGoals(updated); persist({ goals: updated });
    addToast('Meta criada.');
  }
  // Edita só nome, valor objetivo, previsão e ícone — current/history continuam sob
  // responsabilidade de "Adicionar valor" (addGoalFunds), pra não desalinhar o histórico de
  // aportes com o valor atual.
  function editGoal(form) {
    const updated = goals.map((g) => (g.id === form.id ? { ...g, name: form.name, target: form.target, deadline: form.deadline, icon: form.icon } : g));
    setGoals(updated); persist({ goals: updated });
    addToast('Meta atualizada.');
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
    const item = { ...form, id: uid() };
    const updatedRecurring = [...recurring, item];
    const newOccurrences = generateRecurringOccurrences(item, cards, transactions);
    const updatedTransactions = newOccurrences.length > 0 ? [...transactions, ...newOccurrences] : transactions;
    setRecurring(updatedRecurring);
    if (newOccurrences.length > 0) setTransactions(updatedTransactions);
    persist({ recurring: updatedRecurring, transactions: updatedTransactions });
    addToast(newOccurrences.length > 0 ? `Despesa recorrente adicionada e lançada nos próximos ${newOccurrences.length} meses.` : 'Despesa recorrente adicionada.');
  }
  // Atualiza os dados da despesa recorrente e regenera as ocorrências futuras (ainda não
  // vencidas) com os novos valores — o que já venceu/foi lançado no passado não muda, fica
  // como registro histórico do que realmente foi cobrado naquele mês.
  function editRecurring(form) {
    const updatedRecurring = recurring.map((r) => (r.id === form.id ? form : r));
    const todayStr = ymd(new Date());
    const withoutFutureOccurrences = transactions.filter((t) => !(t.recurringId === form.id && t.date >= todayStr));
    const newOccurrences = generateRecurringOccurrences(form, cards, withoutFutureOccurrences);
    const updatedTransactions = [...withoutFutureOccurrences, ...newOccurrences];
    setRecurring(updatedRecurring); setTransactions(updatedTransactions);
    persist({ recurring: updatedRecurring, transactions: updatedTransactions });
    addToast('Despesa recorrente atualizada.');
  }
  function deleteRecurring(r) {
    const updatedRecurring = recurring.filter((x) => x.id !== r.id);
    const todayStr = ymd(new Date());
    // Remove só as ocorrências futuras geradas automaticamente — o que já venceu fica no
    // histórico, mesmo depois de descadastrar a despesa recorrente.
    const updatedTransactions = transactions.filter((t) => !(t.recurringId === r.id && t.date >= todayStr));
    setRecurring(updatedRecurring); setTransactions(updatedTransactions);
    persist({ recurring: updatedRecurring, transactions: updatedTransactions });
    addToast('Despesa recorrente removida — os lançamentos futuros dela também foram removidos.');
  }
  // Cria um lançamento real de hoje a partir de uma despesa recorrente — usada pra backfill
  // manual (ex: mês que ficou de fora do lote automático). Não entra sozinha nos
  // totais/relatórios, isso é o que efetivamente conta a despesa no seu saldo/relatórios.
  // Usa o cartão ou a conta configurados no cadastro da despesa recorrente; itens antigos sem
  // nenhum dos dois caem na primeira conta, pra não quebrar dados já existentes.
  function addRecurringAsTransaction(item) {
    const today = new Date();
    const alreadyExists = transactions.some((t) => t.type === 'despesa' && t.description === item.name && isSameMonth(t.date, today.getFullYear(), today.getMonth()));
    if (alreadyExists) {
      addToast('Esse mês já tem um lançamento dessa despesa recorrente.', 'error');
      return;
    }
    const card = item.cardId ? cards.find((c) => c.id === item.cardId) : null;
    const fallbackAccount = item.accountId || accounts[0]?.id || null;
    addTransaction({
      description: item.name, amount: item.value, category: item.category, type: 'despesa',
      account: card ? card.accountId : fallbackAccount,
      cardId: card ? card.id : null,
      paymentMethod: card ? 'Cartão de crédito' : 'Não informado',
      date: new Date().toISOString().slice(0, 10),
      status: card ? 'Pendente' : 'Pago',
      recurringId: item.id,
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
  function clearAllData() {
    const empty = { transactions: [], accounts: [], cards: [], goals: [], caixinhas: [], recurring: [], benefits: [], investments: [] };
    setTransactions(empty.transactions); setAccounts(empty.accounts); setCards(empty.cards);
    setGoals(empty.goals); setCaixinhas(empty.caixinhas); setRecurring(empty.recurring); setBenefits(empty.benefits); setInvestments(empty.investments);
    persist(empty);
    addToast('Todos os dados foram apagados.');
  }

  const colorThemeClass = settings.colorTheme && settings.colorTheme !== 'default' ? ` theme-${settings.colorTheme}` : '';
  const rootClassSuffix = `${colorThemeClass}${settings.theme === 'dark' ? ' dark' : ''}`;

  // O menu do <Select> (e outros portais) é renderizado fora da árvore normal pra nunca
  // ficar cortado por um card/modal com "overflow: hidden/auto" no caminho. Pra isso, ele
  // precisa ser anexado fora do .cerne-root (que tem overflow-hidden) — mas as variáveis de
  // cor do tema (--bg, --card, --border, --text-soft, etc.) só existem dentro do
  // .cerne-root. Replicando as mesmas classes no <body>, essas variáveis (que são herdadas)
  // ficam disponíveis em qualquer portal anexado em document.body, no tema certo.
  useEffect(() => {
    document.body.className = `cerne-root${rootClassSuffix}`;
    return () => { document.body.className = ''; };
  }, [rootClassSuffix]);

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
  // Separado da comparação mês atual/anterior acima (que a aba Relatórios usa de propósito) —
  // este segue o período selecionado no cabeçalho do Dashboard (mês/trimestre/ano/personalizado).
  const categoryTotalsForPeriod = useMemo(
    () => computeCategoryTotalsForPeriod(transactions, period, customRange),
    [transactions, period, customRange]
  );

  const insights = useMemo(() => {
    if (isLoading) return [];
    return generateInsights({ transactions, goals, monthlyHistory: realMonthlyHistory, accounts, cards, categoryComparison: realCategoryComparison });
  }, [isLoading, transactions, goals, accounts, realMonthlyHistory, realCategoryComparison]);

  const filteredForSearch = useMemo(() => {
    if (!search.trim()) return null;
    const q = search.toLowerCase();
    return transactions.filter((t) => t.description.toLowerCase().includes(q) || t.category.toLowerCase().includes(q));
  }, [search, transactions]);

  // Resumo dos valores encontrados na busca — soma pelo total filtrado, não só pelos
  // 10 primeiros resultados exibidos na lista. Considera só o que já venceu até hoje: uma
  // compra parcelada (ex: "TV" em 10x) já gera as 10 parcelas como lançamentos futuros no
  // array, então somar tudo daria o valor total da compra, não o que de fato já foi gasto
  // até agora. As parcelas futuras aparecem à parte, não somadas ao total.
  const searchTotals = useMemo(() => {
    if (!filteredForSearch || filteredForSearch.length === 0) return null;
    const todayStr = ymd(new Date());
    const realized = filteredForSearch.filter((t) => t.date <= todayStr);
    const future = filteredForSearch.filter((t) => t.date > todayStr);
    const despesas = realized.filter((t) => t.type === 'despesa').reduce((s, t) => s + t.amount, 0);
    const receitas = realized.filter((t) => t.type === 'receita').reduce((s, t) => s + t.amount, 0);
    const futureTotal = future.reduce((s, t) => s + t.amount, 0);
    const parts = [
      despesas > 0 ? `Total gasto: ${formatBRL(despesas)}` : null,
      receitas > 0 ? `Total recebido: ${formatBRL(receitas)}` : null,
    ].filter(Boolean);
    if (future.length > 0) parts.push(`${future.length} parcela(s) futura(s) — ${formatBRL(futureTotal)}`);
    return parts.join(' · ');
  }, [filteredForSearch]);

  // Tonalidades dos cartões de crédito derivadas da cor de destaque escolhida em Configurações
  // (recalcula só quando o tema de cor muda, não a cada render).
  const cardGradients = useMemo(() => buildCardGradients(settings.colorTheme), [settings.colorTheme]);

  const data = {
    transactions, accounts, cards, goals, caixinhas, recurring, benefits, investments, settings, period, customRange, cardGradients,
    monthlyHistory: realMonthlyHistory, categoryComparison: realCategoryComparison, categoryTotalsForPeriod, investmentsTotal, insights,
  };
  const actions = {
    goTo, goToFatura, editTransaction, deleteTransaction, markTransactionPaid, addGoalFunds, deleteGoal,
    addInvestment, editInvestment, deleteInvestment, addRecurringAsTransaction, payCardInvoice, advanceAllFutureInstallments,
  };

  if (isLoading) {
    return (
      <div className={`cerne-root flex h-screen${rootClassSuffix}`} style={{ backgroundColor: 'var(--bg)', height: '100dvh' }}>
        <style>{GLOBAL_STYLES}</style>
        <LoadingScreen />
      </div>
    );
  }

  return (
    <div className={`cerne-root flex h-screen overflow-hidden${rootClassSuffix}`} style={{ backgroundColor: 'var(--bg)', height: '100dvh' }}>
      <style>{GLOBAL_STYLES}</style>
      <Sidebar activePage={activePage} setActivePage={goTo} sidebarOpen={sidebarOpen} setSidebarOpen={setSidebarOpen} onNewTransaction={() => setModal({ type: 'newTransaction' })} dropboxConnected={dropboxConnected} dropboxLastBackup={dropboxLastBackup} dropboxSyncError={dropboxSyncError} onGoToSettings={() => goTo('configuracoes')} />

      <div className="flex-1 flex flex-col overflow-hidden">
        <Header period={period} setPeriod={setPeriod} customRange={customRange} setCustomRange={setCustomRange} search={search} setSearch={setSearch} setSidebarOpen={setSidebarOpen} insights={insights} />
        {!bannerDismissed && insights[0] && <Banner insight={insights[0]} onDismiss={() => setBannerDismissed(true)} />}

        <main ref={mainRef} className="flex-1 overflow-y-auto px-4 md:px-6 lg:px-8 pt-6 pb-24 lg:pb-6 print-area" onScroll={handleContentScroll}>
          {search.trim() ? (
            <Card>
              <SectionTitle subtitle={searchTotals}>Resultados para "{search}"</SectionTitle>
              {filteredForSearch.length === 0 ? (
                <EmptyState icon={Search} title="Nenhum resultado encontrado" description="Tente buscar por outra descrição ou categoria." />
              ) : (
                <div className="space-y-2">
                  {filteredForSearch.slice(0, 10).map((tx) => (
                    <div key={tx.id} className="flex items-center justify-between gap-3 text-sm p-2.5 rounded-xl hover:bg-black/[0.02]">
                      <div className="flex items-center gap-2 min-w-0 flex-1">
                        <span className="shrink-0" style={{ color: 'var(--text-soft)' }}>{formatDateShortYear(tx.date)}</span>
                        <span className="truncate min-w-0 flex-1" style={{ color: 'var(--text)' }}>{tx.description}</span>
                      </div>
                      <span className="tabular-nums font-medium shrink-0" style={{ color: tx.type === 'receita' ? 'var(--income)' : 'var(--expense)' }}>{formatBRL(tx.amount)}</span>
                    </div>
                  ))}
                </div>
              )}
            </Card>
          ) : (
            <>
              {activePage === 'dashboard' && <DashboardPage data={data} actions={actions} />}
              {activePage === 'transacoes' && <TransactionsPage transactions={transactions} accounts={accounts} cards={cards} benefits={benefits} settings={settings} onAdd={addTransaction} onEdit={editTransaction} onDelete={deleteTransaction} onImport={importTransactions} onMarkPaid={markTransactionPaid} onGoToFatura={goToFatura} />}
              {activePage === 'contas' && <AccountsPage accounts={accounts} caixinhas={caixinhas} transactions={transactions} settings={settings} onAddAccount={addAccount} onDeleteAccount={deleteAccount} onSetAccountThreshold={setAccountThreshold} onSetAccountBalance={setAccountBalance} onAddCaixinha={addCaixinha} onDeleteCaixinha={deleteCaixinha} onUpdateCaixinhaValue={updateCaixinhaValue} />}
              {activePage === 'cartoes' && <CardsPage cards={cards} transactions={transactions} accounts={accounts} recurring={recurring} settings={settings} cardGradients={cardGradients} onAdd={addCard} onEdit={editCard} onDelete={deleteCard} onPayInvoice={payCardInvoice} onAdvanceInstallments={advanceAllFutureInstallments} benefits={benefits} onAddBenefit={addBenefit} onDeleteBenefit={deleteBenefit} onUpdateBenefit={updateBenefit} view={cardsView} onChangeView={setCardsView} onMarkPaid={markTransactionPaid} onEditTransaction={editTransaction} onDeleteTransaction={deleteTransaction} />}
              {activePage === 'investimentos' && <InvestmentsPage investments={investments} settings={settings} onAdd={addInvestment} onEdit={editInvestment} onDelete={deleteInvestment} />}
              {activePage === 'metas' && <GoalsPage goals={goals} onAdd={addGoal} onEdit={editGoal} onAddFunds={addGoalFunds} onDelete={deleteGoal} onCompleted={celebrateGoalCompletion} />}
              {activePage === 'recorrentes' && <RecurringExpensesPage recurring={recurring} accounts={accounts} cards={cards} settings={settings} onAdd={addRecurring} onEdit={editRecurring} onDelete={deleteRecurring} onLaunchNow={addRecurringAsTransaction} />}
              {activePage === 'relatorios' && <ReportsPage monthlyHistory={data.monthlyHistory} categoryComparison={data.categoryComparison} transactions={data.transactions} settings={settings} />}
              {activePage === 'configuracoes' && (
                <SettingsPage
                  settings={settings} onChangeSettings={changeSettings} onReset={resetToSampleData} onClearData={clearAllData}
                  dropboxConnected={dropboxConnected} dropboxBusy={dropboxBusy} dropboxLastBackup={dropboxLastBackup} dropboxSyncError={dropboxSyncError}
                  onConnectDropbox={connectDropbox} onDisconnectDropbox={disconnectDropboxAccount}
                  onBackupNow={backupNowToDropbox} onRestoreFromDropbox={restoreFromDropbox}
                  onExportBackup={exportBackupFile} onImportBackup={importBackupFile}
                />
              )}
            </>
          )}
        </main>
      </div>

      {modal?.type === 'newTransaction' && (
        <TransactionForm accounts={accounts} cards={cards} benefits={benefits} transactions={transactions} onSave={(f) => { addTransaction(f); setModal(null); }} onClose={() => setModal(null)} />
      )}
      {syncConflict && (
        <SyncConflictModal date={syncConflict.date} onUseRemote={resolveSyncUseRemote} onKeepLocal={resolveSyncKeepLocal} />
      )}

      <ToastContainer toasts={toasts} />
      {settings.fabEnabled !== false && <FAB onClick={() => setModal({ type: 'newTransaction' })} shrink={fabScrolling} />}
      {celebration && <GoalCelebration key={celebration.id} origin={celebration.origin} onDone={() => setCelebration(null)} />}
    </div>
  );
}
