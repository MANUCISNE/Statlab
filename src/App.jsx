import React, { useState, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  Upload, FileSpreadsheet, FlaskConical, Sparkles, Loader2,
  ChevronRight, ChevronLeft, AlertCircle, CheckCircle2, Download,
  BarChart3, TrendingUp, GitBranch, Activity, Sigma, Grid3x3, ArrowLeft,
  Brain, Lightbulb
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  ScatterChart, Scatter, Line, Legend, Cell
} from "recharts";

// >>> AJUSTE VERCEL: endpoint do proxy serverless (criado em /api/claude.js)
const API_ENDPOINT = "/api/claude";

const STYLES = `
  @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700;9..144,800&family=JetBrains+Mono:wght@400;500;600&family=Inter+Tight:wght@400;500;600;700&display=swap');

  .statlab-root {
    --bg: #0f1115; --bg-elev: #161922; --bg-card: #1a1e29;
    --border: #262b38; --border-strong: #353c4d;
    --text: #e8eaf0; --text-dim: #9098a8; --text-faint: #5a6175;
    --accent: #d4ff3a; --accent-warm: #ff8c42; --accent-cool: #6ab7ff;
    --accent-purple: #a78bfa; --danger: #ff5470; --success: #4ade80;
    --serif: 'Fraunces', Georgia, serif;
    --sans: 'Inter Tight', system-ui, sans-serif;
    --mono: 'JetBrains Mono', monospace;
    background: var(--bg); color: var(--text); font-family: var(--sans);
    min-height: 100vh; padding: 32px 24px;
    background-image:
      radial-gradient(circle at 20% 0%, rgba(212, 255, 58, 0.04) 0%, transparent 40%),
      radial-gradient(circle at 80% 100%, rgba(106, 183, 255, 0.04) 0%, transparent 40%);
  }
  .container { max-width: 1180px; margin: 0 auto; }
  .header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 40px; }
  .brand-mark { font-family: var(--serif); font-style: italic; font-weight: 600; font-size: 38px; letter-spacing: -0.02em; }
  .brand-mark em { color: var(--accent); font-style: italic; }
  .brand-sub { font-family: var(--mono); font-size: 11px; color: var(--text-faint); text-transform: uppercase; letter-spacing: 0.15em; }
  .header-meta { font-family: var(--mono); font-size: 11px; color: var(--text-faint); text-align: right; }

  .mode-tabs { display: grid; grid-template-columns: 1fr 1fr; gap: 0; background: var(--bg-card); border: 1px solid var(--border); border-radius: 3px; padding: 4px; margin-bottom: 16px; }
  .mode-tab { padding: 14px 20px; background: transparent; border: none; color: var(--text-dim); cursor: pointer; font-family: var(--sans); font-size: 13px; font-weight: 500; border-radius: 2px; display: flex; align-items: center; justify-content: center; gap: 10px; transition: all 0.15s; }
  .mode-tab.active { background: var(--bg-elev); color: var(--text); }
  .mode-tab:hover:not(.active) { color: var(--text); }

  .stepper { display: flex; gap: 0; margin-bottom: 32px; border-top: 1px solid var(--border); border-bottom: 1px solid var(--border); }
  .step { flex: 1; padding: 16px 20px; display: flex; align-items: center; gap: 12px; border-right: 1px solid var(--border); }
  .step:last-child { border-right: none; }
  .step-num { font-family: var(--mono); font-size: 11px; color: var(--text-faint); width: 22px; height: 22px; border-radius: 50%; display: flex; align-items: center; justify-content: center; border: 1px solid var(--border-strong); }
  .step.active .step-num { background: var(--accent); color: #000; border-color: var(--accent); font-weight: 700; }
  .step.done .step-num { background: var(--text-dim); color: var(--bg); border-color: var(--text-dim); }
  .step-label { font-size: 12px; color: var(--text-dim); }
  .step.active .step-label { color: var(--text); font-weight: 500; }
  .step.done .step-label { color: var(--text); }

  .card { background: var(--bg-card); border: 1px solid var(--border); padding: 28px; border-radius: 4px; margin-bottom: 16px; }
  .card-title { font-family: var(--serif); font-size: 28px; font-weight: 500; letter-spacing: -0.01em; margin-bottom: 6px; }
  .card-title em { font-style: italic; color: var(--accent); }
  .card-desc { color: var(--text-dim); font-size: 14px; margin-bottom: 24px; line-height: 1.6; }

  .test-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap: 12px; }
  .test-card { background: var(--bg-elev); border: 1px solid var(--border); padding: 20px; cursor: pointer; transition: all 0.15s ease; border-radius: 3px; display: flex; flex-direction: column; gap: 10px; text-align: left; }
  .test-card:hover { border-color: var(--accent); }
  .test-card.selected { border-color: var(--accent); background: rgba(212, 255, 58, 0.04); }
  .test-card-icon { color: var(--accent); }
  .test-card-name { font-family: var(--serif); font-size: 18px; font-weight: 600; letter-spacing: -0.01em; }
  .test-card-desc { color: var(--text-dim); font-size: 12px; line-height: 1.5; }
  .test-card-meta { font-family: var(--mono); font-size: 10px; color: var(--text-faint); text-transform: uppercase; letter-spacing: 0.1em; }
  .test-card-template-btn { margin-top: 8px; padding: 8px 12px; background: transparent; border: 1px solid var(--border-strong); color: var(--text-dim); font-family: var(--mono); font-size: 10px; letter-spacing: 0.1em; text-transform: uppercase; cursor: pointer; border-radius: 2px; display: flex; align-items: center; gap: 6px; justify-content: center; transition: all 0.15s; }
  .test-card-template-btn:hover { color: var(--accent); border-color: var(--accent); background: rgba(212, 255, 58, 0.05); }

  .ai-recommend-card { background: linear-gradient(135deg, rgba(167, 139, 250, 0.08) 0%, rgba(106, 183, 255, 0.04) 100%); border: 1px solid var(--accent-purple); padding: 28px; border-radius: 4px; }
  .ai-badge { display: inline-flex; align-items: center; gap: 6px; background: rgba(167, 139, 250, 0.15); color: var(--accent-purple); padding: 4px 10px; border-radius: 99px; font-family: var(--mono); font-size: 10px; text-transform: uppercase; letter-spacing: 0.1em; margin-bottom: 14px; }
  textarea { background: var(--bg); border: 1px solid var(--border-strong); color: var(--text); padding: 14px 16px; font-family: var(--sans); font-size: 14px; width: 100%; border-radius: 3px; outline: none; resize: vertical; min-height: 100px; line-height: 1.6; box-sizing: border-box; }
  textarea:focus { border-color: var(--accent-purple); }

  .recommend-result { background: var(--bg-elev); border: 1px solid var(--accent-purple); padding: 24px; border-radius: 4px; margin-top: 16px; }
  .recommend-test-name { font-family: var(--serif); font-size: 26px; font-weight: 500; color: var(--accent); margin-bottom: 8px; letter-spacing: -0.01em; }
  .recommend-rationale { font-family: var(--serif); font-size: 16px; line-height: 1.7; color: var(--text); margin: 16px 0; }
  .recommend-mapping { background: var(--bg); border: 1px solid var(--border); padding: 16px; border-radius: 3px; margin-top: 12px; }
  .recommend-mapping-row { display: grid; grid-template-columns: 200px 1fr; padding: 8px 0; border-bottom: 1px solid var(--border); font-size: 13px; }
  .recommend-mapping-row:last-child { border-bottom: none; }
  .recommend-mapping-row .mk { font-family: var(--mono); font-size: 11px; color: var(--text-faint); text-transform: uppercase; letter-spacing: 0.1em; }
  .recommend-mapping-row .mv { font-family: var(--mono); color: var(--accent); }
  .recommend-alts { margin-top: 16px; padding-top: 16px; border-top: 1px dashed var(--border); }
  .recommend-alts-title { font-family: var(--mono); font-size: 10px; text-transform: uppercase; letter-spacing: 0.15em; color: var(--text-faint); margin-bottom: 8px; }
  .alt-item { font-size: 13px; color: var(--text-dim); padding: 6px 0; border-bottom: 1px dotted var(--border); }
  .alt-item:last-child { border-bottom: none; }
  .alt-item strong { color: var(--text); }

  .dropzone { border: 2px dashed var(--border-strong); padding: 64px 32px; text-align: center; border-radius: 4px; cursor: pointer; transition: all 0.2s; background: var(--bg-elev); }
  .dropzone:hover, .dropzone.drag { border-color: var(--accent); background: rgba(212, 255, 58, 0.03); }
  .dropzone-icon { margin: 0 auto 16px; color: var(--text-dim); }
  .dropzone-text { font-family: var(--serif); font-size: 22px; font-weight: 500; margin-bottom: 6px; }
  .dropzone-hint { color: var(--text-dim); font-size: 13px; }

  .file-info { display: flex; align-items: center; gap: 14px; padding: 16px; background: var(--bg-elev); border: 1px solid var(--border); border-radius: 3px; margin-bottom: 20px; }
  .file-info-icon { color: var(--accent); }
  .file-info-name { font-weight: 500; font-size: 14px; }
  .file-info-meta { font-family: var(--mono); font-size: 11px; color: var(--text-faint); }

  .table-wrap { overflow-x: auto; border: 1px solid var(--border); border-radius: 3px; }
  table.preview { width: 100%; border-collapse: collapse; font-size: 12px; }
  table.preview th { background: var(--bg-elev); padding: 10px 14px; text-align: left; font-family: var(--mono); font-size: 10px; text-transform: uppercase; letter-spacing: 0.1em; color: var(--text-dim); border-bottom: 1px solid var(--border); border-right: 1px solid var(--border); white-space: nowrap; }
  table.preview td { padding: 8px 14px; border-bottom: 1px solid var(--border); border-right: 1px solid var(--border); font-family: var(--mono); white-space: nowrap; }
  table.preview tr:last-child td { border-bottom: none; }

  .mapping-row { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; padding: 16px 0; border-bottom: 1px solid var(--border); }
  .mapping-row:last-child { border-bottom: none; }
  .mapping-label { font-family: var(--serif); font-size: 16px; font-weight: 500; }
  .mapping-hint { font-size: 12px; color: var(--text-dim); margin-top: 2px; }

  select, input[type="text"] { background: var(--bg); border: 1px solid var(--border-strong); color: var(--text); padding: 10px 12px; font-family: var(--sans); font-size: 13px; width: 100%; border-radius: 3px; outline: none; box-sizing: border-box; }
  select:focus { border-color: var(--accent); }

  .actions { display: flex; justify-content: space-between; gap: 12px; margin-top: 32px; }
  button.btn { background: var(--accent); color: #000; border: none; padding: 12px 24px; font-family: var(--sans); font-weight: 600; font-size: 14px; cursor: pointer; display: inline-flex; align-items: center; gap: 8px; border-radius: 3px; transition: transform 0.1s; }
  button.btn:hover:not(:disabled) { transform: translateY(-1px); }
  button.btn:disabled { opacity: 0.4; cursor: not-allowed; }
  button.btn.ghost { background: transparent; color: var(--text-dim); border: 1px solid var(--border-strong); }
  button.btn.ghost:hover { color: var(--text); border-color: var(--text-dim); }
  button.btn.purple { background: var(--accent-purple); color: #000; }

  .results-header { display: flex; justify-content: space-between; align-items: flex-end; border-bottom: 1px solid var(--border); padding-bottom: 20px; margin-bottom: 28px; }
  .results-title { font-family: var(--serif); font-size: 36px; font-weight: 500; letter-spacing: -0.02em; }
  .results-title em { font-style: italic; color: var(--accent); }
  .results-subtitle { color: var(--text-dim); font-size: 13px; margin-top: 6px; font-family: var(--mono); }

  .metric-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 1px; background: var(--border); border: 1px solid var(--border); border-radius: 3px; overflow: hidden; margin-bottom: 24px; }
  .metric { background: var(--bg-card); padding: 20px; }
  .metric-label { font-family: var(--mono); font-size: 10px; color: var(--text-faint); text-transform: uppercase; letter-spacing: 0.12em; margin-bottom: 8px; }
  .metric-value { font-family: var(--serif); font-size: 32px; font-weight: 500; letter-spacing: -0.02em; line-height: 1; }
  .metric-value.sig { color: var(--accent); }
  .metric-value.notsig { color: var(--text-dim); }
  .metric-sub { font-family: var(--mono); font-size: 11px; color: var(--text-dim); margin-top: 8px; }

  .section-h { font-family: var(--mono); font-size: 11px; text-transform: uppercase; letter-spacing: 0.15em; color: var(--text-dim); margin: 28px 0 12px; display: flex; align-items: center; gap: 10px; }
  .section-h::after { content: ''; flex: 1; height: 1px; background: var(--border); }

  .interpretation { font-family: var(--serif); font-size: 17px; line-height: 1.7; color: var(--text); background: var(--bg-elev); border-left: 3px solid var(--accent); padding: 20px 24px; border-radius: 0 3px 3px 0; }
  .interpretation strong { color: var(--accent); font-weight: 600; }

  .chart-card { background: var(--bg-elev); border: 1px solid var(--border); padding: 20px; border-radius: 3px; margin-bottom: 16px; }
  .chart-title { font-family: var(--serif); font-size: 18px; font-weight: 500; margin-bottom: 4px; letter-spacing: -0.01em; }
  .chart-desc { font-size: 12px; color: var(--text-dim); margin-bottom: 16px; }

  .alert { display: flex; gap: 12px; padding: 14px 16px; border-radius: 3px; background: rgba(255, 84, 112, 0.08); border: 1px solid rgba(255, 84, 112, 0.3); color: #ffadbb; font-size: 13px; margin-bottom: 16px; }
  .info-alert { display: flex; gap: 12px; padding: 14px 16px; border-radius: 3px; background: rgba(106, 183, 255, 0.06); border: 1px solid rgba(106, 183, 255, 0.3); color: var(--text-dim); font-size: 13px; margin-bottom: 16px; line-height: 1.6; }

  .loading-state { text-align: center; padding: 64px 20px; }
  .loading-state .spin-icon { animation: spin 1.2s linear infinite; color: var(--accent); margin-bottom: 20px; }
  @keyframes spin { to { transform: rotate(360deg); } }
  .loading-state-text { font-family: var(--serif); font-size: 24px; font-style: italic; margin-bottom: 6px; }
  .loading-state-sub { font-family: var(--mono); font-size: 11px; color: var(--text-faint); text-transform: uppercase; letter-spacing: 0.15em; }

  .detail-list { border: 1px solid var(--border); border-radius: 3px; }
  .detail-row { display: grid; grid-template-columns: 200px 1fr; padding: 12px 16px; border-bottom: 1px solid var(--border); font-size: 13px; }
  .detail-row:last-child { border-bottom: none; }
  .detail-row .k { font-family: var(--mono); font-size: 11px; color: var(--text-faint); text-transform: uppercase; letter-spacing: 0.1em; }
  .detail-row .v { font-family: var(--mono); color: var(--text); }
`;

const TESTS = [
  { id: "ttest_indep", name: "Teste t independente", icon: GitBranch, family: "Comparação de médias",
    desc: "Compara médias de 2 grupos independentes (paramétrico).",
    inputs: [
      { key: "group", label: "Coluna de grupo", hint: "Variável categórica com 2 níveis", type: "column" },
      { key: "value", label: "Coluna numérica", hint: "Variável dependente (contínua)", type: "column" },
    ],
    template: {
      headers: ["grupo", "valor"],
      example: [["Controle", 12.4], ["Controle", 11.8], ["Controle", 13.1], ["Controle", 12.9], ["Controle", 11.5], ["Tratamento", 15.2], ["Tratamento", 14.8], ["Tratamento", 16.1], ["Tratamento", 15.5], ["Tratamento", 14.9]],
      instructions: ["TESTE: t de Student independente", "USE QUANDO: comparar a média de 2 grupos diferentes", "EX: medicamento vs placebo, masculino vs feminino", "", "COLUNA 'grupo': nome do grupo (texto) — DEVE TER 2 NÍVEIS", "COLUNA 'valor': valor numérico observado", "", "Cada linha = uma observação. Apague os exemplos antes de usar."],
    },
  },
  { id: "ttest_paired", name: "Teste t pareado", icon: GitBranch, family: "Comparação de médias",
    desc: "Compara medidas pareadas (antes/depois, mesmo sujeito).",
    inputs: [
      { key: "before", label: "Medida 1 (antes)", hint: "Coluna numérica", type: "column" },
      { key: "after", label: "Medida 2 (depois)", hint: "Coluna numérica", type: "column" },
    ],
    template: {
      headers: ["sujeito_id", "antes", "depois"],
      example: [[1, 120, 115], [2, 135, 128], [3, 142, 134], [4, 128, 122], [5, 150, 141], [6, 119, 116], [7, 138, 130], [8, 145, 138]],
      instructions: ["TESTE: t pareado", "USE QUANDO: medir a MESMA pessoa duas vezes (antes/depois)", "EX: pressão arterial antes/depois do tratamento", "", "COLUNA 'sujeito_id': identificador único (opcional)", "COLUNA 'antes': valor na 1ª medida", "COLUNA 'depois': valor na 2ª medida", "", "Cada linha = UM sujeito com suas DUAS medidas pareadas."],
    },
  },
  { id: "anova_oneway", name: "ANOVA one-way", icon: BarChart3, family: "Comparação de médias",
    desc: "Compara médias de 3+ grupos independentes.",
    inputs: [
      { key: "group", label: "Coluna de grupo", hint: "Variável categórica com 3+ níveis", type: "column" },
      { key: "value", label: "Coluna numérica", hint: "Variável dependente", type: "column" },
    ],
    template: {
      headers: ["grupo", "valor"],
      example: [["A", 22.1], ["A", 21.5], ["A", 23.0], ["A", 22.8], ["B", 25.4], ["B", 26.1], ["B", 25.8], ["B", 25.0], ["C", 28.7], ["C", 29.2], ["C", 28.1], ["C", 29.5]],
      instructions: ["TESTE: ANOVA one-way", "USE QUANDO: comparar médias de 3+ grupos independentes", "EX: 3 dosagens de medicamento, 4 escolas, 5 regiões", "", "COLUNA 'grupo': nome do grupo (3+ níveis)", "COLUNA 'valor': valor numérico", "", "Se ANOVA der significativa, será feito post-hoc (Tukey)."],
    },
  },
  { id: "regression_linear", name: "Regressão linear", icon: TrendingUp, family: "Relação entre variáveis",
    desc: "Modela relação linear entre preditora(s) e desfecho.",
    inputs: [
      { key: "y", label: "Variável dependente (Y)", hint: "Desfecho — coluna numérica", type: "column" },
      { key: "x", label: "Variável(eis) preditora(s) (X)", hint: "Ctrl/Cmd+clique para múltiplas", type: "columns" },
    ],
    template: {
      headers: ["y_desfecho", "x1_preditor", "x2_preditor", "x3_preditor"],
      example: [[85, 12, 5, 2.3], [92, 15, 6, 2.8], [78, 10, 4, 1.9], [95, 17, 7, 3.1], [82, 11, 5, 2.1], [88, 14, 6, 2.6], [90, 16, 7, 2.9], [76, 9, 3, 1.7]],
      instructions: ["TESTE: Regressão linear (simples ou múltipla)", "USE QUANDO: estimar como X prediz Y", "EX: pressão arterial em função de idade, peso e dieta", "", "COLUNA 'y_desfecho': variável dependente", "COLUNAS 'x1_preditor', 'x2_preditor'...: independentes (numéricas)", "", "Pode renomear e adicionar/remover colunas X."],
    },
  },
  { id: "correlation", name: "Correlação", icon: Activity, family: "Relação entre variáveis",
    desc: "Mede força e direção de associação entre 2 variáveis.",
    inputs: [
      { key: "x", label: "Variável X", hint: "Coluna numérica", type: "column" },
      { key: "y", label: "Variável Y", hint: "Coluna numérica", type: "column" },
    ],
    template: {
      headers: ["variavel_x", "variavel_y"],
      example: [[1.2, 4.5], [2.1, 5.8], [3.0, 7.1], [1.8, 5.2], [2.5, 6.4], [3.5, 8.0], [1.0, 4.1], [2.8, 6.9], [3.2, 7.5], [2.0, 5.5]],
      instructions: ["TESTE: Correlação (Pearson/Spearman)", "USE QUANDO: medir o quanto 2 variáveis variam juntas (sem causalidade)", "EX: relação entre altura e peso", "", "COLUNA 'variavel_x' e 'variavel_y': numéricas", "Cada linha = um par X,Y."],
    },
  },
  { id: "chisquare", name: "Qui-quadrado", icon: Grid3x3, family: "Variáveis categóricas",
    desc: "Testa associação entre 2 variáveis categóricas.",
    inputs: [
      { key: "var1", label: "Variável categórica 1", hint: "Linhas da tabela", type: "column" },
      { key: "var2", label: "Variável categórica 2", hint: "Colunas da tabela", type: "column" },
    ],
    template: {
      headers: ["variavel_1", "variavel_2"],
      example: [["Fumante", "Doente"], ["Fumante", "Doente"], ["Fumante", "Saudavel"], ["Fumante", "Doente"], ["Nao_fumante", "Saudavel"], ["Nao_fumante", "Saudavel"], ["Nao_fumante", "Doente"], ["Nao_fumante", "Saudavel"], ["Nao_fumante", "Saudavel"], ["Fumante", "Doente"]],
      instructions: ["TESTE: Qui-quadrado (ou Fisher se células < 5)", "USE QUANDO: testar associação entre 2 variáveis categóricas", "EX: tabagismo × doença", "", "Cada linha = UM sujeito (não pré-tabulado).", "A IA constrói a tabela de contingência."],
    },
  },
  { id: "descriptive", name: "Estatística descritiva", icon: Sigma, family: "Resumo",
    desc: "Média, mediana, DP, quartis, normalidade.",
    inputs: [
      { key: "cols", label: "Colunas para descrever", hint: "1+ colunas numéricas", type: "columns" },
      { key: "groupBy", label: "Agrupar por (opcional)", hint: "Coluna categórica ou em branco", type: "column_opt" },
    ],
    template: {
      headers: ["grupo", "variavel_1", "variavel_2", "variavel_3"],
      example: [["A", 12.5, 88, 3.4], ["A", 13.1, 92, 3.6], ["A", 11.9, 85, 3.2], ["B", 15.2, 95, 4.1], ["B", 14.8, 98, 4.3], ["B", 15.5, 101, 4.5]],
      instructions: ["ANÁLISE: Estatística descritiva", "USE QUANDO: quiser sumarizar dados", "", "COLUNA 'grupo' (opcional): para subgrupos", "COLUNAS 'variavel_1'...: numéricas a descrever"],
    },
  },
];

function detectType(values) {
  const clean = values.filter(v => v !== null && v !== undefined && v !== "");
  if (clean.length === 0) return "empty";
  const numeric = clean.filter(v => typeof v === "number" || (!isNaN(parseFloat(v)) && isFinite(v)));
  if (numeric.length / clean.length > 0.85) return "numeric";
  return "categorical";
}

function downloadTemplate(test) {
  const wb = XLSX.utils.book_new();
  const instructionRows = [
    [`StatLab — Template: ${test.name}`], [],
    ...test.template.instructions.map(line => [line]), [],
    ["Próximo: vá para a aba 'Dados' e preencha com suas observações."],
  ];
  const wsInst = XLSX.utils.aoa_to_sheet(instructionRows);
  wsInst["!cols"] = [{ wch: 90 }];
  XLSX.utils.book_append_sheet(wb, wsInst, "Instruções");

  const dataRows = [test.template.headers, ...test.template.example];
  const wsData = XLSX.utils.aoa_to_sheet(dataRows);
  wsData["!cols"] = test.template.headers.map(() => ({ wch: 18 }));
  XLSX.utils.book_append_sheet(wb, wsData, "Dados");
  XLSX.writeFile(wb, `statlab_template_${test.id}.xlsx`);
}

// >>> AJUSTE VERCEL: chamada agora vai para nosso proxy /api/claude
async function callClaude(prompt, maxTokens = 2000) {
  const response = await fetch(API_ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt, max_tokens: maxTokens }),
  });
  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`API ${response.status}: ${errText.slice(0, 200)}`);
  }
  return response.json();
}

function parseJsonFromResponse(json) {
  let text = json.content.filter(c => c.type === "text").map(c => c.text).join("\n");
  text = text.replace(/```json\s*/g, "").replace(/```\s*$/g, "").trim();
  const firstBrace = text.indexOf("{");
  const lastBrace = text.lastIndexOf("}");
  if (firstBrace > -1 && lastBrace > firstBrace) text = text.slice(firstBrace, lastBrace + 1);
  return JSON.parse(text);
}

export default function App() {
  const [mode, setMode] = useState("guided");
  const [step, setStep] = useState(0);
  const [selectedTest, setSelectedTest] = useState(null);
  const [researchQuestion, setResearchQuestion] = useState("");
  const [aiRecommendation, setAiRecommendation] = useState(null);
  const [file, setFile] = useState(null);
  const [data, setData] = useState(null);
  const [columns, setColumns] = useState([]);
  const [mapping, setMapping] = useState({});
  const [result, setResult] = useState(null);
  const [loading, setLoading] = useState(false);
  const [loadingRec, setLoadingRec] = useState(false);
  const [error, setError] = useState(null);
  const [drag, setDrag] = useState(false);
  const fileRef = useRef();

  const currentTest = useMemo(() => TESTS.find(t => t.id === selectedTest), [selectedTest]);
  const columnTypes = useMemo(() => {
    if (!data || !columns.length) return {};
    const types = {};
    columns.forEach(c => { types[c] = detectType(data.map(r => r[c])); });
    return types;
  }, [data, columns]);

  function handleFile(f) {
    setError(null);
    if (!f) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        let sheetName = wb.SheetNames[0];
        if (wb.SheetNames.includes("Dados")) sheetName = "Dados";
        const sheet = wb.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: null });
        if (!json.length) throw new Error("Planilha vazia.");
        setFile(f);
        setData(json);
        setColumns(Object.keys(json[0]));
        setMapping({});
        setAiRecommendation(null);
      } catch (err) {
        setError(`Erro lendo arquivo: ${err.message}`);
      }
    };
    reader.readAsArrayBuffer(f);
  }

  function handleDrop(e) {
    e.preventDefault();
    setDrag(false);
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  }

  function setMap(key, value) { setMapping(m => ({ ...m, [key]: value })); }

  function mappingComplete() {
    if (!currentTest) return false;
    return currentTest.inputs.every(inp => {
      if (inp.type === "column_opt") return true;
      const v = mapping[inp.key];
      if (inp.type === "columns") return Array.isArray(v) && v.length > 0;
      return !!v;
    });
  }

  async function getAiRecommendation() {
    if (!data || !researchQuestion.trim()) return;
    setLoadingRec(true);
    setError(null);
    setAiRecommendation(null);

    try {
      const sample = data.slice(0, 15);
      const columnsSummary = columns.map(c => {
        const vals = data.map(r => r[c]).filter(v => v !== null && v !== undefined && v !== "");
        return { name: c, type: columnTypes[c], uniqueValues: new Set(vals).size, examples: [...new Set(vals)].slice(0, 5) };
      });
      const availableTests = TESTS.map(t => ({ id: t.id, name: t.name, when_to_use: t.desc, required_inputs: t.inputs.map(i => ({ key: i.key, label: i.label, type: i.type })) }));

      const prompt = `You are an expert biostatistician helping a researcher choose the right statistical test.

THEIR RESEARCH QUESTION: "${researchQuestion}"

THEIR DATASET (${data.length} rows, ${columns.length} columns):
Column summary: ${JSON.stringify(columnsSummary, null, 2)}
Sample rows: ${JSON.stringify(sample, null, 2)}

AVAILABLE TESTS: ${JSON.stringify(availableTests, null, 2)}

Task:
1. Choose the SINGLE BEST test.
2. Map their actual column names to the required inputs.
3. Explain WHY in Portuguese (Brazilian), didactically.
4. Suggest 1-2 alternatives.
5. Flag data quality concerns.

Respond ONLY with valid JSON (no markdown, no fences):
{
  "recommended_test_id": "<id>",
  "recommended_test_name": "Nome PT-BR",
  "rationale": "Explicação PT-BR com **negrito** markdown",
  "mapping": { "input_key": "nome_real_coluna" or ["c1","c2"] },
  "alternatives": [{"test_name": "Nome", "when": "quando seria melhor"}],
  "data_quality_notes": "Avisos PT-BR (pode ser vazio)"
}`;

      const json = await callClaude(prompt, 2000);
      setAiRecommendation(parseJsonFromResponse(json));
    } catch (err) {
      setError(`Falha na recomendação: ${err.message}`);
    } finally {
      setLoadingRec(false);
    }
  }

  function acceptRecommendation() {
    if (!aiRecommendation) return;
    setSelectedTest(aiRecommendation.recommended_test_id);
    setMapping(aiRecommendation.mapping || {});
    setStep(2);
  }

  async function runAnalysis() {
    setLoading(true);
    setError(null);
    setStep(3);

    try {
      const usedCols = new Set();
      Object.values(mapping).forEach(v => {
        if (Array.isArray(v)) v.forEach(x => usedCols.add(x));
        else if (v) usedCols.add(v);
      });

      const slice = data.map(row => {
        const o = {};
        usedCols.forEach(c => o[c] = row[c]);
        return o;
      });

      const prompt = `You are an expert biostatistician. Run the analysis and respond ONLY with valid JSON.

ANALYSIS: ${currentTest.name} (id: ${currentTest.id})
DESCRIPTION: ${currentTest.desc}

VARIABLE MAPPING:
${currentTest.inputs.map(inp => `- ${inp.label} (${inp.key}): ${JSON.stringify(mapping[inp.key])}`).join("\n")}

DATASET (${slice.length} rows):
${JSON.stringify(slice.slice(0, 500))}
${slice.length > 500 ? `\n[... ${slice.length - 500} more truncated]` : ""}

Instructions:
1. Compute rigorously. Check assumptions (normality, variance homogeneity).
2. If assumptions violated, ALSO report non-parametric alternative.
3. Compute effect size (Cohen's d, eta², r, etc.).
4. Write didactic interpretation in PT-BR with **bold** markdown.
5. Return chart-ready data.

JSON shape:
{
  "test_performed": "Nome",
  "assumptions": { "checked": ["pressuposto: status — detalhe"], "recommendation": "PT-BR" },
  "metrics": [{"label": "p-valor", "value": "0.023", "significant": true, "sub": "α=0.05"}],
  "summary_table": { "headers": [...], "rows": [[...]] },
  "interpretation": "PT-BR com **negrito**",
  "charts": [{"type": "bar"|"scatter"|"histogram"|"boxplot_summary", "title": "...", "description": "...", "data": [...], "xKey": "name", "yKey": "value"}],
  "additional_notes": "PT-BR opcional"
}`;

      const json = await callClaude(prompt, 4000);
      setResult(parseJsonFromResponse(json));
    } catch (err) {
      setError(`Falha na análise: ${err.message}`);
      setResult(null);
    } finally {
      setLoading(false);
    }
  }

  function reset() {
    setStep(0); setSelectedTest(null); setFile(null); setData(null);
    setColumns([]); setMapping({}); setResult(null); setError(null);
    setAiRecommendation(null); setResearchQuestion("");
  }

  function switchMode(newMode) {
    setMode(newMode); setStep(0); setSelectedTest(null);
    setAiRecommendation(null); setResearchQuestion("");
  }

  const stepsForMode = mode === "ai"
    ? ["Dados", "Pergunta", "Mapeamento", "Resultados"]
    : ["Análise", "Dados", "Mapeamento", "Resultados"];

  return (
    <>
      <style>{STYLES}</style>
      <div className="statlab-root">
        <div className="container">
          <div className="header">
            <div>
              <div className="brand-mark">stat<em>lab</em></div>
              <div className="brand-sub">automated statistical analysis · v2.0</div>
            </div>
            <div className="header-meta">
              <div>{new Date().toLocaleDateString("pt-BR", { day: "2-digit", month: "short", year: "numeric" })}</div>
              <div>powered by claude</div>
            </div>
          </div>

          {step === 0 && (
            <div className="mode-tabs">
              <button className={`mode-tab ${mode === "guided" ? "active" : ""}`} onClick={() => switchMode("guided")}>
                <FlaskConical size={16} /> Tenho um teste em mente
              </button>
              <button className={`mode-tab ${mode === "ai" ? "active" : ""}`} onClick={() => switchMode("ai")}>
                <Brain size={16} /> Não sei qual teste usar
              </button>
            </div>
          )}

          <div className="stepper">
            {stepsForMode.map((label, i) => (
              <div key={i} className={`step ${step === i ? "active" : ""} ${step > i ? "done" : ""}`}>
                <div className="step-num">{step > i ? "✓" : i + 1}</div>
                <div className="step-label">{label}</div>
              </div>
            ))}
          </div>

          {error && (
            <div className="alert">
              <AlertCircle size={18} style={{ flexShrink: 0, marginTop: 1 }} />
              <div>{error}</div>
            </div>
          )}

          {mode === "guided" && (
            <>
              {step === 0 && (
                <div className="card">
                  <div className="card-title">Escolha o <em>protocolo</em> estatístico</div>
                  <div className="card-desc">
                    Selecione o teste. Cada um tem um <strong style={{ color: 'var(--accent)' }}>template Excel pronto</strong> — clique em "Baixar template" no card.
                  </div>
                  <div className="test-grid">
                    {TESTS.map(t => {
                      const Icon = t.icon;
                      return (
                        <div key={t.id} className={`test-card ${selectedTest === t.id ? "selected" : ""}`} onClick={() => setSelectedTest(t.id)}>
                          <Icon size={22} className="test-card-icon" />
                          <div className="test-card-name">{t.name}</div>
                          <div className="test-card-desc">{t.desc}</div>
                          <div className="test-card-meta">{t.family}</div>
                          <button className="test-card-template-btn" onClick={(e) => { e.stopPropagation(); downloadTemplate(t); }}>
                            <Download size={11} /> Baixar template .xlsx
                          </button>
                        </div>
                      );
                    })}
                  </div>
                  <div className="actions">
                    <div />
                    <button className="btn" disabled={!selectedTest} onClick={() => setStep(1)}>
                      Próximo: enviar dados <ChevronRight size={16} />
                    </button>
                  </div>
                </div>
              )}

              {step === 1 && (
                <div className="card">
                  <div className="card-title">Envie sua <em>planilha</em></div>
                  <div className="card-desc">Use o template baixado ou seu próprio .xlsx/.csv.</div>
                  {currentTest && !file && (
                    <div className="info-alert">
                      <Lightbulb size={18} style={{ flexShrink: 0, marginTop: 1, color: 'var(--accent-cool)' }} />
                      <div>
                        Selecionou <strong style={{ color: 'var(--accent-cool)' }}>{currentTest.name}</strong>.{" "}
                        <a href="#" onClick={(e) => { e.preventDefault(); downloadTemplate(currentTest); }} style={{ color: 'var(--accent)', textDecoration: 'underline' }}>Baixar template</a>.
                      </div>
                    </div>
                  )}
                  {!file ? (
                    <div className={`dropzone ${drag ? "drag" : ""}`} onClick={() => fileRef.current.click()} onDragOver={(e) => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)} onDrop={handleDrop}>
                      <Upload size={40} className="dropzone-icon" strokeWidth={1.5} />
                      <div className="dropzone-text">Arraste ou clique</div>
                      <div className="dropzone-hint">.xlsx · .xls · .csv</div>
                      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => handleFile(e.target.files[0])} />
                    </div>
                  ) : (
                    <>
                      <div className="file-info">
                        <FileSpreadsheet size={28} className="file-info-icon" strokeWidth={1.5} />
                        <div style={{ flex: 1 }}>
                          <div className="file-info-name">{file.name}</div>
                          <div className="file-info-meta">{data?.length || 0} linhas · {columns.length} colunas</div>
                        </div>
                        <button className="btn ghost" onClick={() => { setFile(null); setData(null); setColumns([]); }}>Trocar</button>
                      </div>
                      <div className="section-h">Prévia</div>
                      <div className="table-wrap">
                        <table className="preview">
                          <thead>
                            <tr>{columns.map(c => <th key={c}>{c}<div style={{ fontSize: 9, opacity: 0.6, marginTop: 2 }}>{columnTypes[c]}</div></th>)}</tr>
                          </thead>
                          <tbody>
                            {data.slice(0, 5).map((row, i) => (
                              <tr key={i}>{columns.map(c => <td key={c}>{row[c] === null ? "—" : String(row[c])}</td>)}</tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </>
                  )}
                  <div className="actions">
                    <button className="btn ghost" onClick={() => setStep(0)}><ChevronLeft size={16} /> Voltar</button>
                    <button className="btn" disabled={!data} onClick={() => setStep(2)}>Próximo: mapear <ChevronRight size={16} /></button>
                  </div>
                </div>
              )}
            </>
          )}

          {mode === "ai" && (
            <>
              {step === 0 && (
                <div className="card">
                  <div className="card-title">Envie seus <em>dados</em></div>
                  <div className="card-desc">Modo IA: você não precisa saber qual teste. Suba o arquivo, descreva sua pergunta, e a IA recomenda.</div>
                  {!file ? (
                    <div className={`dropzone ${drag ? "drag" : ""}`} onClick={() => fileRef.current.click()} onDragOver={(e) => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)} onDrop={handleDrop}>
                      <Upload size={40} className="dropzone-icon" strokeWidth={1.5} />
                      <div className="dropzone-text">Arraste ou clique</div>
                      <div className="dropzone-hint">.xlsx · .xls · .csv</div>
                      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => handleFile(e.target.files[0])} />
                    </div>
                  ) : (
                    <>
                      <div className="file-info">
                        <FileSpreadsheet size={28} className="file-info-icon" strokeWidth={1.5} />
                        <div style={{ flex: 1 }}>
                          <div className="file-info-name">{file.name}</div>
                          <div className="file-info-meta">{data?.length || 0} linhas · {columns.length} colunas</div>
                        </div>
                        <button className="btn ghost" onClick={() => { setFile(null); setData(null); setColumns([]); }}>Trocar</button>
                      </div>
                      <div className="section-h">Prévia</div>
                      <div className="table-wrap">
                        <table className="preview">
                          <thead>
                            <tr>{columns.map(c => <th key={c}>{c}<div style={{ fontSize: 9, opacity: 0.6, marginTop: 2 }}>{columnTypes[c]}</div></th>)}</tr>
                          </thead>
                          <tbody>
                            {data.slice(0, 5).map((row, i) => (
                              <tr key={i}>{columns.map(c => <td key={c}>{row[c] === null ? "—" : String(row[c])}</td>)}</tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </>
                  )}
                  <div className="actions">
                    <div />
                    <button className="btn purple" disabled={!data} onClick={() => setStep(1)}>Próximo: descrever pergunta <ChevronRight size={16} /></button>
                  </div>
                </div>
              )}

              {step === 1 && (
                <div className="ai-recommend-card">
                  <div className="ai-badge"><Brain size={11} /> assistente ia</div>
                  <div className="card-title">Qual é sua <em>pergunta</em>?</div>
                  <div className="card-desc">Ex: "Quero saber se o medicamento X reduz a pressão arterial vs placebo" ou "Há relação entre sono e desempenho?"</div>
                  <textarea placeholder="Descreva sua pergunta..." value={researchQuestion} onChange={(e) => setResearchQuestion(e.target.value)} />

                  {aiRecommendation && (
                    <div className="recommend-result">
                      <div className="ai-badge"><CheckCircle2 size={11} /> recomendação</div>
                      <div className="recommend-test-name">{aiRecommendation.recommended_test_name}</div>
                      <div className="recommend-rationale" dangerouslySetInnerHTML={{ __html: (aiRecommendation.rationale || "").replace(/\*\*(.+?)\*\*/g, "<strong style='color:var(--accent)'>$1</strong>") }} />
                      <div className="section-h" style={{ margin: "20px 0 8px" }}>Mapeamento sugerido</div>
                      <div className="recommend-mapping">
                        {Object.entries(aiRecommendation.mapping || {}).map(([k, v]) => {
                          const test = TESTS.find(t => t.id === aiRecommendation.recommended_test_id);
                          const inp = test?.inputs.find(i => i.key === k);
                          return (
                            <div className="recommend-mapping-row" key={k}>
                              <div className="mk">{inp?.label || k}</div>
                              <div className="mv">{Array.isArray(v) ? v.join(", ") : v}</div>
                            </div>
                          );
                        })}
                      </div>
                      {aiRecommendation.alternatives && aiRecommendation.alternatives.length > 0 && (
                        <div className="recommend-alts">
                          <div className="recommend-alts-title">Alternativas a considerar</div>
                          {aiRecommendation.alternatives.map((alt, i) => (
                            <div className="alt-item" key={i}><strong>{alt.test_name}</strong> — {alt.when}</div>
                          ))}
                        </div>
                      )}
                      {aiRecommendation.data_quality_notes && aiRecommendation.data_quality_notes.trim() && (
                        <div className="info-alert" style={{ marginTop: 16, marginBottom: 0 }}>
                          <AlertCircle size={16} style={{ flexShrink: 0, marginTop: 1, color: 'var(--accent-warm)' }} />
                          <div><strong style={{ color: 'var(--accent-warm)' }}>Observação:</strong> {aiRecommendation.data_quality_notes}</div>
                        </div>
                      )}
                    </div>
                  )}

                  <div className="actions">
                    <button className="btn ghost" onClick={() => setStep(0)}><ChevronLeft size={16} /> Voltar</button>
                    {!aiRecommendation ? (
                      <button className="btn purple" disabled={!researchQuestion.trim() || loadingRec} onClick={getAiRecommendation}>
                        {loadingRec ? (<><Loader2 size={16} className="spin-icon" /> Analisando...</>) : (<><Brain size={16} /> Pedir recomendação</>)}
                      </button>
                    ) : (
                      <div style={{ display: 'flex', gap: 8 }}>
                        <button className="btn ghost" onClick={() => setAiRecommendation(null)}>Pedir outra</button>
                        <button className="btn" onClick={acceptRecommendation}>Aceitar e continuar <ChevronRight size={16} /></button>
                      </div>
                    )}
                  </div>
                </div>
              )}
            </>
          )}

          {step === 2 && currentTest && (
            <div className="card">
              <div className="card-title">{mode === "ai" ? "Confirme o " : "Mapeie as "}<em>{mode === "ai" ? "mapeamento" : "variáveis"}</em></div>
              <div className="card-desc">
                Análise: <strong style={{ color: "var(--text)" }}>{currentTest.name}</strong> · {currentTest.desc}
                {mode === "ai" && aiRecommendation && " · Pré-preenchido pela IA — revise."}
              </div>
              <div>
                {currentTest.inputs.map(inp => (
                  <div key={inp.key} className="mapping-row">
                    <div>
                      <div className="mapping-label">{inp.label}</div>
                      <div className="mapping-hint">{inp.hint}</div>
                    </div>
                    <div>
                      {inp.type === "columns" ? (
                        <select multiple size={Math.min(6, columns.length)} value={mapping[inp.key] || []} onChange={(e) => setMap(inp.key, Array.from(e.target.selectedOptions).map(o => o.value))}>
                          {columns.map(c => <option key={c} value={c}>{c} ({columnTypes[c]})</option>)}
                        </select>
                      ) : (
                        <select value={mapping[inp.key] || ""} onChange={(e) => setMap(inp.key, e.target.value)}>
                          <option value="">— selecione —</option>
                          {columns.map(c => <option key={c} value={c}>{c} ({columnTypes[c]})</option>)}
                        </select>
                      )}
                    </div>
                  </div>
                ))}
              </div>
              <div className="actions">
                <button className="btn ghost" onClick={() => setStep(1)}><ChevronLeft size={16} /> Voltar</button>
                <button className="btn" disabled={!mappingComplete()} onClick={runAnalysis}><Sparkles size={16} /> Rodar análise</button>
              </div>
            </div>
          )}

          {step === 3 && (
            <div>
              {loading && (
                <div className="card">
                  <div className="loading-state">
                    <Loader2 size={48} className="spin-icon" strokeWidth={1.5} />
                    <div className="loading-state-text">Computando análise...</div>
                    <div className="loading-state-sub">pressupostos · modelo · gráficos</div>
                  </div>
                </div>
              )}
              {!loading && result && (
                <>
                  <div className="card">
                    <div className="results-header">
                      <div>
                        <div className="results-title">Resultado da <em>análise</em></div>
                        <div className="results-subtitle">{result.test_performed} · n = {data.length}</div>
                      </div>
                      <button className="btn ghost" onClick={reset}><ArrowLeft size={16} /> Nova análise</button>
                    </div>
                    {result.metrics && result.metrics.length > 0 && (
                      <div className="metric-grid">
                        {result.metrics.map((m, i) => (
                          <div className="metric" key={i}>
                            <div className="metric-label">{m.label}</div>
                            <div className={`metric-value ${m.significant === true ? "sig" : m.significant === false ? "notsig" : ""}`}>{m.value}</div>
                            {m.sub && <div className="metric-sub">{m.sub}</div>}
                          </div>
                        ))}
                      </div>
                    )}
                    {result.interpretation && (
                      <>
                        <div className="section-h">Interpretação</div>
                        <div className="interpretation" dangerouslySetInnerHTML={{ __html: result.interpretation.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>") }} />
                      </>
                    )}
                    {result.summary_table && result.summary_table.rows?.length > 0 && (
                      <>
                        <div className="section-h">Resumo descritivo</div>
                        <div className="table-wrap">
                          <table className="preview">
                            <thead><tr>{result.summary_table.headers.map((h, i) => <th key={i}>{h}</th>)}</tr></thead>
                            <tbody>
                              {result.summary_table.rows.map((row, i) => (
                                <tr key={i}>{row.map((cell, j) => <td key={j}>{cell}</td>)}</tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </>
                    )}
                    {result.assumptions && (
                      <>
                        <div className="section-h">Pressupostos</div>
                        <div className="detail-list">
                          {result.assumptions.checked?.map((c, i) => (
                            <div className="detail-row" key={i}>
                              <div className="k">Pressuposto {i + 1}</div>
                              <div className="v">{c}</div>
                            </div>
                          ))}
                          {result.assumptions.recommendation && (
                            <div className="detail-row">
                              <div className="k">Recomendação</div>
                              <div className="v" style={{ fontFamily: "var(--sans)" }}>{result.assumptions.recommendation}</div>
                            </div>
                          )}
                        </div>
                      </>
                    )}
                  </div>
                  {result.charts && result.charts.length > 0 && (
                    <div className="card">
                      <div className="section-h" style={{ margin: "0 0 16px" }}>Visualizações</div>
                      {result.charts.map((chart, i) => <ChartRenderer key={i} chart={chart} />)}
                    </div>
                  )}
                  {result.additional_notes && (
                    <div className="card">
                      <div className="section-h" style={{ margin: "0 0 12px" }}>Notas adicionais</div>
                      <div style={{ fontFamily: "var(--serif)", fontSize: 16, lineHeight: 1.7, color: "var(--text-dim)" }}>{result.additional_notes}</div>
                    </div>
                  )}
                </>
              )}
              {!loading && !result && error && (
                <div className="actions">
                  <button className="btn ghost" onClick={() => setStep(2)}><ChevronLeft size={16} /> Voltar</button>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </>
  );
}

function ChartRenderer({ chart }) {
  const accent = "#d4ff3a";
  const accentCool = "#6ab7ff";
  const accentWarm = "#ff8c42";
  const palette = [accent, accentCool, accentWarm, "#4ade80", "#ff5470", "#a78bfa"];

  return (
    <div className="chart-card">
      <div className="chart-title">{chart.title}</div>
      {chart.description && <div className="chart-desc">{chart.description}</div>}
      <ResponsiveContainer width="100%" height={280}>
        {chart.type === "scatter" ? (
          <ScatterChart>
            <CartesianGrid stroke="#262b38" strokeDasharray="3 3" />
            <XAxis type="number" dataKey={chart.xKey || "x"} stroke="#9098a8" fontSize={11} fontFamily="JetBrains Mono" />
            <YAxis type="number" dataKey={chart.yKey || "y"} stroke="#9098a8" fontSize={11} fontFamily="JetBrains Mono" />
            <Tooltip contentStyle={{ background: "#161922", border: "1px solid #353c4d", borderRadius: 3, fontSize: 12 }} />
            <Scatter data={chart.data} fill={accent} />
            {chart.extra?.trendline && (<Line data={chart.extra.trendline} dataKey="y" stroke={accentCool} strokeWidth={2} dot={false} type="linear" />)}
          </ScatterChart>
        ) : chart.type === "boxplot_summary" ? (
          <BarChart data={chart.data}>
            <CartesianGrid stroke="#262b38" strokeDasharray="3 3" />
            <XAxis dataKey={chart.xKey || "name"} stroke="#9098a8" fontSize={11} fontFamily="JetBrains Mono" />
            <YAxis stroke="#9098a8" fontSize={11} fontFamily="JetBrains Mono" />
            <Tooltip contentStyle={{ background: "#161922", border: "1px solid #353c4d", borderRadius: 3, fontSize: 12 }} />
            <Legend wrapperStyle={{ fontSize: 11, fontFamily: "JetBrains Mono" }} />
            <Bar dataKey="min" fill={accentCool} name="Mín" />
            <Bar dataKey="q1" fill="#4ade80" name="Q1" />
            <Bar dataKey="median" fill={accent} name="Mediana" />
            <Bar dataKey="q3" fill={accentWarm} name="Q3" />
            <Bar dataKey="max" fill="#ff5470" name="Máx" />
          </BarChart>
        ) : (
          <BarChart data={chart.data}>
            <CartesianGrid stroke="#262b38" strokeDasharray="3 3" />
            <XAxis dataKey={chart.xKey || "name"} stroke="#9098a8" fontSize={11} fontFamily="JetBrains Mono" />
            <YAxis stroke="#9098a8" fontSize={11} fontFamily="JetBrains Mono" />
            <Tooltip contentStyle={{ background: "#161922", border: "1px solid #353c4d", borderRadius: 3, fontSize: 12 }} />
            <Bar dataKey={chart.yKey || "value"} fill={accent}>
              {chart.data?.map((_, idx) => <Cell key={idx} fill={palette[idx % palette.length]} />)}
            </Bar>
          </BarChart>
        )}
      </ResponsiveContainer>
    </div>
  );
}
