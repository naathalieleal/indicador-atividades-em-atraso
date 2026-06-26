const fs   = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const chalk = require("chalk");

// ═══════════════════════════════════════════════════════════════
//  CONFIG
// ═══════════════════════════════════════════════════════════════
const CONFIG = {
  pastaRelatorios: "C:\\Users\\naath\\Documents\\Indicadores",

  colunas: {
    area:        "Área",
    modulo:      "Modulo",
    prazo:       "Prazo",
    diasAtraso:  "Dias Atraso",
    identificador: "Identificador",
    titulo:      "Título",
    acao:        "Ação",
  },

  percentualCritico: 15,
  atrasoGrave:       30,
};

// ═══════════════════════════════════════════════════════════════
//  CONSTANTES
// ═══════════════════════════════════════════════════════════════
const SEPARADOR       = "═".repeat(70);
const SEPARADOR_FINO  = "─".repeat(70);
const PASTA_HISTORICO = path.join(__dirname, "historico");

// ═══════════════════════════════════════════════════════════════
//  UTILITÁRIOS
// ═══════════════════════════════════════════════════════════════
function colorirPct(pct) {
  const n = Number(pct);
  if (n === 0)                        return chalk.green(`${pct}%`);
  if (n <= CONFIG.percentualCritico)  return chalk.yellow(`${pct}%`);
  return chalk.red(`${pct}%`);
}

function calcularPct(parcial, total) {
  if (!total) return "0.0";
  return ((parcial / total) * 100).toFixed(1);
}

function formatarPrazo(valor) {
  if (!valor) return "";
  if (typeof valor === "number") {
    const d = new Date(Math.round((valor - 25569) * 86400 * 1000));
    return d.toLocaleDateString("pt-BR");
  }
  return String(valor).split(" ")[0];
}

function agora() {
  return new Date().toLocaleString("pt-BR");
}

function carimboDeTempo() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}_${pad(d.getHours())}-${pad(d.getMinutes())}`;
}

// ═══════════════════════════════════════════════════════════════
//  BUSCA AUTOMÁTICA DO EXCEL MAIS RECENTE
// ═══════════════════════════════════════════════════════════════
function buscarArquivoMaisRecente(pasta) {
  if (!fs.existsSync(pasta)) {
    console.error(chalk.red(`\n❌ Pasta não encontrada: ${pasta}`));
    console.error(chalk.yellow("   Verifique CONFIG.pastaRelatorios e tente novamente.\n"));
    process.exit(1);
  }

  const arquivos = fs.readdirSync(pasta)
    .filter(f =>
      f.endsWith(".xlsx") &&
      !f.startsWith("~$") &&
      !f.startsWith(".")
    )
    .map(f => ({
      nome: f,
      caminho: path.join(pasta, f),
      modificado: fs.statSync(path.join(pasta, f)).mtimeMs,
    }))
    .sort((a, b) => b.modificado - a.modificado);

  if (arquivos.length === 0) {
    console.error(chalk.red("\n❌ Nenhum arquivo .xlsx encontrado na pasta configurada."));
    console.error(chalk.yellow(`   Pasta: ${pasta}\n`));
    process.exit(1);
  }

  return arquivos[0];
}

// ═══════════════════════════════════════════════════════════════
//  LEITURA DO EXCEL
// ═══════════════════════════════════════════════════════════════
function lerExcel(caminho) {
  const workbook = XLSX.readFile(caminho);
  const planilha  = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(planilha, { defval: null });
}

// ═══════════════════════════════════════════════════════════════
//  ANÁLISE DOS DADOS
// ═══════════════════════════════════════════════════════════════
function analisarDados(dados) {
  const col   = CONFIG.colunas;
  const areas = {};

  for (const linha of dados) {
    const area   = linha[col.area];
    const modulo = linha[col.modulo];
    if (!area || !modulo) continue;

    if (!areas[area]) {
      areas[area] = { total: 0, emAtraso: 0, atrasoGrave: [], porTipo: {} };
    }

    const a = areas[area];
    a.total++;

    if (!a.porTipo[modulo]) a.porTipo[modulo] = { total: 0, emAtraso: 0 };
    a.porTipo[modulo].total++;

    const dias = Number(linha[col.diasAtraso]);
    if (!isNaN(dias) && dias > 0) {
      a.emAtraso++;
      a.porTipo[modulo].emAtraso++;

      if (dias > CONFIG.atrasoGrave) {
        a.atrasoGrave.push({
          descricao: modulo === "Documento"
            ? `${linha[col.identificador] || ""} — ${linha[col.titulo] || "(sem descrição)"}`
            : `${linha[col.identificador] || ""} — ${linha[col.acao] || ""} — ${linha[col.titulo] || "(sem descrição)"}`,
          tipo:  modulo,
          prazo: formatarPrazo(linha[col.prazo]),
          dias,
        });
      }
    }
  }

  // Ordena áreas por percentual decrescente
  return Object.entries(areas)
    .map(([nome, d]) => ({ nome, ...d, pct: calcularPct(d.emAtraso, d.total) }))
    .sort((a, b) => Number(b.pct) - Number(a.pct));
}

// ═══════════════════════════════════════════════════════════════
//  IMPRESSÃO
// ═══════════════════════════════════════════════════════════════
const linhas = [];
function print(texto = "") {
  console.log(texto);
  linhas.push(texto);
}

function imprimirCabecalho(arquivo) {
  print(SEPARADOR);
  print(`  RELATÓRIO DE ATIVIDADES EM ATRASO`);
  print(`  Gerado em: ${agora()}`);
  print(`  Arquivo:   ${arquivo}`);
  print(SEPARADOR);
}

function imprimirResumo(areas, dados) {
  const totalAtividades = dados.length;
  const totalEmAtraso   = areas.reduce((s, a) => s + a.emAtraso, 0);
  const totalGraves     = areas.reduce((s, a) => s + a.atrasoGrave.length, 0);
  const pctGeral        = calcularPct(totalEmAtraso, totalAtividades);

  print();
  print(chalk.bold("  RESUMO GERAL"));
  print(SEPARADOR_FINO);
  print(`  Áreas identificadas : ${areas.length}`);
  print(`  Total de atividades : ${totalAtividades}`);
  print(`  Em atraso           : ${totalEmAtraso}`);
  print(`  Percentual geral    : ${colorirPct(pctGeral)}`);
  print(`  Atraso > ${CONFIG.atrasoGrave} dias   : ${totalGraves > 0 ? chalk.red(totalGraves) : chalk.green(totalGraves)}`);
  print(SEPARADOR_FINO);
}

function imprimirDetalhamento(areas) {
  print();
  print(chalk.bold("  DETALHAMENTO POR ÁREA"));

  const criticas  = areas.filter(a => Number(a.pct) > CONFIG.percentualCritico);
  const normais   = areas.filter(a => Number(a.pct) <= CONFIG.percentualCritico);

  if (criticas.length) {
    print();
    print(chalk.red(`  ⚠  ÁREAS CRÍTICAS (acima de ${CONFIG.percentualCritico}%)`));
  }

  for (const a of [...criticas, ...normais]) {
    const isCritica = Number(a.pct) > CONFIG.percentualCritico;
    print();
    print(SEPARADOR_FINO);

    const titulo = `  ${isCritica ? "🔴" : "📁"} ${a.nome}`;
    print(isCritica ? chalk.red(chalk.bold(titulo)) : chalk.bold(titulo));
    print(`     Total: ${a.total}  |  Em atraso: ${a.emAtraso}  |  Percentual: ${colorirPct(a.pct)}`);

    for (const [tipo, t] of Object.entries(a.porTipo)) {
      const pctTipo = calcularPct(t.emAtraso, t.total);
      print(`       • ${tipo}: ${t.emAtraso}/${t.total} em atraso (${colorirPct(pctTipo)})`);
    }
  }

  print(SEPARADOR_FINO);
}

function imprimirAtividadesGraves(areas) {
  const graves = areas
    .flatMap(a => a.atrasoGrave.map(g => ({ ...g, area: a.nome })))
    .sort((a, b) => b.dias - a.dias);

  if (!graves.length) {
    print();
    print(chalk.green(`  ✅ Nenhuma atividade com atraso superior a ${CONFIG.atrasoGrave} dias.`));
    return;
  }

  print();
  print(chalk.red(chalk.bold(`  ATIVIDADES COM ATRASO > ${CONFIG.atrasoGrave} DIAS (${graves.length} ocorrência(s))`)));
  print(SEPARADOR_FINO);

  let areaAtual = "";
  for (const g of graves) {
    if (g.area !== areaAtual) {
      print();
      print(chalk.red(`  📁 ${g.area}`));
      areaAtual = g.area;
    }
    print(chalk.red(`     [${g.tipo}] ${g.descricao}`));
    print(chalk.red(`     Prazo: ${g.prazo}  |  Atraso: ${g.dias} dias`));
  }

  print(SEPARADOR_FINO);
}

// ═══════════════════════════════════════════════════════════════
//  SALVAR TXT
// ═══════════════════════════════════════════════════════════════
function salvarTxt() {
  // Remove códigos ANSI de cor para o arquivo de texto
  const semCores = linhas.map(l => l.replace(/\x1B\[[0-9;]*m/g, "")).join("\n");

  // relatorio.txt na raiz
  fs.writeFileSync(path.join(__dirname, "relatorio.txt"), semCores, "utf8");

  // historico/Relatorio_AAAA-MM-DD_HH-mm.txt
  if (!fs.existsSync(PASTA_HISTORICO)) fs.mkdirSync(PASTA_HISTORICO);
  const nomeHistorico = path.join(PASTA_HISTORICO, `Relatorio_${carimboDeTempo()}.txt`);
  fs.writeFileSync(nomeHistorico, semCores, "utf8");

  return nomeHistorico;
}

// ═══════════════════════════════════════════════════════════════
//  MAIN
// ═══════════════════════════════════════════════════════════════
function main() {
  const arquivo = buscarArquivoMaisRecente(CONFIG.pastaRelatorios);

  console.log(chalk.cyan(`\n📂 Arquivo encontrado: ${arquivo.nome}`));
  console.log(chalk.gray(`   Caminho: ${arquivo.caminho}\n`));

  const dados = lerExcel(arquivo.caminho);
  console.log(chalk.gray(`   ${dados.length} linhas carregadas.\n`));

  const areas = analisarDados(dados);

  imprimirCabecalho(arquivo.nome);
  imprimirResumo(areas, dados);
  imprimirDetalhamento(areas);
  imprimirAtividadesGraves(areas);

  print();
  print(SEPARADOR);

  const nomeHistorico = salvarTxt();
  console.log(chalk.cyan(`\n✅ relatorio.txt salvo na pasta do script.`));
  console.log(chalk.gray(`   Histórico: ${nomeHistorico}\n`));
}

main();
