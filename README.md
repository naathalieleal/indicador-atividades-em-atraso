# Relatório de Atividades em Atraso

### Problema

O indicador de atividades em atraso era obtido por meio de uma planilha Excel exportada do sistema.

A análise exigia:

   • filtrar dados;

   • consolidar por área;

   • calcular percentuais;

   • identificar atividades críticas;

   • montar manualmente um relatório para os gestores.

O processo era repetitivo e consumia tempo.



### Solução

Foi desenvolvida uma aplicação em Node.js que automatiza toda a consolidação das informações.

A ferramenta:

   • lê automaticamente a planilha;

   • calcula os indicadores;

   • identifica áreas críticas;

   • lista atividades acima de 30 dias;

   • gera um relatório executivo para envio aos gestores.


## Operação

```
Salve o arquivo na pasta configurada em `CONFIG.pastaRelatorios`

O script escolhe automaticamente o arquivo modificado mais recentemente — não é necessário renomear ou mover o arquivo.
```

Script Node.js lê automaticamente a planilha Excel mais recente de uma pasta,
analisa atividades em atraso por área e exibe o resultado colorido no terminal.
O relatório é salvo em `.txt` para ser copiado e colado manualmente no Outlook.

---

## Instalação

### 1. Instale o Node.js
Baixe em https://nodejs.org (versão LTS recomendada).

### 2. Instale as dependências
Abra o terminal na pasta do script e execute:
```bash
npm install
```

---

## Como configurar

Abra o arquivo `relatorio.js` e edite o bloco `CONFIG` no topo:

```js
const CONFIG = {
  pastaRelatorios: "C:\\Users\\naath\\Documents\\Indicadores", // pasta com os .xlsx
  colunas: {
    area:          "Área",        // cabeçalho da coluna de área
    modulo:        "Modulo",      // cabeçalho da coluna de módulo/tipo
    prazo:         "Prazo",       // cabeçalho da coluna de prazo
    diasAtraso:    "Dias Atraso", // cabeçalho da coluna de dias em atraso
    identificador: "Identificador",
    titulo:        "Título",
    acao:          "Ação",
  },
  percentualCritico: 15, // % a partir do qual a área é marcada como crítica (🔴)
  atrasoGrave:       30, // dias a partir dos quais a atividade aparece no bloco de graves
};
```

> Os nomes das colunas devem ser **idênticos** aos cabeçalhos do Excel (incluindo acentos).

---

## Como executar

No terminal integrado do VSCode, dentro da pasta do script:

```bash
node relatorio.js
```

Ou usando o atalho definido no package.json:

```bash
npm start
```

## Fluxo de utilização diária

```
1. Exportar o Excel atualizado do sistema
   └─ Salvar na pasta: C:\Users\naath\Documents\Indicadores

2. Abrir o terminal no VSCode (Ctrl + `)

3. Executar o script
   └─ node relatorio.js

4. O relatório aparece no terminal com cores:
   🟢 Verde  → 0% de atraso
   🟡 Amarelo → até 15%
   🔴 Vermelho → acima de 15%

5. Abrir o arquivo relatorio.txt gerado na pasta do script

6. Selecionar tudo (Ctrl+A), copiar (Ctrl+C)

7. Colar no corpo do e-mail no Outlook (Ctrl+V)
```

---

## Arquivos gerados

| Arquivo | Descrição |
|---|---|
| `relatorio.txt` | Última execução (sobrescrito a cada rodada) |
| `historico/Relatorio_AAAA-MM-DD_HH-mm.txt` | Cópia com carimbo de data e hora |

---

## Solução de problemas

**"Pasta não encontrada"**
→ Verifique o caminho em `CONFIG.pastaRelatorios`. Use barras duplas `\\` no Windows.

**"Nenhum arquivo .xlsx encontrado"**
→ Confirme que o Excel exportado está na pasta correta e não está aberto (arquivos abertos geram `~$nome.xlsx` que são ignorados).

**Números incorretos / coluna não encontrada**
→ Abra o Excel e confira os cabeçalhos exatos. Acentos e maiúsculas importam.
Atualize os nomes em `CONFIG.colunas`.

**"Cannot find module 'xlsx'" ou 'chalk'**
→ Execute `npm install` novamente na pasta do script.
