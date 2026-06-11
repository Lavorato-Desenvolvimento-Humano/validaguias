# Valida Guias

Ferramenta web simples para cruzar a **planilha de agendamentos** com a **planilha de guias assinadas/enviadas para BM** e gerar uma planilha consolidada indicando, para cada agendamento, a situação da guia.

Tudo roda **100% no navegador** — nenhum dado de paciente é enviado para servidores.

---

## Como usar

1. Abra o arquivo [`index.html`](index.html) no navegador (duplo clique).
2. Selecione (ou arraste) a **Planilha de Agendamentos**.
3. Selecione (ou arraste) a **Planilha de Guias Assinadas**.
4. Clique em **"Cruzar dados e baixar planilha"** — o botão habilita quando os dois arquivos estão carregados.
5. O arquivo `guias_cruzadas.xlsx` é baixado automaticamente.

Formatos aceitos: `.xlsx`, `.xls`, `.csv`.

---

## O que o cruzamento faz

Para cada linha da planilha de agendamentos, o sistema procura uma guia correspondente e preenche a coluna **"Situação da Guia"**.

**Chave de cruzamento:**

| Campo | Regra |
|-------|-------|
| Nome do paciente | normalizado (maiúsculas, sem acentos, espaços colapsados) |
| Competência | mês/ano do atendimento × mês/ano da guia |
| Convênio | igual **ou** um contém o outro (ex.: `CBMDF ABA` ≈ `CBMDF`) |

- **Encontrou guia:** a coluna recebe o **status real** da guia (`ASSINADO` ou `ENVIADO A BM`), para indicar onde ela está.
- **Não encontrou:** a coluna fica vazia (pendente).

**Ordenação do resultado:** Nome do paciente (A→Z) → Convênio (A→Z) → Hora do atendimento (menor→maior).

**Colunas removidas do resultado:** Terapeuta, Especialidade e Obs.

---

## Robustez com planilhas "cruas"

A ferramenta lida com arquivos exportados direto dos sistemas, sem ajuste manual:

- **Detecção automática da linha de cabeçalho** — pula linhas institucionais (nome da clínica, CNPJ, endereço, título) que aparecem antes da tabela.
- **Reconhecimento flexível de colunas** — entende variações de nome de coluna:
  - Paciente: `Assistido`, `Paciente`, `Nome`, `Cliente`
  - Convênio: `Convênio`, `Plano`
  - Hora: `Início`, `Hora`, `Horário`, `Entrada`
  - Competência: coluna única **ou** colunas separadas `mês` + `ano`
- **Datas** — entende texto (`12/06/2026`) e o formato nativo do Excel.

---

## Estrutura do projeto

```
valida_guias/
├── index.html      # estrutura da página (uploads + botão)
├── style.css       # estilos (dropzones, cartões, layout sóbrio)
├── script.js       # leitura das planilhas, cruzamento e geração do .xlsx
├── README.md
├── docs/           # documentação de concepção (arquitetura, BPMN, casos de uso, etc.)
└── prototipos/     # protótipos de tela da concepção inicial
```

> As pastas [`docs/`](docs/) e [`prototipos/`](prototipos/) descrevem a **concepção inicial** de um sistema corporativo completo (React + NestJS + PostgreSQL). A implementação atual foi propositalmente simplificada para uma ferramenta de arquivo único, sem backend. Elas ficam como referência de projeto.

---

## Tecnologia

- HTML + CSS + JavaScript puro (sem build, sem framework).
- [SheetJS (xlsx)](https://sheetjs.com) via CDN para ler e gerar planilhas.

Não há dependências instaladas nem servidor: basta abrir o `index.html`.

---

## Glossário

| Termo | Significado |
|-------|-------------|
| **Guia** | Autorização de procedimento junto ao convênio |
| **BM** | Boletim de Medição / faturamento |
| **Competência** | Mês/ano de referência do atendimento |
| **Situação da Guia** | Status da guia encontrada (`ASSINADO` / `ENVIADO A BM`) ou vazio |
