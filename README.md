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

Para cada linha da planilha de agendamentos, o sistema procura a guia correspondente e preenche o resultado.

**Chave de cruzamento:**

| Campo | Regra |
|-------|-------|
| Nome do paciente | normalizado (maiúsculas, sem acentos, espaços colapsados) |
| Competência | mês/ano do atendimento × mês/ano da guia |
| Convênio | igual **ou** um contém o outro (ex.: `CBMDF ABA` ≈ `CBMDF`) |
| **Procedimento** | categoria do atendimento × categoria da guia (ver abaixo) |

**Resultado por linha:**

- **Guia do mesmo procedimento encontrada:** recebe o **status real** da guia (`ASSINADO` ou `ENVIADO A BM`).
- **Guia encontrada, mas de outro procedimento:** também recebe o status, porém é **sinalizada** na coluna `Obs. Validação` (`Guia de outro procedimento`) para conferência manual.
- **Nenhuma guia:** fica vazia (pendente).

**Ordenação do resultado:** Hora do atendimento (menor→maior) → Nome do paciente (A→Z) → Convênio (A→Z).

**Colunas removidas do resultado:** Especialidade e Obs.

### Validação por procedimento

Como os dois sistemas usam nomes diferentes para o mesmo procedimento, cada serviço (da agenda e da guia) é reduzido a uma **categoria canônica** antes de comparar. Ex.: `SESSÃO DE PSICOLOGIA INDIVIDUAL` e `CONSULTA… COM PSICOLOGO` → `PSICOTERAPIA`; `AVALIACAO NEUROPSICOLOGIA` → `AVALIACAO NEUROPSICOLOGICA`.

Regras especiais:

- **Terapia ABA:** qualquer serviço "COM TERAPIA ABA" (ou pacote ABA) é tratado como categoria **TERAPIA ABA** e validado por uma guia de ABA — não pela especialidade base (fono, psicologia, etc.).
- **Convênios com pacote de horas (Saúde Caixa):** psicoterapia casa com guia de psicoterapia e avaliação neuro com a dela; **todos os demais** procedimentos (fono, psicopedagogia, TO, fisio…) são cobertos por uma guia de **PACOTE**. *(Não há contagem de horas 2/4/6h — apenas a presença do pacote.)* A lista desses convênios fica em `CONVENIOS_PACOTE`, no `script.js`.

---

## Robustez com planilhas "cruas"

A ferramenta lida com arquivos exportados direto dos sistemas, sem ajuste manual:

- **Detecção automática da linha de cabeçalho** — pula linhas institucionais (nome da clínica, CNPJ, endereço, título) que aparecem antes da tabela.
- **Reconhecimento flexível de colunas** — entende variações de nome de coluna:
  - Paciente: `Assistido`, `Paciente`, `Nome`, `Cliente`
  - Convênio: `Convênio`, `Plano`
  - Hora: `Início`, `Hora`, `Horário`, `Entrada`
  - Procedimento: `Serviço`, `Procedimento`, `Especialidade`
  - Competência: coluna única **ou** colunas separadas `mês` + `ano`
- **Datas** — entende texto (`12/06/2026`) e o formato nativo do Excel.
- **Largura das colunas** ajustada automaticamente no arquivo gerado.

> Se a planilha de guias **não** tiver coluna de procedimento, a validação por procedimento é desativada e o sistema volta a casar apenas por nome + competência + convênio.

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

## Colunas do arquivo gerado (`guias_cruzadas.xlsx`)

> Assistido · Data · Início · Término · Convênio · **guias** · Terapeuta · Serviço · Status · **Situação da Guia** · **Obs. Validação**

- **guias** — número da guia encontrada (vem da planilha de guias).
- **Situação da Guia** — status da guia (`ASSINADO` / `ENVIADO A BM`) ou vazio se pendente.
- **Obs. Validação** — `Guia de outro procedimento` quando a guia encontrada não é do procedimento agendado; vazio quando o match é limpo.

---

## Glossário

| Termo | Significado |
|-------|-------------|
| **Guia** | Autorização de procedimento junto ao convênio |
| **BM** | Boletim de Medição / faturamento |
| **Competência** | Mês/ano de referência do atendimento |
| **Categoria** | Procedimento canônico usado para validar atendimento × guia (ex.: `PSICOTERAPIA`, `TERAPIA ABA`, `PACOTE`) |
| **Situação da Guia** | Status da guia encontrada (`ASSINADO` / `ENVIADO A BM`) ou vazio |
