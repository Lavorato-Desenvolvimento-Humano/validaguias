(function () {
        "use strict";

        // ==================================================================
        // TABELA DE-PARA DE PROCEDIMENTOS
        // ------------------------------------------------------------------
        // Cada par é [ nome do serviço como vem nas planilhas , categoria ].
        // A categoria DEVE ser uma destas: TERAPIA ABA, PSICOTERAPIA,
        // FONOAUDIOLOGIA, TERAPIA OCUPACIONAL, PSICOPEDAGOGIA, PSICOMOTRICIDADE,
        // MUSICOTERAPIA, FISIOTERAPIA, NUTRICAO, AVALIACAO NEUROPSICOLOGICA, PACOTE.
        // A busca ignora maiúsculas/minúsculas e acentos.
        // Nomes que NÃO estiverem aqui caem no reconhecimento por palavra-chave
        // e, por fim, na coluna "Especialidade" da agenda.
        // ==================================================================
        var DE_PARA_PROCEDIMENTOS = [
          ["PSICOTERAPIA INDIVIDUAL INFANTIL(SESSÕES DE PSICOTERAPIA ABA)", "TERAPIA ABA"],
          ["SESSÃO DE PSICOPEDAGOGIA", "PSICOPEDAGOGIA"],
          ["CONSULTA/SESSÃO INDIVIDUAL AMBULATORIAL COM PSICOLOGO", "PSICOTERAPIA"],
          ["SESSÃO DE PSICOLOGIA INDIVIDUAL", "PSICOTERAPIA"],
          ["SESSÃO DE PSICOLOGIA COM TERAPIA ABA", "TERAPIA ABA"],
          ["PACOTE - ATENDIMENTO INTEGRADO 2H - PACIENTES TGD/TEA", "PACOTE"],
          ["SESSÃO DE TERAPIA OCUPACIONAL COM TERAPIA ABA", "TERAPIA ABA"],
          ["SESSÃO DE FONOAUDIOLOGIA", "FONOAUDIOLOGIA"],
          ["SESSÃO DE TERAPIA OCUPACIONAL", "TERAPIA OCUPACIONAL"],
          ["AVALIAÇÃO NEUROPSICOLÓGIA", "AVALIACAO NEUROPSICOLOGICA"],
          ["TERAPIA ABA - SESSAO", "TERAPIA ABA"],
          ["SESSÃO DE TERAPIA OCUPACIONAL - PELO MÉTODO ABA", "TERAPIA ABA"],
          ["TERAPIA COM ABA (ANÁLISE AMBULATORIAL DE COMPORTAMENTO APLICADA) - PACOTE", "TERAPIA ABA"],
          ["SESSÃO DE PSICOPEDAGOGIA - PELO MÉTODO ABA", "TERAPIA ABA"],
          ["SESSÃO DE FISIOTERAPIA", "FISIOTERAPIA"],
          ["SESSÃO DE MUSICOTERAPIA COM TERAPIA ABA", "TERAPIA ABA"],
          ["SESSÃO DE PSICOMOTRICIDADE", "PSICOMOTRICIDADE"],
          ["PR P AVAL NEUROPSICOLOGA SESSÕES SUBSEQUENTES", "AVALIACAO NEUROPSICOLOGICA"],
          ["SESSÃO DE TERAPIA COMPORTAMENTAL APLICADA", "TERAPIA ABA"],
          ["PSICOTERAPIA INDIVIDUAL", "PSICOTERAPIA"],
          ["TERAPIA OCUPACIONAL - AVALIACAO DOS COMPONENTES DE DESEMPENHO OCUPACIONAL - SESSOES", "TERAPIA ABA"],
          ["CONSULTA/SESSÃO INDIVIDUAL AMBULATORIAL DE FONOAUDIOLOGIA", "FONOAUDIOLOGIA"],
          ["TRATAMENTO TEA E OUTROS TRANSTORNOS GLOBAIS DO DESENVOLVIMENTO - POR DIA COM PSICOLOGO", "PSICOTERAPIA"],
          ["SESSÃO DE FONOAUDIOLOGIA COM TERAPIA ABA", "TERAPIA ABA"],
          ["SESSÃO DE MUSICOTERAPIA", "MUSICOTERAPIA"],
          ["PSICOPEDAGOGIA INDIVIDUAL", "PSICOPEDAGOGIA"],
          ["SESSOES DE FONOTERAPIA/FONOAUDIOLOGIA", "FONOAUDIOLOGIA"],
          ["SESSÃO DE NUTRIÇÃO", "NUTRICAO"],
          ["TERAPIA ABA - ATENDIMENTO SEMANAL CONFORME ESPECIFICACAO MEDICA", "TERAPIA ABA"],
          ["SESSAO DE PSICOTERAPIA INDIVIDUAL POR PSICOLOGO", "PSICOTERAPIA"],
          // "TRATAMENTO TEA ... POR DIA" (sem psicólogo) fica de fora de propósito:
          // cai no fallback da coluna Especialidade da agenda.
          ["SESSÃO DE PSICOTERAPIA", "PSICOTERAPIA"],
          ["RETARDO DO DESENVOLVIMENTO PSICOMOTOR - TRATAMENTO GLOBAL (FISIOTERAPIA)", "FISIOTERAPIA"],
          ["SESSAO DE ORIENTACAO/ACOLHIMENTO AO FAMILIAR TGD/TEA", "PSICOTERAPIA"],
          ["CONSULTA COM PSICOLOGIA - AVALIAÇÃO", "PSICOTERAPIA"],
          ["FONOAUDIOLOGIA", "FONOAUDIOLOGIA"],
          ["ACOMPANHANTE TERAPEUTICO", "PSICOTERAPIA"],
        ];

        var arqAgenda = document.getElementById("arqAgenda");
        var arqGuias = document.getElementById("arqGuias");
        var btn = document.getElementById("cruzar");
        var status = document.getElementById("status");
        var resumo = document.getElementById("resumo");

        // --- helpers de normalização ---------------------------------------

        function normalizar(v) {
          if (v === null || v === undefined) return "";
          return String(v)
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "") // remove acentos (marcas combinantes U+0300–U+036F)
            .toUpperCase()
            .replace(/\s+/g, " ")
            .trim();
        }

        // preposições/conectivos que não identificam o paciente
        var PREPOSICOES = { DE: 1, DA: 1, DO: 1, DAS: 1, DOS: 1, E: 1 };

        // tokens significativos do nome (sem acentos, sem preposições)
        function tokensNome(nome) {
          return normalizar(nome)
            .split(" ")
            .filter(function (t) {
              return t && !PREPOSICOES[t];
            });
        }

        // nomes casam se o conjunto de tokens de um estiver CONTIDO no outro
        // (ignora ordem e preposições; tolera nome do meio faltando).
        // Exige >= 2 tokens no nome menor, para não casar só pelo primeiro nome.
        // Trata iniciais: uma palavra de 1 letra casa com uma palavra do outro
        // nome que comece com essa letra (ex.: "R" ~ "ROBERTO", "G" ~ "GRACIANO").
        function tokenCasa(t, maior, set) {
          if (set[t]) return true;
          if (t.length === 1)
            return maior.some(function (o) {
              return o[0] === t;
            });
          return maior.some(function (o) {
            return o.length === 1 && o === t[0];
          });
        }
        function nomesCasam(a, b) {
          if (!a.length || !b.length) return false;
          var menor = a.length <= b.length ? a : b;
          var maior = a.length <= b.length ? b : a;
          if (menor.length < 2) return menor.join(" ") === maior.join(" ");
          var set = {};
          maior.forEach(function (t) {
            set[t] = 1;
          });
          return menor.every(function (t) {
            return tokenCasa(t, maior, set);
          });
        }

        // Converte serial do Excel em Date
        function serialParaData(n) {
          var ms = Math.round((n - 25569) * 86400 * 1000);
          return new Date(ms);
        }

        // Extrai "MM/AAAA" (competência) de Date, número-serial ou texto
        function competencia(valor) {
          if (valor === null || valor === undefined || valor === "") return "";
          if (valor instanceof Date && !isNaN(valor)) {
            return pad2(valor.getUTCMonth() + 1) + "/" + valor.getUTCFullYear();
          }
          if (typeof valor === "number") {
            var d = serialParaData(valor);
            return pad2(d.getUTCMonth() + 1) + "/" + d.getUTCFullYear();
          }
          var t = String(valor).trim();
          // dd/mm/aaaa  ou  d-m-aaaa
          var m = t.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
          if (m) {
            return pad2(m[2]) + "/" + ano4(m[3]);
          }
          // mm/aaaa
          m = t.match(/^(\d{1,2})[\/\-.](\d{4})$/);
          if (m) {
            return pad2(m[1]) + "/" + m[2];
          }
          // aaaa-mm-dd
          m = t.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
          if (m) {
            return pad2(m[2]) + "/" + m[1];
          }
          return t.toUpperCase();
        }

        function pad2(n) {
          n = String(n);
          return n.length < 2 ? "0" + n : n;
        }
        function ano4(a) {
          a = String(a);
          return a.length === 2 ? "20" + a : a;
        }

        // valor de hora comparável (minutos desde meia-noite)
        function horaEmMinutos(valor) {
          if (valor instanceof Date && !isNaN(valor))
            return valor.getUTCHours() * 60 + valor.getUTCMinutes();
          if (typeof valor === "number") {
            // fração do dia
            var f = valor - Math.floor(valor);
            return Math.round(f * 24 * 60);
          }
          var m = String(valor).match(/(\d{1,2})[:h](\d{2})/);
          if (m) return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
          return 99999; // sem hora vai para o fim
        }

        // localiza a chave de coluna cujo cabeçalho CONTÉM algum dos sinônimos
        function acharColuna(headers, sinonimos) {
          for (var i = 0; i < headers.length; i++) {
            var h = normalizar(headers[i]);
            for (var j = 0; j < sinonimos.length; j++) {
              if (h.indexOf(sinonimos[j]) !== -1) return headers[i];
            }
          }
          return null;
        }

        // localiza coluna por correspondência EXATA (para tokens curtos como "mes"/"ano",
        // que dariam falso positivo dentro de "PLANO", "TÉRMINO" etc.)
        function acharColunaExata(headers, sinonimos) {
          for (var i = 0; i < headers.length; i++) {
            if (sinonimos.indexOf(normalizar(headers[i])) !== -1)
              return headers[i];
          }
          return null;
        }

        // localiza a linha de cabeçalho real (planilhas cruas trazem linhas de
        // título/identificação da clínica antes da linha das colunas)
        function acharLinhaCabecalho(aoa) {
          var grupos = [
            ["ASSISTIDO", "PACIENTE", "NOME", "CLIENTE"],
            ["CONVENIO", "PLANO"],
            ["COMPETENCIA", "REFERENCIA", "MES", "ANO", "DATA"],
          ];
          for (var i = 0; i < aoa.length && i < 30; i++) {
            var cels = (aoa[i] || []).map(normalizar);
            var acertos = 0;
            grupos.forEach(function (g) {
              if (
                cels.some(function (c) {
                  return (
                    c &&
                    g.some(function (s) {
                      return c.indexOf(s) !== -1;
                    })
                  );
                })
              )
                acertos++;
            });
            if (acertos >= 2) return i;
          }
          return aoa.length ? 0 : -1;
        }

        // remove pontuação do convênio (hífens, barras…), deixando só letras,
        // números e espaços. Ex.: "PRO SAUDE - TJDFT" e "PRO-SAUDE TJDFT" → iguais
        function limparConvenio(c) {
          return String(c)
            .replace(/[^A-Z0-9 ]+/g, " ")
            .replace(/\s+/g, " ")
            .trim();
        }

        // convênios casam se forem iguais OU se um contiver o outro
        // ("CBMDF ABA" ~ "CBMDF"), ignorando pontuação,
        // mas não "BRADESCO SAUDE" x "PORTO SAUDE"
        function convenioCasa(a, b) {
          a = limparConvenio(a);
          b = limparConvenio(b);
          if (!a || !b) return false;
          if (a === b) return true;
          if (a.length >= 4 && b.indexOf(a) !== -1) return true;
          if (b.length >= 4 && a.indexOf(b) !== -1) return true;
          return false;
        }

        // reduz o nome do procedimento (agenda OU guia) a uma CATEGORIA canônica,
        // para validar atendimento x guia mesmo com textos diferentes nos dois sistemas.
        // ORDEM IMPORTA: "ABA" é testado primeiro — uma sessão "FONO COM TERAPIA ABA"
        // é validada por guia de TERAPIA ABA, não pela especialidade base.
        function categoria(servico) {
          var s = normalizar(servico);
          if (!s) return "";
          if (s.indexOf("ABA") !== -1) return "TERAPIA ABA";
          if (s.indexOf("NEUROPSICOL") !== -1) return "AVALIACAO NEUROPSICOLOGICA";
          if (s.indexOf("FONOAUDIOLOG") !== -1) return "FONOAUDIOLOGIA";
          if (s.indexOf("MUSICOTERAP") !== -1) return "MUSICOTERAPIA";
          if (s.indexOf("NUTRI") !== -1) return "NUTRICAO";
          if (s.indexOf("PSICOMOTRIC") !== -1) return "PSICOMOTRICIDADE";
          if (s.indexOf("PSICOPEDAG") !== -1) return "PSICOPEDAGOGIA";
          if (s.indexOf("FISIOTERAP") !== -1) return "FISIOTERAPIA";
          if (s.indexOf("TERAPIA OCUPACIONAL") !== -1) return "TERAPIA OCUPACIONAL";
          if (s.indexOf("PSICOLOG") !== -1 || s.indexOf("PSICOTERAP") !== -1)
            return "PSICOTERAPIA";
          if (s.indexOf("PACOTE") !== -1) return "PACOTE";
          return "OUTROS";
        }

        // mapa normalizado da tabela de-para (montado uma vez)
        var MAPA_CONSOLIDADO = {};
        DE_PARA_PROCEDIMENTOS.forEach(function (par) {
          MAPA_CONSOLIDADO[normalizar(par[0])] = par[1];
        });

        // categoria de um serviço: 1º a tabela de-para curada; senão, palavra-chave
        function consolidar(servico) {
          var n = normalizar(servico);
          if (MAPA_CONSOLIDADO[n]) return MAPA_CONSOLIDADO[n];
          return categoria(servico);
        }

        // categoria de um agendamento: consolida pelo Serviço; se não identificar
        // (OUTROS), usa a coluna Especialidade da agenda (profissional → categoria)
        function categoriaAgenda(servico, especialidade) {
          var c = consolidar(servico);
          if (!c || c === "OUTROS") {
            var ce = categoria(especialidade);
            if (ce && ce !== "OUTROS") return ce;
          }
          return c;
        }

        // convênios cujas sessões (exceto psicoterapia e avaliação neuro) são
        // cobertas por uma guia de PACOTE de horas, e não por guia da especialidade
        var CONVENIOS_PACOTE = ["SAUDE CAIXA"];
        function isConvenioPacote(conv) {
          return CONVENIOS_PACOTE.some(function (c) {
            return conv.indexOf(c) !== -1;
          });
        }

        // convênios que RENOVAM a guia toda semana: por especialidade usamos
        // sempre a guia mais recente (por data de criação), que corresponde à
        // semana do atendimento. Avaliação neuro é exceção (dura mais).
        var CONVENIOS_RENOVA_SEMANAL = ["BRADESCO"];
        function isConvenioRenovaSemanal(conv) {
          return CONVENIOS_RENOVA_SEMANAL.some(function (c) {
            return conv.indexOf(c) !== -1;
          });
        }

        // converte a data de criação da guia em número comparável (timestamp);
        // aceita Date nativo, "AAAA-MM-DD HH:MM:SS" e "DD/MM/AAAA"
        function tempoCriacao(v) {
          if (v instanceof Date && !isNaN(v)) return v.getTime();
          if (v == null || v === "") return 0;
          var s = String(v);
          var m = s.match(
            /(\d{4})\D(\d{1,2})\D(\d{1,2})(?:\D+(\d{1,2})\D(\d{1,2})(?:\D(\d{1,2}))?)?/,
          );
          if (m)
            return new Date(
              +m[1],
              +m[2] - 1,
              +m[3],
              +(m[4] || 0),
              +(m[5] || 0),
              +(m[6] || 0),
            ).getTime();
          var d = s.match(/(\d{1,2})\D(\d{1,2})\D(\d{4})/);
          if (d) return new Date(+d[3], +d[2] - 1, +d[1]).getTime();
          return 0;
        }

        // --- leitura de arquivo --------------------------------------------

        function lerPlanilha(file) {
          return new Promise(function (resolve, reject) {
            var reader = new FileReader();
            reader.onload = function (e) {
              try {
                var wb = XLSX.read(new Uint8Array(e.target.result), {
                  type: "array",
                  cellDates: true,
                });
                var ws = wb.Sheets[wb.SheetNames[0]];
                var aoa = XLSX.utils.sheet_to_json(ws, {
                  header: 1,
                  defval: "",
                  blankrows: false,
                });
                var hi = acharLinhaCabecalho(aoa);
                if (hi === -1) {
                  resolve({ headers: [], rows: [] });
                  return;
                }

                // cabeçalhos (preenche colunas sem nome para não colidirem)
                var headers = (aoa[hi] || []).map(function (h, idx) {
                  h = String(h).trim();
                  return h === "" ? "Coluna_" + (idx + 1) : h;
                });

                // monta as linhas a partir da linha seguinte ao cabeçalho
                var rows = [];
                for (var r = hi + 1; r < aoa.length; r++) {
                  var arr = aoa[r] || [];
                  if (
                    arr.every(function (c) {
                      return String(c).trim() === "";
                    })
                  )
                    continue; // pula linhas vazias
                  var obj = {};
                  headers.forEach(function (h, idx) {
                    obj[h] = arr[idx] !== undefined ? arr[idx] : "";
                  });
                  rows.push(obj);
                }
                resolve({ headers: headers, rows: rows });
              } catch (err) {
                reject(err);
              }
            };
            reader.onerror = function () {
              reject(new Error("Falha ao ler o arquivo."));
            };
            reader.readAsArrayBuffer(file);
          });
        }

        // --- estado da UI ---------------------------------------------------

        function atualizarBotao() {
          btn.disabled = !(arqAgenda.files.length && arqGuias.files.length);
        }

        arqAgenda.addEventListener("change", function () {
          document.getElementById("statusAgenda").textContent = arqAgenda.files
            .length
            ? arqAgenda.files[0].name
            : "";
          limparResultado();
          atualizarBotao();
        });
        arqGuias.addEventListener("change", function () {
          document.getElementById("statusGuias").textContent = arqGuias.files
            .length
            ? arqGuias.files[0].name
            : "";
          limparResultado();
          atualizarBotao();
        });

        function limparResultado() {
          status.innerHTML = "";
          resumo.style.display = "none";
        }

        // --- processamento --------------------------------------------------

        btn.addEventListener("click", function () {
          limparResultado();
          btn.disabled = true;
          status.textContent = "Processando...";

          Promise.all([
            lerPlanilha(arqAgenda.files[0]),
            lerPlanilha(arqGuias.files[0]),
          ])
            .then(function (res) {
              cruzar(res[0], res[1]);
            })
            .catch(function (err) {
              status.innerHTML =
                '<span class="erro">Erro ao processar: ' +
                (err.message || err) +
                "</span>";
              atualizarBotao();
            });
        });

        function cruzar(agenda, guias) {
          // --- localizar colunas (agendamentos) ---
          var colNomeA = acharColuna(agenda.headers, [
            "ASSISTIDO",
            "PACIENTE",
            "NOME",
            "CLIENTE",
          ]);
          var colConvA = acharColuna(agenda.headers, ["CONVENIO", "PLANO"]);
          var colDataA = acharColuna(agenda.headers, ["DATA", "DT"]);
          var colHoraA = acharColuna(agenda.headers, [
            "INICIO",
            "HORA",
            "HORARIO",
            "ENTRADA",
          ]);
          var colServA = acharColuna(agenda.headers, [
            "SERVICO",
            "PROCEDIMENTO",
          ]);
          var colEspA = acharColuna(agenda.headers, ["ESPECIALIDADE"]);

          // --- localizar colunas (guias) ---
          var colNomeG = acharColuna(guias.headers, [
            "ASSISTIDO",
            "PACIENTE",
            "NOME",
            "CLIENTE",
          ]);
          var colConvG = acharColuna(guias.headers, ["CONVENIO", "PLANO"]);
          var colStatusG = acharColuna(guias.headers, ["STATUS", "SITUACAO"]);
          var colGuiaG = acharColuna(guias.headers, ["GUIA"]);
          var colServicoG = acharColuna(guias.headers, [
            "SERVICO",
            "PROCEDIMENTO",
            "ESPECIALIDADE",
            
          ]);
          var colDataCriacaoG = acharColuna(guias.headers, [
            "CRIACAO",
            "EMISSAO",
            "DATA-CRIACAO",
          ]);
          // competência pode vir como coluna única OU separada em mês + ano
          var colCompG = acharColuna(guias.headers, [
            "COMPETENCIA",
            "REFERENCIA",
          ]);
          var colMesG = acharColunaExata(guias.headers, ["MES", "MÊS"]);
          var colAnoG = acharColunaExata(guias.headers, ["ANO"]);

          var faltando = [];
          if (!colNomeA) faltando.push("Nome do paciente (agendamentos)");
          if (!colConvA) faltando.push("Convênio (agendamentos)");
          if (!colDataA) faltando.push("Data (agendamentos)");
          if (!colNomeG) faltando.push("Nome do paciente (guias)");
          if (!colConvG) faltando.push("Convênio (guias)");
          if (!colCompG && !(colMesG && colAnoG))
            faltando.push("Competência (guias) — coluna única ou mês+ano");
          if (faltando.length) {
            status.innerHTML =
              '<span class="erro">Não encontrei as colunas: ' +
              faltando.join("; ") +
              ". Verifique os cabeçalhos das planilhas.</span>";
            atualizarBotao();
            return;
          }

          // competência de uma linha de guia (mês+ano separados ou coluna única)
          function compDaGuia(g) {
            if (colMesG && colAnoG)
              return pad2(g[colMesG]) + "/" + ano4(g[colAnoG]);
            return competencia(g[colCompG]);
          }

          // só validamos por procedimento se as duas planilhas tiverem a coluna
          var validarProcedimento = !!(colServA && colServicoG);

          // --- índice das guias por COMPETÊNCIA → lista de guias ---
          // (o nome é comparado por tokens na busca, o convênio com tolerância,
          //  e a categoria valida o procedimento; o status indica onde a guia está)
          var indiceGuias = {};
          guias.rows.forEach(function (g) {
            var comp = compDaGuia(g);
            (indiceGuias[comp] = indiceGuias[comp] || []).push({
              conv: normalizar(g[colConvG]),
              status: colStatusG ? String(g[colStatusG]).trim() : "",
              guia: colGuiaG ? g[colGuiaG] : "",
              cat: colServicoG ? consolidar(g[colServicoG]) : "",
              criacao: colDataCriacaoG ? tempoCriacao(g[colDataCriacaoG]) : 0,
              tokens: tokensNome(g[colNomeG]),
            });
          });

          // retorna a primeira guia ainda NÃO usada de uma lista (ou null)
          function primeiraLivre(lista) {
            for (var i = 0; i < lista.length; i++) {
              if (!lista[i].usada) return lista[i];
            }
            return null;
          }

          // escolhe a guia para um agendamento, CONSUMINDO uma guia por sessão:
          // cada guia de procedimento é atribuída no máximo a uma sessão, para que
          // cada agendamento tenha o SEU próprio número de guia.
          // retorna { guia, obs } ou null (pendente). obs != "" sinaliza divergência.
          function escolherGuia(convA, agendaCat, guiasDoPaciente) {
            // só guias do mesmo convênio (com tolerância de nome)
            var doConvenio = guiasDoPaciente.filter(function (g) {
              return convenioCasa(convA, g.conv);
            });
            if (!doConvenio.length) return null;

            // sem validação por procedimento → consome qualquer guia livre (1 por sessão)
            if (!validarProcedimento) {
              var qualquer = primeiraLivre(doConvenio);
              if (qualquer) {
                qualquer.usada = true;
                return { guia: qualquer, obs: "" };
              }
              return null;
            }

            // 1) match limpo: guia da MESMA categoria do agendamento
            if (agendaCat && agendaCat !== "OUTROS") {
              var mesmaCat = doConvenio.filter(function (g) {
                return g.cat === agendaCat;
              });
              if (mesmaCat.length) {
                // Bradesco (renovação semanal), exceto avaliação neuro: usa SEMPRE
                // a guia mais recente por data — corresponde à semana do atendimento;
                // é compartilhada entre as sessões (não consome) e ignora as antigas.
                if (
                  isConvenioRenovaSemanal(convA) &&
                  agendaCat !== "AVALIACAO NEUROPSICOLOGICA"
                ) {
                  var recente = mesmaCat[0];
                  for (var k = 1; k < mesmaCat.length; k++) {
                    if (mesmaCat[k].criacao > recente.criacao)
                      recente = mesmaCat[k];
                  }
                  return { guia: recente, obs: "" };
                }
                // demais: distribui uma guia distinta por sessão (consome)
                var livre = primeiraLivre(mesmaCat);
                if (livre) {
                  livre.usada = true;
                  return { guia: livre, obs: "" };
                }
                // há guias dessa categoria, mas todas já foram atribuídas a outras
                // sessões (mais sessões do que guias) → pendente
                return null;
              }
            }

            // 2) convênio-pacote (ex.: Saúde Caixa): psicoterapia e avaliação neuro
            //    casam pela própria categoria (acima); o RESTO é coberto por uma
            //    guia de PACOTE de horas — que é COMPARTILHADA (não se consome)
            if (
              isConvenioPacote(convA) &&
              agendaCat !== "PSICOTERAPIA" &&
              agendaCat !== "AVALIACAO NEUROPSICOLOGICA"
            ) {
              for (var i = 0; i < doConvenio.length; i++) {
                if (doConvenio[i].cat === "PACOTE")
                  return { guia: doConvenio[i], obs: "" };
              }
            }

            // 3) tem guia no convênio/mês, mas de outro procedimento → assina sinalizando
            //    (apenas informativo; não consome a guia)
            return { guia: doConvenio[0], obs: "Guia de outro procedimento" };
          }

          // --- cruzar cada agendamento ---
          var encontradas = 0;
          var sinalizadas = 0;
          var porStatus = {};
          var saida = agenda.rows.map(function (a) {
            var convA = normalizar(a[colConvA]);
            var agendaCat = colServA
              ? categoriaAgenda(a[colServA], colEspA ? a[colEspA] : "")
              : "";
            // candidatos: guias da mesma competência cujo nome casa por tokens
            var compA = competencia(a[colDataA]);
            var tokensA = tokensNome(a[colNomeA]);
            var candidatos = (indiceGuias[compA] || []).filter(function (g) {
              return nomesCasam(tokensA, g.tokens);
            });
            var res = escolherGuia(convA, agendaCat, candidatos);

            var linha = Object.assign({}, a);
            if (res) {
              encontradas++;
              var st = res.guia.status || "Guia encontrada";
              linha["Situação da Guia"] = st;
              linha["guias"] = res.guia.guia != null ? res.guia.guia : "";
              linha["Obs. Validação"] = res.obs;
              if (res.obs) sinalizadas++;
              porStatus[st] = (porStatus[st] || 0) + 1;
            } else {
              linha["Situação da Guia"] = "";
              linha["guias"] = "";
              linha["Obs. Validação"] = "";
            }
            return linha;
          });

          // --- ordenar: Convênio A→Z, depois Nome A→Z, depois Hora crescente ---
          saida.sort(function (x, y) {
            var c = normalizar(x[colConvA]).localeCompare(
              normalizar(y[colConvA]),
            );
            if (c !== 0) return c;
            var n = normalizar(x[colNomeA]).localeCompare(
              normalizar(y[colNomeA]),
            );
            if (n !== 0) return n;
            if (colHoraA)
              return horaEmMinutos(x[colHoraA]) - horaEmMinutos(y[colHoraA]);
            return 0;
          });

          // --- gerar e baixar a planilha ---
          // colunas a remover do resultado (por nome normalizado)
          var COLUNAS_EXCLUIR = ["ESPECIALIDADE", "OBS"];
          var ordemColunas = agenda.headers.filter(function (h) {
            return COLUNAS_EXCLUIR.indexOf(normalizar(h)) === -1;
          });
          // insere "guias" (número da guia) logo após a coluna de Convênio
          var idxConv = ordemColunas.indexOf(colConvA);
          if (idxConv !== -1) {
            ordemColunas.splice(idxConv + 1, 0, "guias");
          } else {
            ordemColunas.push("guias");
          }
          if (ordemColunas.indexOf("Situação da Guia") === -1)
            ordemColunas.push("Situação da Guia");
          if (ordemColunas.indexOf("Obs. Validação") === -1)
            ordemColunas.push("Obs. Validação");
          // monta cada linha SÓ com as colunas desejadas — o "header" do
          // json_to_sheet apenas ordena; chaves extras seriam anexadas ao fim
          var saidaFinal = saida.map(function (linha) {
            var o = {};
            ordemColunas.forEach(function (col) {
              o[col] = linha[col] !== undefined ? linha[col] : "";
            });
            return o;
          });
          var ws = XLSX.utils.json_to_sheet(saidaFinal, { header: ordemColunas });

          // largura automática das colunas (senão saem coladas/cortadas no Excel):
          // para cada coluna, usa o maior conteúdo entre o cabeçalho e as células
          ws["!cols"] = ordemColunas.map(function (col) {
            var maxLen = col.length;
            saidaFinal.forEach(function (linha) {
              var val = linha[col] == null ? "" : String(linha[col]);
              if (val.length > maxLen) maxLen = val.length;
            });
            return { wch: Math.min(maxLen + 2, 50) }; // +2 de folga, teto de 50
          });

          var wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, "Cruzamento");
          XLSX.writeFile(wb, "guias_cruzadas.xlsx");

          // --- resumo discreto ---
          var total = saida.length;
          var pendentes = total - encontradas;
          status.textContent = "Planilha gerada e baixada: guias_cruzadas.xlsx";
          // detalhamento por status (ex.: "Assinado", "Enviado a BM")
          var detalhe = Object.keys(porStatus)
            .map(function (st) {
              return (
                "<span><b>" +
                porStatus[st] +
                "</b> " +
                st.toLowerCase() +
                "</span>"
              );
            })
            .join("");
          var sinalizadasHtml = sinalizadas
            ? "<span><b>" +
              sinalizadas +
              "</b> com guia de outro procedimento (conferir)</span>"
            : "";
          resumo.innerHTML =
            "<span><b>" +
            total +
            "</b> agendamentos</span>" +
            "<span><b>" +
            encontradas +
            "</b> com guia encontrada</span>" +
            "<span><b>" +
            pendentes +
            "</b> pendentes</span>" +
            detalhe +
            sinalizadasHtml;
          resumo.style.display = "block";
          atualizarBotao();
        }
      })();