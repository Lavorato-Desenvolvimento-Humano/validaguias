(function () {
        "use strict";

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
            .replace(/[̀-ͯ]/g, "") // remove acentos
            .toUpperCase()
            .replace(/\s+/g, " ")
            .trim();
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

        // convênios casam se forem iguais OU se um contiver o outro
        // ("CBMDF ABA" ~ "CBMDF"), mas não "BRADESCO SAUDE" x "PORTO SAUDE"
        function convenioCasa(a, b) {
          if (!a || !b) return false;
          if (a === b) return true;
          if (a.length >= 4 && b.indexOf(a) !== -1) return true;
          if (b.length >= 4 && a.indexOf(b) !== -1) return true;
          return false;
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

          // --- índice das guias por NOME + COMPETÊNCIA → lista de {convênio, status} ---
          // (o convênio é checado depois com tolerância, pois os sistemas usam rótulos diferentes;
          //  guarda o status para indicar ONDE a guia está: "Assinado" ou "Enviado a BM")
          var indiceGuias = {};
          guias.rows.forEach(function (g) {
            var chave = normalizar(g[colNomeG]) + "|" + compDaGuia(g);
            (indiceGuias[chave] = indiceGuias[chave] || []).push({
              conv: normalizar(g[colConvG]),
              status: colStatusG ? String(g[colStatusG]).trim() : "",
              guia: colGuiaG ? g[colGuiaG] : "",
            });
          });

          // --- cruzar cada agendamento ---
          var encontradas = 0;
          var porStatus = {};
          var saida = agenda.rows.map(function (a) {
            var convA = normalizar(a[colConvA]);
            var chave =
              normalizar(a[colNomeA]) + "|" + competencia(a[colDataA]);
            var guiasDoPaciente = indiceGuias[chave] || [];
            var achada = null;
            for (var i = 0; i < guiasDoPaciente.length; i++) {
              if (convenioCasa(convA, guiasDoPaciente[i].conv)) {
                achada = guiasDoPaciente[i];
                break;
              }
            }
            var linha = Object.assign({}, a);
            if (achada) {
              encontradas++;
              var st = achada.status || "Guia encontrada";
              linha["Situação da Guia"] = st;
              linha["guias"] = achada.guia != null ? achada.guia : "";
              porStatus[st] = (porStatus[st] || 0) + 1;
            } else {
              linha["Situação da Guia"] = "";
              linha["guias"] = "";
            }
            return linha;
          });

          // --- ordenar: Hora crescente (principal), depois Nome A→Z e Convênio A→Z ---
          saida.sort(function (x, y) {
            if (colHoraA) {
              var h = horaEmMinutos(x[colHoraA]) - horaEmMinutos(y[colHoraA]);
              if (h !== 0) return h;
            }
            var n = normalizar(x[colNomeA]).localeCompare(
              normalizar(y[colNomeA]),
            );
            if (n !== 0) return n;
            return normalizar(x[colConvA]).localeCompare(
              normalizar(y[colConvA]),
            );
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
            detalhe;
          resumo.style.display = "block";
          atualizarBotao();
        }
      })();