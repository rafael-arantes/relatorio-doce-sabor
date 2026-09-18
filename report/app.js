/* Relatório Financeiro — Doce Sabor
 * Carrega data.json e renderiza a interface. Sem dependências externas. */
(function () {
  "use strict";

  var state = { data: null, month: null };

  var FORMA_COLORS = {
    "Pix": "#0f766e",
    "Dinheiro": "#10b981",
    "Cartão": "#f59e0b",
    "Boleto": "#ef4444",
    "Outro": "#64748b",
  };

  function el(id) { return document.getElementById(id); }

  function formatBRL(value) {
    return (Math.round(value * 100) / 100).toLocaleString("pt-BR", {
      style: "currency", currency: "BRL",
    });
  }

  function formatPct(value) {
    return (value || 0).toLocaleString("pt-BR", {
      maximumFractionDigits: 1, minimumFractionDigits: 1,
    }) + "%";
  }

  function monthLabel(ym) {
    var parts = ym.split("-");
    var d = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
    var s = d.toLocaleDateString("pt-BR", { month: "long", year: "numeric" });
    return s.charAt(0).toUpperCase() + s.slice(1);
  }

  function colorForForma(value) {
    return FORMA_COLORS[value] || "#64748b";
  }

  function icon(name) {
    var paths = {
      "arrow-down": '<line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/>',
      "arrow-up": '<line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/>',
      "calendar": '<rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>',
      "dollar": '<line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>',
    };
    return (
      '<svg viewBox="0 0 24 24" width="24" height="24" fill="none" stroke="currentColor" ' +
      'stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
      (paths[name] || "") + "</svg>"
    );
  }

  function esc(text) {
    return String(text == null ? "" : text)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function renderHero(m) {
    var positive = m.lucro >= 0;
    var hero = el("hero");
    hero.className = "hero " + (positive ? "hero--positive" : "hero--negative");
    hero.innerHTML =
      '<div class="hero__top">' +
      '<span class="hero__label">' + (positive ? "Lucro" : "Prejuízo") + " do mês</span>" +
      '<span class="hero__month">' + monthLabel(state.month) + "</span>" +
      "</div>" +
      '<div class="hero__value">' + formatBRL(m.lucro) + "</div>" +
      '<div class="hero__note">Entradas ' + formatBRL(m.entradas) +
      " · Saídas " + formatBRL(m.saidas) +
      " · Fixas " + formatBRL(m.fixas) + "</div>";
  }

  function renderKpis(m) {
    var kpis = [
      { icon: icon("arrow-down"), label: "Entradas", sub: "receitas do mês", value: m.entradas, tone: "green" },
      { icon: icon("arrow-up"), label: "Saídas", sub: "despesas variáveis", value: m.saidas, tone: "red" },
      { icon: icon("calendar"), label: "Despesas fixas", sub: "contas fixas", value: m.fixas, tone: "amber" },
      {
        icon: icon("dollar"), label: "Saldo", sub: "entradas − saídas", value: m.saldo,
        tone: m.saldo >= 0 ? "auto-pos" : "auto-neg",
      },
    ];
    el("bigNumbers").innerHTML = kpis.map(function (k) {
      var valueClass = "";
      if (k.tone === "auto-pos") valueClass = "kpi__value--pos";
      if (k.tone === "auto-neg") valueClass = "kpi__value--neg";
      return (
        '<div class="kpi kpi--' + k.tone + '">' +
        '<div class="kpi__head">' +
        '<span class="kpi__icon">' + k.icon + "</span>" +
        '<span class="kpi__label">' + k.label + "</span>" +
        "</div>" +
        '<div class="kpi__value ' + valueClass + '">' + formatBRL(k.value) + "</div>" +
        '<div class="kpi__sub">' + k.sub + "</div>" +
        "</div>"
      );
    }).join("");
  }

  function rankRow(item, index, opts) {
    var o = opts || {};
    var muted = o.muted;
    var label = o.label != null ? o.label : (index + 1);
    var top = !muted && index < 3;
    var pct = item.pct || 0;
    var subtitle = item.subtitle ||
      (item.quantidade
        ? item.quantidade + (item.quantidade === 1 ? " lançamento" : " lançamentos")
        : "");
    var classes = "rank__item" +
      (muted ? " rank__item--muted" : "") +
      (top ? " rank__item--top" : "");
    return (
      '<li class="' + classes + '">' +
      '<div class="rank__top">' +
      '<span class="rank__rank">' + label + "</span>" +
      '<div class="rank__info">' +
      '<span class="rank__name">' + esc(item.nome) + "</span>" +
      (subtitle ? '<span class="rank__count">' + esc(subtitle) + "</span>" : "") +
      "</div>" +
      '<div class="rank__right">' +
      '<span class="rank__value">' + formatBRL(item.valor) + "</span>" +
      '<span class="rank__pct">' + formatPct(pct) + "</span>" +
      "</div>" +
      "</div>" +
      '<div class="rank__bar"><span class="rank__bar-fill" style="width:' +
      Math.max(1.5, pct) + '%"></span></div>' +
      "</li>"
    );
  }

  function renderRank(listId, items, options) {
    var opts = options || {};
    var list = el(listId);
    if (!items.length) {
      list.innerHTML = '<div class="empty">Sem dados para este mês.</div>';
      return;
    }

    var defLimit = opts.limit || 0;
    var expanded = !!list._expanded;
    var collapsed = !expanded && defLimit > 0 && items.length > defLimit;
    var visible = collapsed ? items.slice(0, defLimit) : items;

    var html = visible.map(function (item, i) { return rankRow(item, i); }).join("");

    if (collapsed) {
      var rest = items.slice(defLimit);
      var restVal = rest.reduce(function (s, x) { return s + x.valor; }, 0);
      var total = items.reduce(function (s, x) { return s + x.valor; }, 0);
      var noun = opts.noun || "itens";
      html += rankRow(
        {
          nome: "Outros " + noun,
          valor: restVal,
          pct: total ? restVal / total * 100 : 0,
          subtitle: rest.length + " " + noun,
        },
        defLimit,
        { muted: true, label: "+" },
      );
    }

    list.innerHTML = html;

    if (defLimit > 0 && items.length > defLimit) {
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "rank__toggle";
      var noun = opts.noun || "itens";
      btn.textContent = collapsed
        ? "Ver todos os " + items.length + " " + noun
        : "Mostrar menos";
      btn.addEventListener("click", function () {
        list._expanded = !list._expanded;
        renderRank(listId, items, options);
      });
      list.appendChild(btn);
    }
  }

  function renderDonut(donutId, legendId, items) {
    var donut = el(donutId);
    var legend = el(legendId);
    if (!items.length) {
      donut.style.background = "#e5e7eb";
      donut.innerHTML = "";
      legend.innerHTML = '<li class="empty">Sem dados.</li>';
      return;
    }
    var total = items.reduce(function (s, it) { return s + it.valor; }, 0);
    var cumulative = 0;
    var stops = items.map(function (item) {
      var start = cumulative;
      cumulative += item.valor / total * 100;
      return colorForForma(item.nome) + " " + start + "% " + cumulative + "%";
    });
    donut.style.background = "conic-gradient(" + stops.join(", ") + ")";
    donut.innerHTML =
      '<div class="donut__center"><div class="donut__total">' + formatBRL(total) +
      '</div><div class="donut__caption">total</div></div>';
    legend.innerHTML = items.map(function (item) {
      return (
        '<li class="legend__row">' +
        '<span class="legend__swatch" style="background:' + colorForForma(item.nome) + '"></span>' +
        '<span class="legend__name">' + esc(item.nome) + "</span>" +
        '<span class="legend__val">' + formatBRL(item.valor) + "</span>" +
        '<span class="legend__pct">' + formatPct(item.pct) + "</span>" +
        "</li>"
      );
    }).join("");
  }

  function renderFixas(m) {
    var list = el("fixasList");
    if (!m.fixasDetalhe.length) {
      list.innerHTML = '<li class="empty">Sem despesas fixas cadastradas.</li>';
      return;
    }
    list.innerHTML = m.fixasDetalhe.map(function (item) {
      return (
        "<li><span class='fixas__name'>" + esc(item.nome) + "</span>" +
        "<span class='fixas__value'>" + formatBRL(item.valor) + "</span></li>"
      );
    }).join("");
  }

  function buildMonthNav() {
    var select = el("monthSelect");
    select.innerHTML = state.data.meses.map(function (ym) {
      return '<option value="' + ym + '">' + monthLabel(ym) + "</option>";
    }).join("");
    select.addEventListener("change", function () { setMonth(select.value); });
    el("prevMonth").addEventListener("click", function () {
      var i = state.data.meses.indexOf(state.month);
      if (i > 0) setMonth(state.data.meses[i - 1]);
    });
    el("nextMonth").addEventListener("click", function () {
      var i = state.data.meses.indexOf(state.month);
      if (i >= 0 && i < state.data.meses.length - 1) setMonth(state.data.meses[i + 1]);
    });
  }

  function setMonth(ym) {
    if (!state.data.porMes[ym]) return;
    state.month = ym;
    var i = state.data.meses.indexOf(ym);
    el("monthSelect").value = ym;
    el("prevMonth").disabled = i <= 0;
    el("nextMonth").disabled = i >= state.data.meses.length - 1;
    render(ym);
  }

  function render(ym) {
    var m = state.data.porMes[ym];
    renderHero(m);
    renderKpis(m);
    renderRank("categoryRank", m.categorias);
    renderDonut("donutEntradas", "legendEntradas", m.entradasPorForma);
    renderDonut("donutSaidas", "legendSaidas", m.saidasPorForma);
    renderRank("supplierRank", m.fornecedores, { limit: 8, noun: "fornecedores" });
    renderFixas(m);
  }

  function renderFooter() {
    var d = state.data;
    var parts = ["Fonte: " + esc(d.fonte)];
    if (d.fonteAtualizadaEm) {
      parts.push("Planilha atualizada em " + new Date(d.fonteAtualizadaEm).toLocaleString("pt-BR"));
    }
    parts.push("Relatório gerado em " + new Date(d.geradoEm).toLocaleString("pt-BR"));
    el("footer").innerHTML = parts.join(" · ");
  }

  var lastGeradoEm = null;
  var toastTimer = null;

  function showToast(message) {
    var t = el("toast");
    if (!t) return;
    t.textContent = message;
    t.classList.add("toast--show");
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove("toast--show"); }, 5000);
  }

  function applyData(data, firstLoad) {
    state.data = data;
    lastGeradoEm = data.geradoEm || null;
    if (!data.meses || !data.meses.length) {
      el("bigNumbers").innerHTML = '<div class="empty">Nenhum dado disponível.</div>';
      return;
    }
    if (firstLoad) {
      buildMonthNav();
      setMonth(data.meses[data.meses.length - 1]);
    } else {
      var keep = state.month && data.porMes[state.month]
        ? state.month
        : data.meses[data.meses.length - 1];
      setMonth(keep);
    }
    renderFooter();
  }

  function loadData() {
    return fetch("data.json").then(function (r) {
      if (!r.ok) throw new Error("HTTP " + r.status);
      return r.json();
    });
  }

  function boot() {
    loadData()
      .then(function (data) { applyData(data, true); })
      .catch(function (err) {
        el("bigNumbers").innerHTML =
          '<div class="empty">Não foi possível carregar os dados (' + esc(err.message) + ").</div>";
      });

    // Verifica silenciosamente por novos dados a cada minuto.
    setInterval(function () {
      loadData()
        .then(function (data) {
          if (data.geradoEm && data.geradoEm !== lastGeradoEm) applyData(data, false);
        })
        .catch(function () { /* silencioso */ });
    }, 60_000);

    // Botão de atualização: pede ao servidor para re-ler a planilha e, só depois
    // que ele terminar, recarrega o data.json (na sequência, não em paralelo).
    var btn = el("refreshBtn");
    if (btn) {
      btn.addEventListener("click", function () {
        btn.classList.add("refresh-btn--loading");
        var viaServer = false;
        fetch("refresh", { method: "POST" })
          .then(function (res) {
            viaServer = res.ok;
            return res.ok ? res.json() : null;
          })
          .catch(function () { return null; })
          .then(function () { return loadData(); })
          .then(function (data) {
            applyData(data, false);
            showToast(
              viaServer
                ? "Dados atualizados ✓"
                : "Sem servidor de dados — o relatório está estático e não lê a planilha ao vivo.",
            );
          })
          .catch(function () { showToast("Não foi possível atualizar os dados."); })
          .then(function () {
            btn.classList.remove("refresh-btn--loading");
            btn.classList.add("refresh-btn--done");
            setTimeout(function () { btn.classList.remove("refresh-btn--done"); }, 1600);
          });
      });
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
