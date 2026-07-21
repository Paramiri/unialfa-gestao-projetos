/* App — Diretrizes para a Gestão de Projetos UNIALFA */
(function () {
  'use strict';

  const root = document.getElementById('app');

  function esc(s) {
    return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function artChip(code) {
    const a = ARTEFATOS[code];
    if (!a) return '';
    if (a.arquivo) {
      return `<a class="art-chip ready" href="${encodeURI(a.arquivo)}" target="_blank" rel="noopener">${esc(a.nome)}<span class="arrow">↗</span></a>`;
    }
    return `<span class="art-chip missing" title="Ainda não construído">${esc(a.nome)} · a construir</span>`;
  }

  function gateBadgeHtml(gateId, size) {
    const g = GATES[gateId];
    return `<button class="gate-badge" data-gate="${g.id}" title="${esc(g.nome)}">
      <span class="gate-diamond"></span>
      <span class="g-lbl">${esc(g.nome.replace('Gate ', 'GATE '))}</span>
    </button>`;
  }

  function gateCardHtml(gateId) {
    const g = GATES[gateId];
    return `<div class="gate-card" id="gate-${g.id}">
      <div class="gc-head"><span class="gc-diamond"></span><h3>${esc(g.nome)}</h3></div>
      <div class="gc-grid">
        <div class="gc-field"><span>Quando</span><b>${esc(g.quando)}</b></div>
        <div class="gc-field"><span>Quem decide</span><b>${esc(g.quemDecide)}</b></div>
        <div class="gc-field"><span>Sobre qual artefato</span><b>${esc(g.sobreArtefato)}</b></div>
        <div class="gc-field"><span>Efeito</span><b>${esc(g.efeito)}</b></div>
      </div>
    </div>`;
  }

  function tag(d) {
    return d.id.replace('D', '');
  }

  // ---------------- Sidebar ----------------
  function renderSidebar(activeId) {
    const seqItems = SEQUENCIAIS.map(d => `
      <a class="nav-item ${activeId === d.id ? 'active' : ''}" href="#/${d.id}">
        <span class="tag">${tag(d)}</span><span class="lbl">${esc(d.titulo)}</span>
      </a>`).join('');
    const tvItems = TRANSVERSAIS.map(d => `
      <a class="nav-item ${activeId === d.id ? 'active' : ''}" href="#/${d.id}">
        <span class="tag">${tag(d)}</span><span class="lbl">${esc(d.titulo)}</span><span class="pill">Transv.</span>
      </a>`).join('');

    return `
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-brand">
          <div class="logo">UNIALFA<small>GERÊNCIA DE PROJETOS</small></div>
          <p>Diretrizes, Estratégias e Práticas</p>
        </div>
        <nav class="sidebar-nav">
          <a class="nav-item ${activeId === 'home' ? 'active' : ''}" href="#/">
            <span class="tag">◎</span><span class="lbl">Visão geral</span>
          </a>
          <div class="sidebar-section-label">Fluxo sequencial (D01–D05)</div>
          ${seqItems}
          <div class="sidebar-section-label">Transversais</div>
          ${tvItems}
          <div class="sidebar-section-label">Ferramentas</div>
          <a class="nav-item" href="14 - validador-projetos.html" target="_blank" rel="noopener">
            <span class="tag">◈</span><span class="lbl">Validador de Projetos</span>
          </a>
          <div class="sidebar-section-label">Gestão</div>
          <a class="nav-item" href="13 - administracao-usuarios.html" target="_blank" rel="noopener">
            <span class="tag">⚙</span><span class="lbl">Administração</span>
          </a>
        </nav>
        <div class="sidebar-foot">12 artefatos institucionais · 2 gates de autorização</div>
      </aside>
      <div class="scrim" id="scrim"></div>
      <div class="mobile-bar">
        <button id="menuBtn" aria-label="Abrir menu">☰</button>
        <span class="mb-title">UNIALFA · DIRETRIZES</span>
      </div>`;
  }

  // ---------------- Home / Map ----------------
  function renderHome() {
    const seqNodes = SEQUENCIAIS.map((d, i) => {
      let out = `<a class="flow-node" href="#/${d.id}">
          <span class="num">${tag(d)}</span>
          <h3>${esc(d.titulo)}</h3>
          <p>${esc(d.resumo)}</p>
        </a>`;
      return out;
    });

    // interleave arrows / gates between sequential nodes
    let flowHtml = '';
    SEQUENCIAIS.forEach((d, i) => {
      flowHtml += seqNodes[i];
      if (i < SEQUENCIAIS.length - 1) {
        if (d.id === 'D01') flowHtml += gateBadgeHtml('gate1');
        else if (d.id === 'D02') flowHtml += gateBadgeHtml('gate2');
        else flowHtml += `<span class="flow-arrow">→</span>`;
      }
    });

    const tvHtml = TRANSVERSAIS.map(d => {
      const chips = d.grupos[0].estrategias.map(e => `<span class="tv-chip">${esc(e.id)}</span>`).join('');
      return `<a class="transversal-track" href="#/${d.id}">
        <span class="tv-tag">${tag(d)}</span>
        <span class="tv-body">
          <h4>${esc(d.titulo)} <em style="font-style:normal;color:var(--muted-2);font-weight:600">— transversal, D01 a D05</em></h4>
          <p>${esc(d.resumo)}</p>
          <span class="tv-chips">${chips}</span>
        </span>
      </a>`;
    }).join('');

    return `
      <div class="hero">
        <h1>Diretrizes para a Gestão de Projetos — UNIALFA</h1>
        <p>7 diretrizes (D01–D07), 61 estratégias e 12 artefatos institucionais que orientam o ciclo de vida de um projeto na UNIALFA — da demanda ao encerramento, com governança e liderança presentes o tempo todo. Clique em qualquer etapa do mapa para ver o detalhe.</p>
      </div>

      <div class="comm-band">Comunicação <span>— atravessa todas as diretrizes, do D01 ao D07</span></div>

      <div class="map-wrap">
        <div class="flow-row">${flowHtml}</div>
        <div>${tvHtml}</div>
        <div class="legend">
          <span><i style="background:var(--red)"></i> Diretriz sequencial (D01–D05)</span>
          <span><i style="background:var(--ink)"></i> Diretriz transversal (D06–D07)</span>
          <span><i style="background:var(--gold);transform:rotate(45deg)"></i> Gate de autorização institucional</span>
        </div>
      </div>

      <div class="home-grid">
        <div class="info-card">
          <h3>Os 2 Gates de autorização</h3>
          <ul>
            <li><b>Gate 1 — Triagem</b> (em D01.4/D01.5): o Gestor Responsável aprova ou rejeita a entrada da demanda no portfólio.</li>
            <li><b>Gate 2 — Pactuação</b> (fim do D02): o mais crítico — o Dono do Negócio pactua o Relatório de Entregas e Benefícios com a Alta Gestão, autorizando o início da execução.</li>
          </ul>
        </div>
        <div class="info-card">
          <h3>Artefatos sempre acessíveis</h3>
          <ul>
            <li><b>Ata de Reunião</b> (FORALF00340) — qualquer interação formal, do D01 ao D07.</li>
            <li><b>SMP</b> (FORALF00343) — solicitação de mudança, a qualquer momento da execução/controle.</li>
            <li><b>Relatório de Situação de Projetos</b> (FORALF11) — cobre também o Reporte de Resultados, recorrente, ligado à governança (D06.5–D06.7).</li>
          </ul>
          <p style="margin-top:10px;font-size:11px;color:var(--muted-2)">Esses três não pertencem à esteira sequencial D01→D05 — não bloqueiam nem são bloqueados por ela.</p>
        </div>
      </div>`;
  }

  // ---------------- Detail ----------------
  function estrategiaHtml(e) {
    const chips = (e.artefatos || []).map(artChip).join('');
    return `<div class="acc-item" data-acc>
      <div class="acc-head">
        <button class="acc-trigger" data-toggle>
          <span class="acc-id">${esc(e.id)}</span>
          <span class="acc-title" title="${esc(e.titulo)}">${esc(e.titulo)}</span>
          <span class="acc-chev">▾</span>
        </button>
        ${chips ? `<div class="acc-quicklinks">${chips}</div>` : ''}
      </div>
      <div class="acc-body">
        <div class="acc-body-in">
          <ul>${e.pontos.map(p => `<li>${esc(p)}</li>`).join('')}</ul>
        </div>
      </div>
    </div>`;
  }

  function renderDetail(id) {
    const d = getDiretriz(id);
    if (!d) return renderHome();

    const entregaLi = d.entrega.map(x => `<li>${esc(x)}</li>`).join('');
    const pessoasLi = d.pessoas.map(x => `<li>${esc(x)}</li>`).join('');

    const gateBefore = Object.values(GATES).find(g => g.diretriz === d.id && false); // reserved
    let strategiesHtml = '';

    d.grupos.forEach(grp => {
      if (grp.titulo) {
        strategiesHtml += `<div class="strat-group-title">${esc(grp.id)} — ${esc(grp.titulo)}</div>`;
      } else {
        strategiesHtml += `<div class="strategies-title">Estratégias recomendadas</div>`;
      }
      strategiesHtml += `<div class="accordion">${grp.estrategias.map(estrategiaHtml).join('')}</div>`;

      // insert gate card right after the strategy it follows
      Object.values(GATES).forEach(g => {
        if (grp.estrategias.some(e => e.id === g.apos)) {
          strategiesHtml += gateCardHtml(g.id);
        }
      });
    });

    return `
      <div class="crumb"><a href="#/">Visão geral</a> <span>/</span> <span>${esc(d.id)} — ${esc(d.titulo)}</span></div>

      <div class="d-head">
        <div class="d-head-band">
          <div class="d-num">${tag(d)}</div>
          <div class="d-title">
            <h1>${esc(d.id)} — ${esc(d.titulo)}</h1>
            <div class="d-badge">${d.tipo === 'transversal' ? 'Diretriz transversal · presente do D01 ao D05' : `Diretriz sequencial · etapa ${d.ordem} de 5`}</div>
          </div>
          <div class="badge-pill">${d.tipo === 'transversal' ? 'Transversal' : 'Sequencial'}</div>
        </div>
        <div class="d-head-body">
          <div><h4>O que é?</h4><p>${esc(d.oQueE)}</p></div>
          <div><h4>Por que é importante?</h4><p>${esc(d.porQue)}</p></div>
        </div>
        <div class="ficha">
          <div><h4>O que a diretriz entrega?</h4><ul>${entregaLi}</ul></div>
          <div><h4>Quem está envolvido?</h4><ul>${pessoasLi}</ul></div>
        </div>
      </div>

      ${strategiesHtml}
    `;
  }

  // ---------------- Router ----------------
  function currentId() {
    const h = location.hash.replace('#/', '').trim();
    return h || 'home';
  }

  function render() {
    const id = currentId();
    const body = id === 'home' ? renderHome() : renderDetail(id);
    root.innerHTML = `
      ${renderSidebar(id)}
      <main class="main"><div class="main-inner">${body}</div></main>
    `;
    wireEvents();
    window.scrollTo(0, 0);
  }

  function wireEvents() {
    root.querySelectorAll('[data-toggle]').forEach(btn => {
      btn.addEventListener('click', () => {
        btn.closest('.acc-item').classList.toggle('open');
      });
    });
    root.querySelectorAll('[data-gate]').forEach(btn => {
      btn.addEventListener('click', () => {
        const g = GATES[btn.getAttribute('data-gate')];
        location.hash = `#/${g.diretriz}`;
        setTimeout(() => {
          const el = document.getElementById(`gate-${g.id}`);
          if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }, 60);
      });
    });
    const menuBtn = document.getElementById('menuBtn');
    const sidebar = document.getElementById('sidebar');
    const scrim = document.getElementById('scrim');
    if (menuBtn) {
      menuBtn.addEventListener('click', () => {
        sidebar.classList.add('open');
        scrim.classList.add('show');
      });
      scrim.addEventListener('click', () => {
        sidebar.classList.remove('open');
        scrim.classList.remove('show');
      });
      root.querySelectorAll('.nav-item').forEach(a => a.addEventListener('click', () => {
        sidebar.classList.remove('open');
        scrim.classList.remove('show');
      }));
    }
  }

  window.addEventListener('hashchange', render);
  window.addEventListener('DOMContentLoaded', render);
})();
