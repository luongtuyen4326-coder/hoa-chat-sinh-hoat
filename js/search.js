/* ═══════════════════════════════════════════════════
   search.js — Logic trang tra cứu hai chiều
   Chiều 1: Tra theo hóa chất  → CHEMICALS  → chemical.html?id=X
   Chiều 2: Tra theo sản phẩm → PRODUCTS   → products.html?id=X
   (đúng theo Mục 2.5.1 — "Lớp xử lý" hỗ trợ tra cứu 2 chiều)
   ═══════════════════════════════════════════════════ */

const input      = document.getElementById('searchInput');
const clearBtn   = document.getElementById('clearBtn');
const resultList = document.getElementById('resultList');

let currentMode = 'chemical'; // 'chemical' | 'product'

/* ── Đổi chiều tra cứu ── */
function setMode(el, mode) {
  document.querySelectorAll('.mode-tag').forEach(t => t.classList.remove('active'));
  el.classList.add('active');
  currentMode = mode;

  document.getElementById('quickTagsChemical').style.display = mode === 'chemical' ? 'flex' : 'none';
  document.getElementById('quickTagsProduct').style.display  = mode === 'product'  ? 'flex' : 'none';
  input.placeholder = mode === 'chemical'
    ? 'Nhập tên hóa chất, tên đồng nghĩa, công thức hoặc số CAS...'
    : 'Nhập tên sản phẩm, nhà sản xuất hoặc nhóm sản phẩm...';

  if (input.value.trim()) doSearch();
}

/* ── Hiển thị trạng thái ── */
function showState(state) {
  document.getElementById('stateEmpty').style.display   = state === 'empty'   ? 'block' : 'none';
  document.getElementById('stateNone').style.display    = state === 'none'    ? 'block' : 'none';
  document.getElementById('stateLoading').style.display = state === 'loading' ? 'block' : 'none';

  const header = document.getElementById('resultsHeader');
  const isResults = state === 'results';
  header.style.display     = isResults ? 'flex' : 'none';
  resultList.style.display = isResults ? 'flex' : 'none';
  if (isResults) resultList.style.flexDirection = 'column';
}

function clearSearch() {
  input.value = '';
  clearBtn.style.display = 'none';
  resultList.innerHTML = '';
  showState('empty');
  input.focus();
}

function quickSearch(term) {
  input.value = term;
  clearBtn.style.display = 'block';
  doSearch();
}

function hplBadge(hplLevelId) { return hplBadgeHTML(hplLevelId); }

/* ── Render kết quả hóa chất ── */
function renderChemicalResults(results) {
  resultList.innerHTML = results.map(c => `
    <a class="result-item" href="chemical.html?id=${c.ChemicalID}" title="Xem hồ sơ ${c.ChemicalName}">
      <div class="result-formula">${formatFormula(c.Formula)}</div>
      <div class="result-info">
        <div class="result-name">${c.ChemicalName}</div>
        <div class="result-iupac">
          ${(c.Synonyms||[]).join(', ')}${c.CASNumber ? ' &nbsp;·&nbsp; CAS: ' + c.CASNumber : ''}${c.ChemicalGroup ? ' &nbsp;·&nbsp; ' + c.ChemicalGroup : ''}
        </div>
        <div class="result-tags">
          ${hplBadge(c.HPLLevelID)}
        </div>
      </div>
      <div class="result-arrow">›</div>
    </a>
  `).join('');
}

/* ── Render kết quả sản phẩm ── */
function renderProductResults(results) {
  resultList.innerHTML = results.map(p => {
    const cat = getCategoryById(p.CategoryID);
    return `
    <a class="result-item" href="products.html?id=${p.ProductID}" title="Xem chi tiết ${p.ProductName}">
      <div class="result-formula" style="font-size:1.6rem">${cat ? cat.icon : '📦'}</div>
      <div class="result-info">
        <div class="result-name">${p.ProductName}</div>
        <div class="result-iupac">${cat ? cat.CategoryName : ''}${p.Manufacturer ? ' &nbsp;·&nbsp; ' + p.Manufacturer : ''}</div>
        <div class="result-tags">${productHPLBadgeHTML(p)}</div>
      </div>
      <div class="result-arrow">›</div>
    </a>`;
  }).join('');
}

/* ── Thực hiện tìm kiếm ── */
function doSearch() {
  const query = input.value.trim();
  clearBtn.style.display = query ? 'block' : 'none';
  if (!query) { showState('empty'); return; }

  showState('loading');
  setTimeout(() => {
    try {
      if (currentMode === 'chemical') {
        const results = searchChemicals(query);
        if (!results.length) { showState('none'); return; }
        document.getElementById('resultsCount').innerHTML =
          `Tìm thấy <strong>${results.length}</strong> hóa chất cho "<strong>${query}</strong>"`;
        renderChemicalResults(results);
      } else {
        const results = searchProducts(query);
        if (!results.length) { showState('none'); return; }
        document.getElementById('resultsCount').innerHTML =
          `Tìm thấy <strong>${results.length}</strong> sản phẩm cho "<strong>${query}</strong>"`;
        renderProductResults(results);
      }
      showState('results');
    } catch (err) {
      console.error('[search.js] Lỗi tìm kiếm:', err);
      showState('none');
    }
  }, 150);
}

/* ── Sự kiện bàn phím ── */
input.addEventListener('keydown', e => {
  if (e.key === 'Enter') doSearch();
  if (e.key === 'Escape') clearSearch();
});

input.addEventListener('input', () => {
  clearBtn.style.display = input.value ? 'block' : 'none';
  clearTimeout(input._debounce);
  input._debounce = setTimeout(() => {
    const q = input.value.trim();
    if (q.length >= 2) doSearch();
    else if (!q) showState('empty');
  }, 350);
});

/* ── Khởi động ── */
document.addEventListener('DOMContentLoaded', function () {
  const params = new URLSearchParams(window.location.search);
  const q = params.get('q');
  const mode = params.get('mode');
  if (mode === 'product') {
    setMode(document.querySelector('.mode-tag[data-mode="product"]'), 'product');
  }
  if (q) {
    input.value = q;
    clearBtn.style.display = 'block';
    doSearch();
  } else {
    showState('empty');
  }
});
