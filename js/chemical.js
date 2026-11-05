/* ═══════════════════════════════════════════════════
   chemical.js — Trang hồ sơ hóa chất (Cấp 2) — dữ liệu thật
   ═══════════════════════════════════════════════════ */

function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[ch]));
}

function renderNotFound() {
  document.title = 'Không tìm thấy — CSDL Hóa Chất Trong Sinh Hoạt';
  document.getElementById('pageContent').innerHTML = `
    <div class="not-found-box">
      <div class="not-found-icon">🔬</div>
      <h2>Không tìm thấy hồ sơ hóa chất</h2>
      <p>ID không hợp lệ hoặc hóa chất chưa có trong cơ sở dữ liệu.</p>
      <a href="search.html" class="btn btn-primary">← Quay lại tra cứu</a>
    </div>`;
}

function renderHPLCard(c) {
  const hpl = getHPLLevel(c.HPLLevelID) || getHPLLevel(0);
  return `
    <div class="danger-card ${hpl.cssClass}">
      <div class="danger-header">
        <div class="danger-icon">${hpl.icon}</div>
        <div>
          <div class="danger-title">${hpl.LevelName}</div>
          <div class="danger-level">Mức cảnh báo: ${hpl.HPLCode}</div>
        </div>
      </div>
      ${c.HPLBasis ? `<div class="danger-text"><strong>Căn cứ:</strong> ${esc(c.HPLBasis)}</div>` : ''}
      ${c.HPLVerification ? `<div class="danger-text" style="margin-top:6px"><strong>Trạng thái xác minh:</strong> ${esc(c.HPLVerification)}</div>` : ''}
      ${c.HPLSourceURL ? `<div class="danger-text" style="margin-top:6px;word-break:break-all"><strong>Nguồn GHS/SDS:</strong> <a href="${esc(c.HPLSourceURL)}" target="_blank" rel="noopener" style="color:var(--accent)">${esc(c.HPLSourceURL)}</a></div>` : ''}
      ${c.CheckNotes ? `<div class="danger-text" style="margin-top:6px"><strong>Cần lưu ý:</strong> ${esc(c.CheckNotes)}</div>` : ''}
      ${c.HPLLevelID === 0 ? `<div class="danger-text" style="margin-top:6px"><strong>Lưu ý:</strong> chưa đủ bằng chứng GHS/SDS để phân mức. Thiếu dữ liệu không đồng nghĩa với an toàn.</div>` : ''}
      <p style="font-size:.72rem;color:var(--muted);margin-top:10px;line-height:1.5">* ${esc(HPL_SCOPE_NOTE)}</p>
    </div>`;
}

function renderGHSSection(c) {
  const pics = c.Pictograms || [];
  const hs = c.HStatements || [];
  return `
    ${pics.length ? `<div class="ghs-pictograms">${pics.map(code => {
        const p = GHS_PICTOGRAMS[code] || { name: code, icon: '⚠️' };
        return `<div class="ghs-pic" title="${code} — ${p.name}">${p.icon}</div>`;
      }).join('')}</div>` : ''}
    <div class="quick-info-row"><span class="qi-label">Signal word</span><span class="qi-value">${esc(c.SignalWord || '—')}</span></div>
    <div class="quick-info-row"><span class="qi-label">Nhóm nguy hại</span><span class="qi-value">${esc(c.HazardGroups || '—')}</span></div>
    <div class="quick-info-row"><span class="qi-label">Hạng nguy hại</span><span class="qi-value">${esc(c.HazardCategories || '—')}</span></div>
    <div class="quick-info-row"><span class="qi-label">H statements</span><span class="qi-value">${hs.length ? hs.map(esc).join(', ') : '—'}</span></div>`;
}

function renderEcotox(c) {
  const e = c.Ecotox || {};
  const rows = [
    ['LC₅₀ (cá; mg/L)', e.LC50], ['EC₅₀ (Daphnia/tảo; mg/L)', e.EC50],
    ['NOEC/EC₁₀ (mg/L)', e.NOEC], ['log Kow', e.logKow], ['BCF', e.BCF],
    ['Phân hủy sinh học', e.Biodegradation], ['Aquatic Acute/Chronic (GHS)', e.AquaticClass],
    ['Đánh giá độc tính sinh thái', e.Assessment], ['Chất lượng/trạng thái dữ liệu', e.DataQuality],
  ];
  const src = e.EcoSource ? e.EcoSource.split(/\s*\|\s*/).filter(Boolean) : [];
  return `
    <div class="info-card">
      <div class="section-label">🌊 Dữ liệu độc tính sinh thái</div>
      ${rows.map(([k, v]) => `
        <div class="quick-info-row" style="align-items:flex-start;gap:12px">
          <span class="qi-label" style="flex:0 0 38%">${k}</span>
          <span class="qi-value" style="text-align:right;${isDataAvailable(v) ? '' : 'color:var(--muted)'}">${esc(v || '—')}</span>
        </div>`).join('')}
      <p style="font-size:.74rem;color:var(--muted);margin-top:10px;line-height:1.5">
        “NF” = chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên — không đồng nghĩa hóa chất không độc.</p>
      ${src.length ? `<div style="margin-top:8px;font-size:.78rem;line-height:1.6">Nguồn: ${src.map(s =>
        /^https?:\/\//.test(s) ? `<a href="${esc(s)}" target="_blank" rel="noopener" style="color:var(--accent)">${esc(s)}</a>` : esc(s)).join('<br/>')}</div>` : ''}
    </div>`;
}

function renderProducts(chemId) {
  const links = getProductsByChemicalId(chemId);
  if (!links.length) return `<p style="color:var(--muted);font-size:.9rem;padding:12px 0">Chưa có sản phẩm nào được liên kết với hóa chất này.</p>`;
  return `<div style="display:flex;flex-direction:column;gap:8px">
    ${links.map(({ product: p, LabelName, LinkConfidence }) => {
      const trusted = ['BỘ GỐC – ĐÃ GẮN','XÁC MINH TỪ THÀNH PHẦN'].includes(LinkConfidence);
      const cat = getCategoryById(p.CategoryID);
      return `
      <a href="products.html?id=${encodeURIComponent(p.ProductID)}" class="danger-point"
         style="text-decoration:none;color:inherit;border:1px solid var(--border);border-radius:12px;padding:12px 14px;">
        <span class="danger-point-icon">${cat ? cat.icon : '🛒'}</span>
        <span><strong>${esc(p.ProductName)}</strong><br/>
          <span style="font-size:.78rem;color:var(--muted)">${cat ? esc(cat.CategoryName) : ''}${LabelName ? ' · Tên trên nhãn: ' + esc(LabelName) : ''}${LinkConfidence && !trusted ? ' · <span style="color:#b45309">⚠ ' + esc(LinkConfidence.toLowerCase()) + '</span>' : ''}</span>
        </span>
      </a>`;
    }).join('')}
  </div>`;
}

function renderResources(chemId) {
  const resources = getResourcesByChemicalId(chemId);
  if (!resources.length) return `<p style="color:var(--muted);font-size:.9rem;padding:12px 0">Chưa có học liệu số cho hóa chất này.</p>`;
  return `<div style="display:flex;flex-direction:column;gap:10px">${resources.map(r => {
    const t = r.resourceType || {};
    const isInfo = r.ResourceTypeID === 2;
    const href = isInfo ? r.FileURL : resourceOpenUrl(r);
    return `
    <a href="${esc(href)}" target="_blank" rel="noopener" class="danger-point"
       style="text-decoration:none;color:inherit;border:1px solid var(--border);border-radius:12px;padding:12px 14px;flex-wrap:wrap;">
      <span class="danger-point-icon">${t.icon || '📄'}</span>
      <span><strong>${esc(r.Title)}</strong><br/>
        <span style="font-size:.78rem;color:var(--muted)">${esc(t.TypeName || '')}${r.Description ? ' — ' + esc(r.Description) : ''}</span></span>
      ${isInfo ? `<img src="${esc(r.FileURL)}" alt="${esc(r.Title)}" loading="lazy" style="width:100%;max-height:340px;object-fit:contain;border-radius:8px;background:#fff;margin-top:8px">` : ''}
    </a>`;
  }).join('')}</div>`;
}

function renderReferences(c) {
  if (!c.References || !c.References.length) return '';
  return `
    <div class="info-card">
      <div class="section-label">🔗 Nguồn tài liệu tham khảo</div>
      <ul style="margin:0;padding-left:1.1em;display:flex;flex-direction:column;gap:6px;">
        ${c.References.map(r => {
          const m = r.match(/^(.*?)(https?:\/\/\S+)$/);
          return `<li style="font-size:.85rem;line-height:1.6">${m
            ? `${esc(m[1])}<a href="${esc(m[2])}" target="_blank" rel="noopener" style="color:var(--accent)">${esc(m[2])}</a>`
            : esc(r)}</li>`;
        }).join('')}
      </ul>
    </div>`;
}

function renderChemicalPage(c) {
  document.title = `${c.ChemicalName} — CSDL Hóa Chất Trong Sinh Hoạt`;
  const hpl = getHPLLevel(c.HPLLevelID) || getHPLLevel(0);
  const synonyms = (c.Synonyms || []).join(' · ');
  const nProducts = getProductsByChemicalId(c.ChemicalID).length;

  document.getElementById('pageContent').innerHTML = `
    <div class="breadcrumb">
      <a href="index.html">Trang chủ</a><span>›</span>
      <a href="chemicals.html">Hóa chất</a><span>›</span>
      <span>${esc(c.ChemicalName)}</span>
    </div>

    <div class="chem-header">
      <div class="chem-icon">⚗️</div>
      <div class="chem-header-info">
        <div class="chem-formula-display">${formatFormula(esc(c.Formula || ''))}</div>
        <div class="chem-name-display">${esc(c.ChemicalName)}</div>
        <div class="chem-name-alt">${esc(synonyms)}</div>
        <div class="chem-badges">
          <span class="badge" style="background:${hpl.ColorCode}22;color:${hpl.ColorCode};border-color:${hpl.ColorCode}55">${hpl.icon} ${hpl.HPLCode} — ${hpl.LevelName}</span>
          ${c.CASNumber ? `<span class="badge badge-muted">CAS: ${esc(c.CASNumber)}</span>` : ''}
          ${c.ChemicalGroup ? `<span class="badge badge-accent">${esc(c.ChemicalGroup)}</span>` : ''}
        </div>
      </div>
    </div>

    ${c.IsNew ? `<div class="notice-box"><strong>Hóa chất bổ sung (03/11/2026).</strong> Hồ sơ này mới xác minh sự có mặt trong sản phẩm;
      chưa thuộc bộ 143 hóa chất đã phân mức HPL của bài báo. Các trường GHS, ảnh hưởng sức khỏe/môi trường và độc tính sinh thái
      có thể còn “chưa rà soát”, thiếu dữ liệu không đồng nghĩa với an toàn.</div>` : ''}

    <div class="chem-layout">
      <div class="col-left">
        <div class="info-card">
          <div class="section-label">🩺 Ảnh hưởng đến sức khỏe</div>
          <p style="${c.HealthEffects ? '' : 'color:var(--muted)'}">${esc(c.HealthEffects || (c.IsNew ? 'Chưa rà soát.' : 'Chưa có thông tin.'))}</p>
        </div>
        <div class="info-card">
          <div class="section-label">🌍 Ảnh hưởng đến môi trường</div>
          <p style="${c.EnvEffects ? '' : 'color:var(--muted)'}">${esc(c.EnvEffects || (c.IsNew ? 'Chưa rà soát.' : 'Chưa có thông tin.'))}</p>
        </div>

        ${renderEcotox(c)}

        <div class="products-section">
          <div class="section-label">🛒 Sản phẩm có chứa hóa chất này (${nProducts})</div>
          ${renderProducts(c.ChemicalID)}
        </div>

        <div class="products-section">
          <div class="section-label">🎓 Học liệu số liên quan</div>
          ${renderResources(c.ChemicalID)}
        </div>

        ${renderReferences(c)}

        ${c.StatusNote ? `<div class="info-card"><div class="section-label">📝 Trạng thái / ghi chú dữ liệu</div><p>${esc(c.StatusNote)}</p></div>` : ''}
      </div>

      <div class="col-right">
        ${renderHPLCard(c)}

        <div class="quick-info-card">
          <div style="padding:14px 18px 10px;border-bottom:1px solid var(--border);">
            <span style="font-size:.75rem;font-weight:600;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);">☣️ Phân loại GHS</span>
          </div>
          <div style="padding:6px 18px 14px">${renderGHSSection(c)}</div>
        </div>

        <div class="quick-info-card">
          <div style="padding:14px 18px 10px;border-bottom:1px solid var(--border);">
            <span style="font-size:.75rem;font-weight:600;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);">📋 Thông tin nhanh</span>
          </div>
          <div class="quick-info-row"><span class="qi-label">Mã hồ sơ</span><span class="qi-value">${esc(c.ChemicalID)}</span></div>
          <div class="quick-info-row"><span class="qi-label">Số CAS</span><span class="qi-value">${esc(c.CASNumber || '—')}</span></div>
          <div class="quick-info-row"><span class="qi-label">Công thức</span><span class="qi-value" style="font-family:var(--font-mono)">${formatFormula(esc(c.Formula || '')) || '—'}</span></div>
          <div class="quick-info-row"><span class="qi-label">Nhóm hóa chất</span><span class="qi-value">${esc(c.ChemicalGroup || '—')}</span></div>
          <div class="quick-info-row"><span class="qi-label">Nhóm SP có thể chứa</span><span class="qi-value">${esc(c.ProductCategoriesText || '—')}</span></div>
        </div>

        <div style="display:flex;flex-direction:column;gap:10px;">
          ${c.CASNumber ? `<a href="https://pubchem.ncbi.nlm.nih.gov/#query=${encodeURIComponent(c.CASNumber)}" target="_blank" rel="noopener" class="btn btn-outline" style="justify-content:center;">🌐 Tra cứu thêm tại PubChem</a>` : ''}
          <button onclick="goBackSafe()" class="btn btn-ghost" style="justify-content:center;cursor:pointer;">← Quay lại</button>
        </div>
      </div>
    </div>`;
}

document.addEventListener('DOMContentLoaded', function () {
  try {
    const id = new URLSearchParams(window.location.search).get('id');
    const chem = id ? getChemicalById(id) : null;
    if (!chem) { renderNotFound(); return; }
    renderChemicalPage(chem);
  } catch (err) {
    console.error('[chemical.js]', err);
    document.getElementById('pageContent').innerHTML = `<div style="text-align:center;padding:80px 0;color:var(--danger)"><h2>Lỗi tải trang</h2><p style="color:var(--muted)">${esc(err.message)}</p></div>`;
  }
});

function goBackSafe() {
  if (document.referrer && document.referrer !== window.location.href) history.back();
  else window.location.href = 'chemicals.html';
}
