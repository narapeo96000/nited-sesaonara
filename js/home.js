(() => {
  const list = document.getElementById('schoolList');
  const count = document.getElementById('schoolCount');
  const search = document.getElementById('schoolSearch');
  const schoolFilter = document.getElementById('schoolFilter');
  const status = document.getElementById('directoryStatus');
  let allSchools = [];
  document.getElementById('year').textContent = new Date().getFullYear() + 543;

  function render() {
    const term = search.value.trim().toLocaleLowerCase('th');
    const visible = allSchools.filter(s => {
      const matchesSearch = `${s.name || ''} ${s.district || ''} ${s.code || ''} ${s.director || ''} ${s.address || ''}`.toLocaleLowerCase('th').includes(term);
      const matchesSchool = !schoolFilter.value || String(s.schoolId) === schoolFilter.value;
      return matchesSearch && matchesSchool;
    });
    count.textContent = visible.length;
    if (!visible.length) {
      list.innerHTML = '<div class="directory-state">ไม่พบโรงเรียนที่ตรงกับคำค้น</div>';
      return;
    }
    list.innerHTML = visible.map(s => {
      const website = normalizeWebsite(s.website);
      const position = s.latitude && s.longitude ? `${s.latitude},${s.longitude}` : s.coordinates || '';
      const directMapUrl = /^https?:\/\//i.test(String(s.coordinates || '').trim()) ? normalizeWebsite(s.coordinates) : '';
      const mapUrl = directMapUrl || (position ? `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(position)}` : '');
      return `<article class="school-card"><div class="school-card-content"><b>${escapeHtml(s.name || 'ไม่ระบุชื่อ')}</b><small class="school-code">รหัสโรงเรียน ${escapeHtml(s.code || '—')}</small>${s.director ? `<small>ผู้บริหาร: ${escapeHtml(s.director)}</small>` : ''}${s.address ? `<small class="school-address"><span aria-hidden="true">⌖</span> ${escapeHtml(s.address)}</small>` : ''}<div class="school-links">${website ? `<a href="${escapeHtml(website)}" target="_blank" rel="noopener noreferrer">เว็บไซต์ ↗</a>` : ''}${mapUrl ? `<a href="${mapUrl}" target="_blank" rel="noopener noreferrer">ดูพิกัดแผนที่ ↗</a>` : ''}</div></div></article>`;
    }).join('');
  }

  function normalizeWebsite(value) {
    const input = String(value || '').trim();
    if (!input) return '';
    const url = /^https?:\/\//i.test(input) ? input : `https://${input}`;
    try {
      const parsed = new URL(url);
      return ['http:', 'https:'].includes(parsed.protocol) ? parsed.href : '';
    } catch (_) { return ''; }
  }

  function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
  }

  search.addEventListener('input', render);
  schoolFilter.addEventListener('change', render);
  loadSchoolDirectory().then(schools => {
    allSchools = schools;
    schools.forEach(school => {
      const option = document.createElement('option');
      option.value = school.schoolId;
      option.textContent = `${school.name || 'ไม่ระบุชื่อโรงเรียน'}${school.code ? ` (${school.code})` : ''}`;
      schoolFilter.appendChild(option);
    });
    render();
    status.textContent = 'ข้อมูลโรงเรียนจากชีต Schools · แสดงผู้บริหาร ที่อยู่ เว็บไซต์ และพิกัดเมื่อมีข้อมูล';
  }).catch(() => {
    list.innerHTML = '<div class="directory-state">ยังโหลดรายชื่อโรงเรียนไม่ได้ โปรดลองใหม่ภายหลัง</div>';
    status.textContent = 'ตรวจสอบการเชื่อมต่อ Apps Script';
  });
})();
