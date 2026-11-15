(() => {
  const list = document.getElementById('schoolList');
  const count = document.getElementById('schoolCount');
  const search = document.getElementById('schoolSearch');
  const status = document.getElementById('directoryStatus');
  let allSchools = [];
  document.getElementById('year').textContent = new Date().getFullYear() + 543;

  function render() {
    const term = search.value.trim().toLocaleLowerCase('th');
    const visible = allSchools.filter(s => `${s.name || ''} ${s.district || ''} ${s.code || ''} ${s.address || ''}`.toLocaleLowerCase('th').includes(term));
    count.textContent = visible.length;
    if (!visible.length) {
      list.innerHTML = '<div class="directory-state">ไม่พบโรงเรียนที่ตรงกับคำค้น</div>';
      return;
    }
    list.innerHTML = visible.map(s => `<article class="school-card"><span><b>${escapeHtml(s.name || 'ไม่ระบุชื่อ')}</b><small>${escapeHtml(s.district || 'จังหวัดนราธิวาส')}${s.code ? ` · รหัส ${escapeHtml(s.code)}` : ''}</small>${s.address ? `<small class="school-address"><span aria-hidden="true">⌖</span> ${escapeHtml(s.address)}</small>` : ''}</span><span class="school-arrow" aria-hidden="true">↗</span></article>`).join('');
  }

  function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
  }

  function loadSchools() {
    const callback = `nitedSchools_${Date.now()}_${Math.floor(Math.random() * 100000)}`;
    const script = document.createElement('script');
    const timeout = setTimeout(() => finish(new Error('หมดเวลารอข้อมูล')), 12000);
    function finish(error, payload) {
      clearTimeout(timeout);
      delete window[callback];
      script.remove();
      if (error || !payload || !payload.success || !Array.isArray(payload.schools)) {
        list.innerHTML = '<div class="directory-state">ยังโหลดข้อมูลไม่ได้ โปรดลองใหม่ภายหลัง หรือติดต่อสำนักงาน</div>';
        status.textContent = 'รายชื่อแสดงเมื่อบริการชีตพร้อมใช้งาน';
        return;
      }
      allSchools = payload.schools;
      render();
      status.textContent = 'ข้อมูลรายชื่อจากชีต Schools · แสดงเฉพาะข้อมูลสาธารณะ';
    }
    window[callback] = payload => finish(null, payload);
    script.onerror = () => finish(new Error('เชื่อมต่อไม่สำเร็จ'));
    script.src = `${CONFIG.APPS_SCRIPT_URL}?api=publicSchools&callback=${encodeURIComponent(callback)}`;
    document.head.appendChild(script);
  }

  search.addEventListener('input', render);
  loadSchools();
})();
