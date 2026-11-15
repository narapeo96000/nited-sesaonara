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
      const matchesSearch = `${s.name || ''} ${s.district || ''} ${s.code || ''} ${s.address || ''}`.toLocaleLowerCase('th').includes(term);
      const matchesSchool = !schoolFilter.value || String(s.schoolId) === schoolFilter.value;
      return matchesSearch && matchesSchool;
    });
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

  search.addEventListener('input', render);
  schoolFilter.addEventListener('change', render);
  loadSchoolDirectory().then(schools => {
    allSchools = schools;
    schools.forEach(school => {
      const option = document.createElement('option');
      option.value = school.schoolId;
      option.textContent = school.name || 'ไม่ระบุชื่อโรงเรียน';
      schoolFilter.appendChild(option);
    });
    render();
    status.textContent = 'ข้อมูลชื่อโรงเรียนและที่ตั้งจากชีต Schools';
  }).catch(() => {
    list.innerHTML = '<div class="directory-state">ยังโหลดรายชื่อโรงเรียนไม่ได้ โปรดลองใหม่ภายหลัง</div>';
    status.textContent = 'ตรวจสอบการเชื่อมต่อ Apps Script';
  });
})();
