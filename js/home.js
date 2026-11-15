(() => {
  const list = document.getElementById('schoolList');
  const count = document.getElementById('schoolCount');
  const search = document.getElementById('schoolSearch');
  const schoolFilter = document.getElementById('schoolFilter');
  const status = document.getElementById('directoryStatus');
  let allSchools = [];
  document.getElementById('year').textContent = new Date().getFullYear() + 543;
  renderAuthenticatedMenu();

  function renderAuthenticatedMenu() {
    let user = null;
    try { user = JSON.parse(localStorage.getItem('nited_user') || 'null'); } catch (_) { user = null; }
    if (!user) return;

    const nav = document.querySelector('.topbar nav');
    nav.classList.add('authenticated');
    document.querySelector('.topbar').classList.add('authenticated');
    const role = String(user.role || '').trim().toLowerCase();
    const privileged = ['admin', 'supervisor'].includes(role);
    const links = [
      ['pages/dashboard.html', '&#127968; แดชบอร์ด'],
      ['pages/files.html', '&#128194; หลักฐานโรงเรียน'],
      ['pages/reports.html', '&#128202; รายงานโรงเรียน']
    ];
    if (role === 'admin') links.push(['pages/admin.html', '&#9881; ผู้ดูแลระบบ']);
    const appointment = `<div class="home-nav-dropdown"><a href="pages/booking.html">&#128197; กำหนดนิเทศ <span aria-hidden="true">⌄</span></a>${privileged ? '<div class="home-nav-submenu"><a href="pages/evaluation.html">&#128221; ทำแบบนิเทศ</a></div>' : ''}</div>`;
    nav.innerHTML = links.slice(0,1).map(([href,label]) => `<a href="${href}">${label}</a>`).join('') + appointment + links.slice(1).map(([href,label]) => `<a href="${href}">${label}</a>`).join('') + '<a class="nav-button" href="#" id="homeLogout">&#10140; ออกจากระบบ</a>';
    document.getElementById('homeLogout').addEventListener('click', event => {
      event.preventDefault();
      try {
        const activeUser = JSON.parse(localStorage.getItem('nited_user') || 'null');
        if (activeUser && activeUser.token && CONFIG.APPS_SCRIPT_URL) {
          fetch(CONFIG.APPS_SCRIPT_URL, { method:'POST', mode:'cors', redirect:'follow', keepalive:true,
            headers:{'Content-Type':'text/plain'}, body:JSON.stringify({action:'logout',token:activeUser.token}) });
        }
      } catch (_) {}
      localStorage.removeItem('nited_user');
      window.location.reload();
    });
  }

  function render() {
    const term = search.value.trim().toLocaleLowerCase('th');
    let signedUser = null;
    try { signedUser = JSON.parse(localStorage.getItem('nited_user') || 'null'); } catch (_) {}
    const visible = allSchools.filter(s => {
      if (signedUser && signedUser.role === 'school' && String(s.schoolId) !== String(signedUser.schoolId)) return false;
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
    let loggedIn = null;
    try { loggedIn = JSON.parse(localStorage.getItem('nited_user') || 'null'); } catch (_) {}
    const schoolOptions = loggedIn && loggedIn.role === 'school'
      ? schools.filter(school => String(school.schoolId) === String(loggedIn.schoolId))
      : schools;
    schoolOptions.forEach(school => {
      const option = document.createElement('option');
      option.value = school.schoolId;
      option.textContent = `${school.name || 'ไม่ระบุชื่อโรงเรียน'}${school.code ? ` (${school.code})` : ''}`;
      schoolFilter.appendChild(option);
    });
    try {
      if (loggedIn && loggedIn.role === 'school') {
        schoolFilter.value = String(loggedIn.schoolId || '');
        schoolFilter.disabled = true;
        search.disabled = true;
        search.placeholder = 'แสดงข้อมูลเฉพาะโรงเรียนของคุณ';
      }
    } catch (_) {}
    render();
    status.textContent = 'ข้อมูลโรงเรียนจากชีต Schools · แสดงผู้บริหาร ที่อยู่ เว็บไซต์ และพิกัดเมื่อมีข้อมูล';
  }).catch(() => {
    list.innerHTML = '<div class="directory-state">ยังโหลดรายชื่อโรงเรียนไม่ได้ โปรดลองใหม่ภายหลัง</div>';
    status.textContent = 'ตรวจสอบการเชื่อมต่อ Apps Script';
  });

  const brandMark = document.querySelector('.brand-mark:not(.brand-logo-image)');
  if (brandMark) {
    const image = document.createElement('img');
    image.className = 'brand-mark brand-logo-image';
    image.alt = 'ตราสัญลักษณ์ สพม.นราธิวาส';
    image.src = 'https://www.sesaonara.go.th/wp-content/uploads/elementor/thumbs/smpnt-web-header3-pm81vc0p83fwfylect5mcnanbnw7ct0kusm9zvfjqe.png';
    image.onerror = () => { const fallback = document.createElement('span'); fallback.className = 'brand-mark'; fallback.textContent = 'น'; image.replaceWith(fallback); };
    brandMark.replaceWith(image);
  }
})();
