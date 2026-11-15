/* Shared school directory loader for GitHub Pages and Apps Script frontends. */
let schoolDirectoryRequest = null;

function loadSchoolDirectory() {
  if (schoolDirectoryRequest) return schoolDirectoryRequest;

  schoolDirectoryRequest = new Promise((resolve, reject) => {
    const callback = `nitedDirectory_${Date.now()}_${Math.floor(Math.random() * 100000)}`;
    const script = document.createElement('script');
    let finished = false;

    const cleanup = () => {
      if (finished) return;
      finished = true;
      clearTimeout(timeout);
      delete window[callback];
      script.remove();
    };

    const timeout = setTimeout(() => {
      cleanup();
      reject(new Error('หมดเวลารอรายชื่อโรงเรียนจากชีต Schools'));
    }, 15000);

    window[callback] = result => {
      cleanup();
      if (!result || !result.success || !Array.isArray(result.schools)) {
        reject(new Error(result && result.message || 'โหลดรายชื่อโรงเรียนไม่สำเร็จ'));
        return;
      }
      resolve(result.schools);
    };

    script.onerror = () => {
      cleanup();
      reject(new Error('เชื่อมต่อ Apps Script เพื่อโหลดรายชื่อโรงเรียนไม่สำเร็จ'));
    };
    script.src = `${CONFIG.APPS_SCRIPT_URL}?api=publicSchools&callback=${encodeURIComponent(callback)}`;
    document.head.appendChild(script);
  }).catch(error => {
    schoolDirectoryRequest = null;
    throw error;
  });

  return schoolDirectoryRequest;
}

function populateSchoolDropdown(target, placeholder, selectedId) {
  const select = typeof target === 'string' ? document.getElementById(target) : target;
  if (!select) return Promise.reject(new Error('ไม่พบ dropdown โรงเรียน'));

  return loadSchoolDirectory().then(schools => {
    const activeUser = typeof APP !== 'undefined' ? APP.user : null;
    const visibleSchools = activeUser && activeUser.role === 'school'
      ? schools.filter(school => String(school.schoolId) === String(activeUser.schoolId))
      : schools;
    const selected = selectedId || select.value;
    select.replaceChildren();

    const first = document.createElement('option');
    first.value = '';
    first.textContent = placeholder || '-- เลือกโรงเรียน --';
    select.appendChild(first);

    visibleSchools.forEach(school => {
      const option = document.createElement('option');
      option.value = school.schoolId;
      option.textContent = school.name || 'ไม่ระบุชื่อโรงเรียน';
      option.dataset.code = school.code || '';
      option.dataset.district = school.district || '';
      option.dataset.address = school.address || '';
      select.appendChild(option);
    });

    if (selected) select.value = selected;
    if (activeUser && activeUser.role === 'school') select.disabled = true;
    return visibleSchools;
  });
}
