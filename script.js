function injectAge(row) {
  if (row.DBirthDate) {
    const deathDate = excelDateTimeToJSDate(row.DDeathDate);
    const birthDate = excelDateTimeToJSDate(row.DBirthDate);
    if (birthDate === 'Invalid Date') return row;

    const use =  new calculateDateDifference(deathDate, birthDate);
    const diff = use.jsBuiltIn();
    console.log({diff})
    return {
      DBirthDate: { t: 'd', v: birthDate },
      DDeathDate: { t: 'd', v: deathDate },
      Y: { t: 'n', v: Math.floor(diff.years) },
      M: { t: 'n', v: Math.floor(diff.months) },
      D: { t: 'n', v: Math.floor(diff.days) },
    };
  }
  return row;
}


document.addEventListener('DOMContentLoaded', () => {
  const fileInput = document.getElementById('fileInput');
  const processBtn = document.getElementById('process-btn');
  const state = {
    workbook: null,
    set(key, value) {
      this[key] = value;
    },
    get(key) {
      return this[key] || null;
    },
  };

  fileInput.addEventListener('change', readExcelFile(state));
  processBtn.addEventListener('click', processExcel(state));
});
