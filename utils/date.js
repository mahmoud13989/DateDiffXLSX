function excelDateTimeToJSDate(excelDateTimeValue) {
  if (!excelDateTimeValue) {
    return 'Invalid Date';
  }
  if (typeof excelDateTimeValue != 'number') {
    return 'Invalid Date';
  }

  const parsedDate = XLSX.SSF.parse_date_code(excelDateTimeValue);
  const result = new Date(parsedDate.y, parsedDate.m, parsedDate.d);

  return result;
}

function diffLuxonDates(greaterJSDate, smallerJSDate) {
  const luxonGreaterJSDate = luxon.DateTime.fromJSDate(greaterJSDate);
  const luxonSmallerJSDate = luxon.DateTime.fromJSDate(smallerJSDate);

  return luxonGreaterJSDate
    .diff(luxonSmallerJSDate, ['years', 'months', 'days'])
    .toObject();
}
