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

// function calculateDateDifference(greaterJSDate, smallerJSDate) {
//   return function () {
//     const luxonGreaterJSDate = luxon.DateTime.fromJSDate(greaterJSDate, {
//       zone: 'Africa/Cairo',
//     });
//     const luxonSmallerJSDate = luxon.DateTime.fromJSDate(smallerJSDate, {
//       zone: 'Africa/Cairo',
//     });

//     return luxonGreaterJSDate
//       .diff(luxonSmallerJSDate, ['years', 'months', 'days'])
//       .toObject();
//   };
// }

class calculateDateDifference {
  constructor(greaterJSDate, smallerJSDate) {
    this.greaterJSDate = greaterJSDate;
    this.smallerJSDate = smallerJSDate;
  }
  luxon() {
    if (!luxon) {
      console.log('luxon is not defined properly');
    }
    const luxonGreaterJSDate = luxon.DateTime.fromJSDate(this.greaterJSDate, {
      zone: 'Africa/Cairo',
    });
    const luxonSmallerJSDate = luxon.DateTime.fromJSDate(this.smallerJSDate, {
      zone: 'Africa/Cairo',
    });

    return luxonGreaterJSDate
      .diff(luxonSmallerJSDate, ['years', 'months', 'days'])
      .toObject();
  }
  jsBuiltIn() {
    const diffTime = Math.abs(this.greaterJSDate - this.smallerJSDate);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

    const years = Math.floor(diffDays / 365.25); // Adjust for leap years
    const remainingDaysAfterYears = diffDays % 365.25;
    const months = Math.floor(remainingDaysAfterYears / 30.15); // Average days in a month
    const days = Math.floor(remainingDaysAfterYears % 30.15);

    return { years, months, days };
  }
}
