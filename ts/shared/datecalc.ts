function isLeapYear(year: number): boolean {
    return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

function getDaysInMonth(year: number, month: number) {
    return [31, isLeapYear(year) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month];
}

// Note that adding months isn't straight forward because of edge cases round leap years, and
// months with differing number of days.
// The test cases below are consistent with the adding of months by the D365 AddMonths calculated
// field function, Excel and moment.js.  Note moment.js not used because it's a large bundle.
//
// TEST	                       INPUT        OUTPUT (if 2 months added)
// Start of Month              01/09/2019	01/11/2019
// End of leap Year	           31/12/2020	28/02/2021
// End of non - leap Year	   31/12/2019	29/02/2020
// End of June	               30/06/2019	30/08/2019
// End of July	               31/07/2019	30/09/2019
// Leap Year End of Feb	       29/02/2020	29/04/2020
// Non Leap Year End of Feb	   28/02/2019	28/04/2019
export function addMonths(date: Date, value: number): Date {
    const d = new Date(date);
    const n = date.getDate();
    d.setDate(1);
    d.setMonth(d.getMonth() + value);
    d.setDate(Math.min(n, getDaysInMonth(d.getFullYear(), d.getMonth())));
    return d;
}

export function addYears(dt: Date, years: number): Date {
    const result = new Date(dt);
    result.setFullYear(result.getFullYear() + years);
    return result;
}
