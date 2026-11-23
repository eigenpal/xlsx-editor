// Date serials in the two date systems (ECMA-376 Part 1, 18.17.4.1) and the date and time
// sections of number formats. `x:workbookPr/@date1904` (`CT_WorkbookPr`, sml.xsd, default
// false) selects the 1904 system. The 1900 system counts 1900 as a leap year: serial 60 is
// 1900-02-29, and serial 0 shows as 1900-01-00. The 1904 system counts from 1904-01-01. Time
// rounds to the shown precision (whole seconds, or the fractional seconds of the code), and the
// rounding carries into the minutes, hours, and date.

import type { Section } from '../../store/number-format/parse.ts';
import type { Pad, Rendered, Separators } from './render-number.ts';

/**
 * The date of a serial: year, month 1 to 12, day 0 to 31, and weekday 0 (Sunday) to 6. Day 0
 * is only serial 0 of the 1900 system (1900-01-00).
 */
export interface CalendarDate {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  /** 0 is Sunday. */
  readonly weekday: number;
}

/** The largest serial that is a date: 9999-12-31. */
const MAX_SERIAL_1900 = 2958465;
const MAX_SERIAL_1904 = MAX_SERIAL_1900 - 1462;

/** Days from 1970-01-01 to a proleptic Gregorian date. */
function daysFromCivil(y: number, m: number, d: number): number {
  const yy = m <= 2 ? y - 1 : y;
  const era = Math.floor(yy / 400);
  const yoe = yy - era * 400;
  const doy = Math.floor((153 * (m + (m > 2 ? -3 : 9)) + 2) / 5) + d - 1;
  const doe = yoe * 365 + Math.floor(yoe / 4) - Math.floor(yoe / 100) + doy;
  return era * 146097 + doe - 719468;
}

function civilFromDays(z: number): { year: number; month: number; day: number } {
  const zz = z + 719468;
  const era = Math.floor(zz / 146097);
  const doe = zz - era * 146097;
  const yoe = Math.floor(
    (doe - Math.floor(doe / 1460) + Math.floor(doe / 36524) - Math.floor(doe / 146096)) / 365
  );
  const doy = doe - (365 * yoe + Math.floor(yoe / 4) - Math.floor(yoe / 100));
  const mp = Math.floor((5 * doy + 2) / 153);
  const day = doy - Math.floor((153 * mp + 2) / 5) + 1;
  const month = mp < 10 ? mp + 3 : mp - 9;
  return { year: yoe + era * 400 + (month <= 2 ? 1 : 0), month, day };
}

const EPOCH_1900 = daysFromCivil(1899, 12, 30);
const EPOCH_1904 = daysFromCivil(1904, 1, 1);

/**
 * The calendar date of a whole serial day, or undefined outside 0 to 9999-12-31. In the 1900
 * system, serial 60 is 1900-02-29, a day that the 1900 system counts.
 */
export function serialToDate(serial: number, date1904: boolean): CalendarDate | undefined {
  const day = Math.floor(serial);
  if (!Number.isFinite(serial) || day < 0 || day > (date1904 ? MAX_SERIAL_1904 : MAX_SERIAL_1900))
    return undefined;
  if (date1904) {
    const z = EPOCH_1904 + day;
    return { ...civilFromDays(z), weekday: mod(z + 4, 7) };
  }
  // The 1900 system: serial 1 is a Sunday, and every serial keeps that weekday count.
  const weekday = mod(day + 6, 7);
  if (day === 0) return { year: 1900, month: 1, day: 0, weekday };
  if (day === 60) return { year: 1900, month: 2, day: 29, weekday };
  return { ...civilFromDays(EPOCH_1900 + (day < 60 ? day + 1 : day)), weekday };
}

const mod = (a: number, b: number): number => ((a % b) + b) % b;

/**
 * The serial of an ISO 8601 date or date-time (`2024-03-01`, `2024-03-01T12:30:00`, an optional
 * fraction of a second and `Z`), or undefined when the text is not a valid date in the system.
 */
export function dateToSerial(iso: string, date1904: boolean): number | undefined {
  if (iso.length < 10 || iso.length > 40) return undefined;
  const year = digitsAt(iso, 0, 4);
  const month = digitsAt(iso, 5, 2);
  const day = digitsAt(iso, 8, 2);
  if (
    iso[4] !== '-' ||
    iso[7] !== '-' ||
    year === undefined ||
    month === undefined ||
    day === undefined
  )
    return undefined;
  if (month < 1 || month > 12 || day < 1 || day > daysInMonth(year, month)) return undefined;
  if (year < (date1904 ? 1904 : 1900)) return undefined;
  let fraction = 0;
  if (iso.length > 10) {
    const time = readTime(iso.slice(11));
    if ((iso[10] !== 'T' && iso[10] !== ' ') || time === undefined) return undefined;
    fraction = time;
  }
  const z = daysFromCivil(year, month, day);
  let serial: number;
  if (date1904) serial = z - EPOCH_1904;
  else {
    serial = z - EPOCH_1900;
    // Before 1900-03-01 the 1900 system is one day ahead of a plain count.
    if (serial < 61) serial -= 1;
  }
  if (serial < 0) return undefined;
  return serial + fraction;
}

function daysInMonth(year: number, month: number): number {
  if (month === 2) return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0 ? 29 : 28;
  return month === 4 || month === 6 || month === 9 || month === 11 ? 30 : 31;
}

function digitsAt(text: string, at: number, count: number): number | undefined {
  let value = 0;
  for (let k = at; k < at + count; k++) {
    const d = text.charCodeAt(k) - 48;
    if (!(d >= 0 && d <= 9)) return undefined;
    value = value * 10 + d;
  }
  return value;
}

/** `HH:MM[:SS[.fff]][Z]` as a fraction of a day. */
function readTime(text: string): number | undefined {
  let body = text.endsWith('Z') ? text.slice(0, -1) : text;
  const h = digitsAt(body, 0, 2);
  const m = digitsAt(body, 3, 2);
  if (body[2] !== ':' || h === undefined || m === undefined || h > 23 || m > 59) return undefined;
  let s = 0;
  if (body.length > 5) {
    const ss = digitsAt(body, 6, 2);
    if (body[5] !== ':' || ss === undefined || ss > 59) return undefined;
    s = ss;
    body = body.slice(8);
    if (body.length > 0) {
      if (body[0] !== '.' || body.length < 2) return undefined;
      const frac = digitsAt(body, 1, body.length - 1);
      if (frac === undefined) return undefined;
      s += frac / 10 ** (body.length - 1);
    }
  } else if (body.length !== 5) return undefined;
  return (h * 3600 + m * 60 + s) / 86400;
}

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
]; // prettier-ignore
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

const two = (n: number): string => String(n).padStart(2, '0');

/**
 * Renders a non-negative serial through a date section, or undefined when the serial is not a
 * date (negative or past 9999-12-31): the caller shows the error fill.
 */
export function renderDate(
  serial: number,
  section: Section,
  sep: Separators,
  date1904: boolean
): Rendered | undefined {
  if (!(serial >= 0)) return undefined;
  const unit = 10 ** section.subsec;
  const perDay = 86400 * unit;
  const ticks = Math.round(Number((serial * perDay).toPrecision(15)));
  const days = Math.floor(ticks / perDay);
  const date = serialToDate(days, date1904);
  if (!date) return undefined;
  const inDay = ticks - days * perDay;
  const sub = inDay % unit;
  const secs = Math.floor(inDay / unit);
  const hour = Math.floor(secs / 3600);
  const minute = Math.floor(secs / 60) % 60;
  const second = secs % 60;
  const totalSecs = Math.floor(ticks / unit);
  let text = '';
  let fill: Rendered['fill'];
  const pads: Pad[] = [];
  for (const t of section.tokens) {
    switch (t.kind) {
      case 'lit':
        if (t.pad !== undefined) pads.push({ index: text.length, char: t.pad });
        text += t.text;
        break;
      case 'fill':
        if (!fill) fill = { char: t.char, index: text.length };
        break;
      case 'ampm':
        text += hour < 12 ? t.am : t.pm;
        break;
      case 'subsec':
        text += sep.decimal + String(sub).padStart(section.subsec, '0').slice(0, t.digits);
        break;
      case 'elapsed': {
        const value =
          t.part === 'h'
            ? Math.floor(totalSecs / 3600)
            : t.part === 'n'
              ? Math.floor(totalSecs / 60)
              : totalSecs;
        text += String(value).padStart(t.length, '0');
        break;
      }
      case 'date':
        text += datePart(t.part, t.length, date, section, { hour, minute, second });
        break;
      default:
        break;
    }
  }
  const out: Rendered = fill ? { text, fill } : { text };
  if (pads.length) out.pads = pads;
  return out;
}

function datePart(
  part: string,
  length: number,
  date: CalendarDate,
  section: Section,
  time: { hour: number; minute: number; second: number }
): string {
  switch (part) {
    case 'y':
      return length <= 2 ? two(date.year % 100) : String(date.year).padStart(4, '0');
    case 'm':
      if (length === 1) return String(date.month);
      if (length === 2) return two(date.month);
      if (length === 3) return MONTHS[date.month - 1]!.slice(0, 3);
      if (length === 4) return MONTHS[date.month - 1]!;
      return MONTHS[date.month - 1]![0]!;
    case 'd':
      if (length === 1) return String(date.day);
      if (length === 2) return two(date.day);
      if (length === 3) return DAYS[date.weekday]!.slice(0, 3);
      return DAYS[date.weekday]!;
    case 'h': {
      const h = section.ampm ? (time.hour % 12 === 0 ? 12 : time.hour % 12) : time.hour;
      return length >= 2 ? two(h) : String(h);
    }
    case 'n':
      return length >= 2 ? two(time.minute) : String(time.minute);
    case 's':
      return length >= 2 ? two(time.second) : String(time.second);
    default:
      return '';
  }
}
