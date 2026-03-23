/** UtilsClickUp.gs
 * ClickUp 연동 공통 유틸.
 */

function stringValue_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function toIntOrNull_(v) {
  if (v === null || v === undefined || v === '') return null;
  var n = Number(v);
  return isNaN(n) ? null : parseInt(n, 10);
}

function colToNumber_(colA1) {
  var s = stringValue_(colA1).toUpperCase();
  if (!s) return null;
  var n = 0;
  for (var i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
}

function formatDateYmd_(value, tz) {
  if (!value) return '';
  var d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return stringValue_(value);
  return Utilities.formatDate(d, tz || 'Etc/UTC', 'yyyy-MM-dd');
}

function toClickUpDueDateMs_(value, tz) {
  if (!value) return null;
  var d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return null;
  var ymd = Utilities.formatDate(d, tz || 'Etc/UTC', 'yyyy-MM-dd');
  var parts = ymd.split('-');
  var normalized = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 9, 0, 0, 0);
  return normalized.getTime();
}

function compressSpace_(s) {
  return stringValue_(s).replace(/\s+/g, ' ').trim();
}

function joinNonEmpty_(arr, sep) {
  return (arr || []).map(function(v) { return compressSpace_(v); }).filter(function(v) { return !!v; }).join(sep || ' ');
}

function parseFirstParenText_(text) {
  var s = stringValue_(text);
  var m = s.match(/\(([^)]+)\)/);
  return m && m[1] ? compressSpace_(m[1]) : '';
}

function safeSlug_(text) {
  var s = romanizeKoreanSimple_(stringValue_(text).toUpperCase());
  s = s.replace(/[^A-Z0-9]+/g, '-').replace(/-+/g, '-').replace(/^-|-$/g, '');
  return s || 'NA';
}

function romanizeKoreanSimple_(text) {
  var s = stringValue_(text);
  if (!s) return '';

  var lead = ['G','KK','N','D','TT','R','M','B','PP','S','SS','','J','JJ','CH','K','T','P','H'];
  var vowel = ['A','AE','YA','YAE','EO','E','YEO','YE','O','WA','WAE','OE','YO','U','WEO','WE','WI','YU','EU','YI','I'];
  var tail = ['','K','K','KS','N','NJ','NH','T','L','LK','LM','LP','LS','LT','LP','LH','M','P','PS','T','T','NG','T','T','K','T','P','H'];

  var out = '';
  for (var i = 0; i < s.length; i++) {
    var code = s.charCodeAt(i);
    if (code >= 0xAC00 && code <= 0xD7A3) {
      var t = code - 0xAC00;
      var l = Math.floor(t / 588);
      var v = Math.floor((t % 588) / 28);
      var c = t % 28;
      out += (lead[l] || '') + (vowel[v] || '') + (tail[c] || '');
    } else {
      out += s.charAt(i);
    }
  }
  return out;
}

function shortHash6_(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, stringValue_(text));
  var hex = bytes.map(function(b) {
    var h = (b < 0 ? b + 256 : b).toString(16);
    return h.length === 1 ? '0' + h : h;
  }).join('').toUpperCase();
  return hex.substring(0, 6);
}
