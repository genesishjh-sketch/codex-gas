/** 안전한 템플릿 파서: 허용된 필드/함수/연산자만 지원 */
function buildTemplateContextFromMilestoneRow_(rowObj) {
  return {
    project_code: rowObj.project_code || '',
    project_name: cleanupProjectName_(rowObj.project_code || ''),
    section: rowObj.section || '',
    step_name: rowObj.step_name || '',
    plan_date: rowObj.plan_date || '',
    done_date: rowObj.done_date || '',
    manager: rowObj.manager || ''
  };
}

function cleanupProjectName_(projectCode) {
  var text = (projectCode || '').toString().trim();
  return text.replace(/^\d{6}\s+/, '');
}

function renderSafeTemplate_(template, context) {
  if (!template) return '';
  var parser = new SafeTemplateParser_(template, context || {});
  return parser.parse();
}

function SafeTemplateParser_(text, context) {
  this.text = text || '';
  this.context = context || {};
  this.pos = 0;
}

SafeTemplateParser_.prototype.parse = function() {
  var result = this.parseConcat_();
  this.skipSpaces_();
  if (!this.eof_()) throw new Error('템플릿 구문 오류(남은 토큰): ' + this.text.slice(this.pos));
  return valueToString_(result);
};

SafeTemplateParser_.prototype.parseConcat_ = function() {
  var left = this.parsePrimary_();
  this.skipSpaces_();
  while (this.peek_() === '&') {
    this.pos++;
    var right = this.parsePrimary_();
    left = valueToString_(left) + valueToString_(right);
    this.skipSpaces_();
  }
  return left;
};

SafeTemplateParser_.prototype.parsePrimary_ = function() {
  this.skipSpaces_();
  var ch = this.peek_();
  if (!ch) throw new Error('템플릿 구문 오류: 예상치 못한 종료');

  if (ch === '"') return this.parseString_();
  if (ch === '(') {
    this.pos++;
    var value = this.parseConcat_();
    this.expect_(')');
    return value;
  }
  if (/[0-9]/.test(ch)) return this.parseNumber_();
  if (/[A-Za-z_]/.test(ch)) {
    var identifier = this.parseIdentifier_();
    this.skipSpaces_();
    if (this.peek_() === '(') {
      return this.parseFunctionCall_(identifier);
    }
    return this.resolveField_(identifier);
  }

  throw new Error('템플릿 구문 오류: 지원하지 않는 토큰 ' + ch);
};

SafeTemplateParser_.prototype.parseFunctionCall_ = function(name) {
  var fn = name.toUpperCase();
  this.expect_('(');
  var args = [];
  this.skipSpaces_();
  if (this.peek_() !== ')') {
    while (true) {
      args.push(this.parseConcat_());
      this.skipSpaces_();
      if (this.peek_() === ',') {
        this.pos++;
        continue;
      }
      break;
    }
  }
  this.expect_(')');
  return this.executeFunction_(fn, args);
};

SafeTemplateParser_.prototype.executeFunction_ = function(fn, args) {
  if (fn === 'TEXT') {
    var dateVal = args[0];
    var fmt = valueToString_(args[1] || 'yyyy-MM-dd');
    if (!dateVal) return '';
    var dateObj = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(dateObj.getTime())) return '';
    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), fmt);
  }

  var source = valueToString_(args[0] || '');
  if (fn === 'LEFT') return source.substring(0, parseInt(args[1], 10) || 0);
  if (fn === 'RIGHT') {
    var count = parseInt(args[1], 10) || 0;
    return count <= 0 ? '' : source.slice(-count);
  }
  if (fn === 'MID') {
    var start = (parseInt(args[1], 10) || 1) - 1;
    var length = parseInt(args[2], 10) || 0;
    if (start < 0) start = 0;
    return source.substr(start, length);
  }
  if (fn === 'LEN') return source.length;

  throw new Error('허용되지 않은 함수: ' + fn);
};

SafeTemplateParser_.prototype.resolveField_ = function(name) {
  if (!this.context.hasOwnProperty(name)) {
    throw new Error('허용되지 않은 필드: ' + name);
  }
  return this.context[name];
};

SafeTemplateParser_.prototype.parseString_ = function() {
  this.expect_('"');
  var result = '';
  while (!this.eof_()) {
    var ch = this.text.charAt(this.pos++);
    if (ch === '"') return result;
    result += ch;
  }
  throw new Error('문자열 종료 따옴표가 없습니다.');
};

SafeTemplateParser_.prototype.parseNumber_ = function() {
  var start = this.pos;
  while (!this.eof_() && /[0-9]/.test(this.peek_())) this.pos++;
  return parseInt(this.text.slice(start, this.pos), 10);
};

SafeTemplateParser_.prototype.parseIdentifier_ = function() {
  var start = this.pos;
  while (!this.eof_() && /[A-Za-z0-9_]/.test(this.peek_())) this.pos++;
  return this.text.slice(start, this.pos);
};

SafeTemplateParser_.prototype.skipSpaces_ = function() {
  while (!this.eof_() && /\s/.test(this.peek_())) this.pos++;
};

SafeTemplateParser_.prototype.expect_ = function(char) {
  this.skipSpaces_();
  if (this.peek_() !== char) throw new Error('템플릿 구문 오류: "' + char + '" 기대');
  this.pos++;
};

SafeTemplateParser_.prototype.peek_ = function() {
  return this.text.charAt(this.pos);
};

SafeTemplateParser_.prototype.eof_ = function() {
  return this.pos >= this.text.length;
};

function valueToString_(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(value);
}
