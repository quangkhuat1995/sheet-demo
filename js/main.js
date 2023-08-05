/* oss.sheetjs.com (C) 2014-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */

/** drop target **/
var _file = document.getElementById('file');
var _output = document.getElementById('output');
var _doParseDataBtn = document.getElementById('doParseDataBtn');

/** Spinner **/
var spinner;

var _workstart = function () {
  // spinner = new Spinner();
  console.log('[START] performance....', window.performance);
};
var _workend = function () {
  // spinner.stop();
  console.log('[END] performance....', window.performance);
};

/** Alerts **/
var _badfile = function () {
  window.alert(
    'This file does not appear to be a valid Excel file.  If we made a mistake, please report this issue to <a href="https://git.sheetjs.com/sheetjs/sheetjs/issues/">our bug tracker</a> so we can take a look.'
  );
};

var _pending = function () {
  alertify.alert('Please wait until the current file is processed.');
};

var _large = function (len, cb) {
  window.alert(
    'This file is ' +
      len +
      ' bytes and may take a few moments.  Your browser may lock up during this process'
  );
};

var _failed = function (e) {
  console.log(e, e.stack);
  window.alert(
    'We unfortunately dropped the ball here.  Please test the file using the <a href="/js-xlsx/">raw parser</a>.  If there are issues with the file processor, please report this issue to <a href="https://git.sheetjs.com/sheetjs/sheetjs/issues/">our bug tracker</a> so we can make things right.'
  );
};

function renderOutput(json) {
  const DISPLAY_ALL_JSON = false;

  const jsonString = JSON.stringify(json, null, 4);
  // don't display all
  const chunked = jsonString.substring(0, 1000);

  const displayValue = DISPLAY_ALL_JSON ? jsonString : chunked;

  _output.innerHTML = `<code>${displayValue}</code>
  <br/>


  <h2>Value are chunked down</h2>
  `;
}

var _onsheet = function (json) {
  /* set up table headers */
  var L = 0;
  json.forEach(function (r) {
    if (L < r.length) L = r.length;
  });
  console.log(L);
  for (var i = json[0].length; i < L; ++i) {
    json[0][i] = '';
  }

  /* load data */
  console.log('json', json);
  renderOutput(json);
};

/** Drop it like it's hot **/
DropSheet({
  file: _file,
  parseBtn: _doParseDataBtn,
  on: {
    workstart: _workstart,
    workend: _workend,
    sheet: _onsheet,
    foo: 'bar',
  },
  errors: {
    badfile: _badfile,
    pending: _pending,
    failed: _failed,
    large: _large,
    foo: 'bar',
  },
});
