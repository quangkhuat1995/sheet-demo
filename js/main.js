/* oss.sheetjs.com (C) 2014-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */

/** drop target **/
var _target = document.getElementById('drop');
var _file = document.getElementById('file');
var _grid = document.getElementById('grid');
var _output = document.getElementById('output');
var _doParseDataBtn = document.getElementById('doParseDataBtn');

/** Spinner **/
var spinner;

var _workstart = function () {
  // spinner = new Spinner().spin(_target);
};
var _workend = function () {
  // spinner.stop();
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

/* make the buttons for the sheets */
var make_buttons = function (sheetnames, cb) {
  var buttons = document.getElementById('buttons');
  buttons.innerHTML = '';
  sheetnames.forEach(function (s, idx) {
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.name = 'btn' + idx;
    btn.text = s;
    var txt = document.createElement('h3');
    txt.innerText = s;
    btn.appendChild(txt);
    btn.addEventListener(
      'click',
      function () {
        cb(idx);
      },
      false
    );
    buttons.appendChild(btn);
    buttons.appendChild(document.createElement('br'));
  });
};

function renderOutput(json) {
  _output.innerHTML = `<code>${JSON.stringify(json, null, 4)}</code>`;
}

function _resize() {
  // _grid.style.height = window.innerHeight - 200 + 'px';
  // _grid.style.width = window.innerWidth - 200 + 'px';
}
window.addEventListener('resize', _resize);

var _onsheet = function (json, sheetnames, select_sheet_cb) {
  // document.getElementById('footnote').style.display = 'none';

  make_buttons(sheetnames, select_sheet_cb);

  /* show grid */
  // _grid.style.display = 'block';
  _resize();

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
  drop: _target,
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
