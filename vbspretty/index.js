const vbspretty = require('vbspretty')
var bsource = vbspretty({
    level: 1,
    indentChar: '\t',
    breakLineChar: '\r\n',
    breakOnSeperator: false,
    removeComments: false,
    source: require('fs').readFileSync('./vbspretty-unpretty.vbs').toString()
  });

  require('fs').writeFileSync('./vbspretty-pretty-js.vbs', bsource)