const vbspretty = require('vbspretty')
var bsource = vbspretty({
    level: 0,
    indentChar: '    ',
    breakLineChar: '\n',
    breakOnSeperator: false,
    removeComments: false,
    source: require('fs').readFileSync('./index-unpretty.vbs').toString()
  });

  require('fs').writeFileSync('./index-pretty-js.vbs', bsource)