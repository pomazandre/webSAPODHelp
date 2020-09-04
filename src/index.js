var Express = require('express')
var App = Express();
const Port = 3000;
var helpWriter = require('./HelpWriter/help-writer');

App.listen(Port, (Err) => {
  if (Err) {
    return console.log('error :', Err)
  }
  helpWriter.parse();
  //console.log(`webHelp is listening on ${Port}`)
})
