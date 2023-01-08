// This code is adapted from https://github.com/garethdmm/SpreadsheetMagic

const properties = PropertiesService.getScriptProperties();
var API_KEY = properties.getProperty('OPENAI_API_KEY');

var preface = "I'm an intelligent bot. I will answer your questions. If I am not sure, I will reply with 'Unknown'.\
\
"

var NUM_TOKENS = 2048;

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();

  var menuItems = [
    {name: 'Fill with GPT-3', functionName: 'gpt3fill'},
  ];

  spreadsheet.addMenu('GPT3', menuItems);
}

function _callAPI(prompt) {
  var data = {
    'prompt': prompt,
    'max_tokens': NUM_TOKENS,
    'temperature': 0,
  };

  var options = {
    'method' : 'post',
    'contentType': 'application/json',
    'payload' : JSON.stringify(data),
    'headers': {
      Authorization: 'Bearer ' + API_KEY,
    },
  };

  response = UrlFetchApp.fetch(
    'https://api.openai.com/v1/engines/text-davinci-003/completions',
    options,
  );

  return JSON.parse(response.getContentText())['choices'][0]['text']
}

function gpt3(question){
  var prompt = preface + "Question: " + question + "\Answer:"
  return _callAPI(prompt);
}

function gpt3fill() {
  var spreadsheet = SpreadsheetApp.getActive();
  var range = spreadsheet.getActiveRange();
  var num_rows = range.getNumRows();
  var num_cols = range.getNumColumns();

  for (var i=1; i<num_rows + 1; i++) {
    question = range.getCell(i,1).getValue();
    fill_cell = range.getCell(i, num_cols);

    result = gpt3(question);

    fill_cell.setValue([result]);
  }
}
