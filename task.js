const fs = require('fs');
const readline = require('readline-sync');
const xlsx = require('xlsx');

const argv = require('optimist').argv;

let inputDir;
let outputDir;

function checkArgs (argv) {
  if (argv.help) {
    console.log('Use:\ntask --inputDir \'path to dir\' --outputDir \'new path to dir\'');
    process.exit(1);
  }
  if (argv.inputDir !== undefined) {
    inputDir = argv.inputDir;
  } else {
    inputDir = readline.question('Input file name | file absolute path:\n');
  }
  if (argv.outputDir !== undefined) {
    outputDir = argv.outputDir;
  } else {
    outputDir = inputDir;
  }
}

function checkFolder (path) {
  if (!fs.existsSync(path)) {
    throw new Error(path + ' doesn\'t exists');
  } else if (!fs.statSync(path).isDirectory()) {
    throw new Error(path + ' isn\'t a folder');
  }
}

function jsonToXlsx (file, outputDir) {
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
  }
  const files = fs.readdirSync(file);
  files.forEach((fileName, index, array) => {
    if (fs.statSync(file + '\\' + fileName).isFile() && fileName.endsWith('.json')) {
      createJsonFile(file, fileName, outputDir);
    } else {
      if (fs.statSync(file + '\\' + fileName).isDirectory()) {
        jsonToXlsx(file + '\\' + fileName, outputDir + '\\' + fileName);
      }
    }
  });
}

function createJsonFile (file, fileName, outputDir) {
  const fullPath = fs.realpathSync(file + '\\' + fileName);
  const newPath = fs.realpathSync(outputDir) + '\\' + fileName.replace('.json', '') + '.xlsx';
  const jsonFile = openJsonFile(fullPath);
  const ws = xlsx.utils.json_to_sheet([jsonFile]);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, fileName.replace('.json', ''));
  console.log('Json file - \n\t' + fullPath + '\nwas found.\n----------');
  xlsx.writeFile(wb, newPath);
  console.log('XLSX file - \n\t' + newPath + '\nwas succesfully created.\n------------------------');
}

function openJsonFile (filePath) {
  const jsonObject = require(filePath);
  const properties = Object.getOwnPropertyNames(jsonObject);
  properties.forEach(property => {
    if (typeof jsonObject[property] === 'object') {
      jsonObject[property] = JSON.stringify(jsonObject[property]);
    }
  });
  return jsonObject;
}

checkArgs(argv);
checkFolder(inputDir);
checkFolder(outputDir);
jsonToXlsx(inputDir, outputDir);
