let fs = require('fs')
let readline = require('readline-sync')
let xlsx = require('xlsx')
let path
let argv = require('optimist').argv

if (Object.keys(argv).length > 2) {
  if (argv.inputDir === undefined || argv.outputDir === undefined) {
    console.log("Use \n--inputDir 'path'\n--outputDir 'path'")
  } else {
    renameFolder(argv.inputDir, argv.outputDir)
  }
} else {
  body()
}

function body () {
  do {
    path = readline.question('Input file name | file absolute path:\n')
  } while (!checkFolder(path))

  console.log('--------\nJson files:')
  jsonToXlsx(path)
}

function checkFolder (pathToFolder) {
  if (!fs.existsSync(pathToFolder)) {
    console.clear()
    console.log(pathToFolder + ' doesn\'t exists')
    return false
  } else if (!fs.statSync(pathToFolder).isDirectory()) {
    console.clear()
    console.log(pathToFolder + ' isn\'t a folder')
    return false
  } else {
    console.log('Ok')
    return true
  }
}

function jsonToXlsx (object) {
  let files = fs.readdirSync(object)
  for (let i = 0; i < files.length; i++) {
    if (fs.statSync(object + '\\' + files[i]).isFile() && files[i].endsWith('.json')) {
      let path = fs.realpathSync(object + '\\' + files[i])
      let jsonFile = require(path)
      let ws = xlsx.utils.json_to_sheet([jsonFile])
      let wb = xlsx.utils.book_new()
      xlsx.utils.book_append_sheet(wb, ws, files[i])
      console.log(path)
      try {
        xlsx.writeFile(wb, path.replace('.json', '.xlsx'))
        console.log('  - ' + path.replace('.json', '.xlsx') + ' was succesfully created.')
      } catch (error) {
        console.log(error.message)
      }
    } else {
      if (fs.statSync(object + '\\' + files[i]).isDirectory()) {
        jsonToXlsx(object + '\\' + files[i])
      }
    }
  }
}

function renameFolder (currFolder, newFolder) {
  console.log(fs.existsSync(currFolder) + '!' + fs.statSync(currFolder).isDirectory())
  if (fs.existsSync(currFolder) && fs.statSync(currFolder).isDirectory()) {
    if (fs.existsSync(newFolder) && fs.statSync(newFolder).isDirectory()) {
      var path = require('path')
      var folderPath = fs.realpathSync(currFolder)
      var folderName = path.basename(folderPath)
      fs.renameSync(currFolder, newFolder + '\\' + folderName)
      console.log(folderPath + ' succesfully moved to ' + fs.realpathSync(newFolder + '\\' + folderName))
    } else {
      throw new Error(newFolder + ' is not exist OR is not a directory')
    }
  } else {
    throw new Error(currFolder + ' is not exist OR is not a directory')
  }
}
