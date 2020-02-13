const fs = require('fs')
const path = require('path')
const xlsx = require('node-xlsx');
const AdmZip = require('adm-zip');

const regZip = /[^\B]+.zip$/g

// 例：D:\\FolderA\\FolderB\\FolderC
//    |---------1---------| |---2---|

// 1. Path-of-Parent-Folder
let targetPrePath = ''
// 2. Path-of-Target-Folder
let targetCurrentPath = ''

main()

// --------------------------

function main () {
    checkParams()

    let dirTree = {}
    
    try {
        const targetPath = path.resolve(targetPrePath, targetCurrentPath)
        const stats = fs.statSync(targetPath)
        if (!stats.isDirectory()) throw new Error(`invalid target path: ${targetPath}`)

        readDirRecursively(targetPrePath, targetCurrentPath, dirTree)
    }
    catch(err) {
        console.log(err)
        return
    }
    console.log(dirTree)
    console.log('------------------')

    let rows = []
    rows.push([path.resolve(targetPrePath, targetCurrentPath)])
    
    createData(dirTree, rows)

    const exportData = [{
        name: "sheet1",
        data: rows
    }]
    
    writeExcel(exportData)
}

// --------------------------

function checkParams() {
    let args = process.argv.splice(2)
    if (args && args.length < 2) throw new Error(`Parameter error: ${args}`)

    targetPrePath =  args[0]
    targetCurrentPath = args[1]

    if (!targetCurrentPath) throw new Error(`Parameter error: ${args}`)

    const fullPath = path.resolve(targetPrePath, targetCurrentPath)
    try {
        const stats = fs.statSync(fullPath)
    }
    catch(err) {
        throw new Error(err)
    }
}

function readDirRecursively (prePath, currentPath, dirTree) {
    const fullPath = path.resolve(prePath, currentPath)

    const stats = fs.statSync(fullPath)

    // folder
    if (stats.isDirectory()) {
        dirTree[currentPath] = {}

        let files = fs.readdirSync(fullPath)
        files.forEach((item) => {
            readDirRecursively(fullPath, item, dirTree[currentPath])
        })
    }
    // .zip file
    else if(stats.isFile() && regZip.test(fullPath)) {
        // TODO
        const zip = new AdmZip(fullPath);
        const zipEntries = zip.getEntries();
        let zipDirTree = dirTree[currentPath] = {}
        
        zipEntries.forEach(function(zipEntry, i, arr) {
            let fullPath = zipEntry.entryName.toString()
            let arrPath = fullPath.split('/')
            let rootObj = zipDirTree

            if (zipEntry.isDirectory == false) {
                let fileName = arrPath.pop()
                while (arrPath && arrPath.length) {
                    const folder = arrPath.shift()
                    if(!folder) break
                    if (!rootObj[folder]) {
                        rootObj[folder] = {}
                    }
                    rootObj = rootObj[folder]
                }
                rootObj[fileName] = fileName
            }
            else {
                while(arrPath) {
                    const folder = arrPath.shift()
                    if(!folder) break
                    if (!rootObj[folder]) {
                        rootObj[folder] = {}
                    }
                    rootObj = rootObj[folder]
                }
            }
        });
    }
    // other file
    else if (stats.isFile()) {
        dirTree[currentPath] = currentPath
        return
    }
}

function createData(obj, rows, indents = '') {
    for(item in obj) {
        let cols = []
        indents.split(',').forEach(indent=>{
            cols.push(indent)
        })

        if(toString.call(obj[item]) === '[object Object]'){            
            cols.push(item)
            rows.push(cols)
            createData(obj[item], rows, indents + ',')
        } else {
            cols.push(obj[item])
            rows.push(cols)
        }
        console.log(cols)
    }
}

function writeExcel(data) {
    const buffer = xlsx.build(data);
    const fileName = `${targetCurrentPath}.xlsx`
    const fileFullPath = path.resolve(targetPrePath, targetCurrentPath, fileName)

    // write excel file
    fs.writeFile(fileName, buffer, function(err) {
        if (err) {
            console.log("Write failed: " + err + '\r\n-->Please close the file and try it again.');
            throw new Error(err)
        }

        console.log(`Write completed.\r\nCheck for: ${fileFullPath}`);
    });
}