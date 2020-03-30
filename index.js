const fs = require('fs')
const path = require('path')
const xlsx = require('node-xlsx');
const AdmZip = require('adm-zip');

const regZip = /^.+(?:(\.zip)+)$/ig

// start running.
main()

// --------------------------

/**
 * Entry method
 *
 */
function main () {
    let { targetPath, outputPath } = checkParams()

    console.log(`
        /_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
        /_/
        /_/ Processing...    
        /_/ Pleate wait... 
        /_/
        /_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    `)

    let dirTree = {}
    let targetPrePath, targetCurrentPath
    
    try {
        // joined with [\]
        if (targetPath.includes('\\')) {
            [ targetPrePath, targetCurrentPath ] = devidePath(targetPath, '\\')
            
        }
        // joined with [/]
        else if (targetPath.includes('/')) {
            [ targetPrePath, targetCurrentPath ] = devidePath(targetPath, '/')
        }
        // disk or invalid path
        else {
            [ targetPrePath, targetCurrentPath ] = devidePath(targetPath, '')
        }

        readDirRecursively(targetPrePath, targetCurrentPath, dirTree)
    }
    catch(err) {
        console.log(err)
        return
    }
    // console.log(dirTree)

    let rows = []
    rows.push([path.resolve(targetPrePath, targetCurrentPath)])
    
    createData(dirTree, rows)

    const exportData = [{
        name: "sheet1",
        data: rows
    }]
    
    writeExcel(exportData, targetCurrentPath, outputPath)
}

// --------------------------

/**
 * check for parameters from CLI
 *
 * @returns { targetPath, outputPath }
 */
function checkParams() {
    let args = process.argv.splice(2)

    // no parameters recieved
    if (!args || !args.length) throw new Error(`no-parameter error.\r\nPlease input full path of the target folder.\r\nYou can also specify an output path.`)

    let targetPath = ''
    let outputPath = ''

    try {        
        // 対象フォルダ
        targetPath = args[0]
        const statsTarget = fs.statSync(targetPath)
        if (!statsTarget.isDirectory()) throw new Error(`invalid target path: ${targetPath}`)

        // 出力先
        outputPath = (args.length > 1) ? args[1] : __dirname
        const statsOutPath = fs.statSync(outputPath)
        if (!statsOutPath.isDirectory()) {
            console.log(`
                invalid output path: ${outputPath}
                using ${__dirname} instead.
            `)
            outputPath = __dirname
        }
    }
    catch(err) {
        throw new Error(err)
    }

    console.log(`
        /_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
        /_/
        /_/ Target Folder: ${targetPath}    
        /_/ Output Path  : ${outputPath}    
        /_/
        /_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    `)

    return { targetPath, outputPath }
}

/**
 * devide path for current folder name
 *
 * @param {String} targetPath
 * @param {String | null} seperator
 * @returns
 */
function devidePath (targetPath, seperator) {
    let targetPrePath
    let targetCurrentPath

    // [\]または[/]で分割する場合
    if (seperator) {
        let arr = targetPath.split(seperator)
        
        targetCurrentPath = arr.pop()
        while (targetCurrentPath === '') {
            targetCurrentPath = arr.pop()
        }
        targetPrePath = arr.join(seperator)  
    }
    // disk の場合
    else {
        targetPrePath = ''
        targetCurrentPath = targetPath
    }

    return [targetPrePath, targetCurrentPath]
}

/**
 * read directory recursively
 *
 * @param {String} prePath
 * @param {String} currentPath
 * @param {Object} dirTree
 */
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
        const zip = new AdmZip(fullPath);
        const zipEntries = zip.getEntries();
        let zipDirTree = dirTree[currentPath] = {}
        
        zipEntries.forEach(function(zipEntry) {
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

/**
 * create data for excel
 *
 * @param {Object} obj
 * @param {Array} rows
 * @param {String} [indents='']
 */
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
        // console.log(cols)
    }
}

/**
 * create excel and write data
 *
 * @param {Array} data
 * @param {String} targetCurrentPath
 * @param {String} outputPath
 */
function writeExcel(data, targetCurrentPath, outputPath) {
    const buffer = xlsx.build(data);
    const fileName = `${targetCurrentPath}.xlsx`
    const fileFullPath = path.resolve(outputPath, fileName)

    // write excel file
    fs.writeFile(fileName, buffer, function(err) {
        if (err) {
            console.log("Write failed: " + err + '\r\n-->Please close the file and try it again.');
            throw new Error(err)
        }

        console.log(`Write completed.\r\nCheck for: ${fileFullPath}`);
    });
}