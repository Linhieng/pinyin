import json from "./新华词典.json" assert {type: 'json'}
import ExcelJS from "exceljs"
import fs from 'fs'

const VOWEL = ['a', 'o', 'e', 'i', 'u', 'ü', 'ai', 'ei', 'ui', 'ao', 'ou', 'iu', 'ie', 'üe', 'er', 'an', 'en', 'in', 'un', 'ün', 'ang', 'eng', 'ing', 'ong']
const INITIAL = ['', 'b', 'p', 'm', 'f', 'd', 't', 'n', 'l', 'g', 'k', 'h', 'j', 'q', 'x', 'zh', 'sh', 'sh', 'r', 'z', 'c', 's', 'y', 'w']


const isOrder = true
const isROW_concat_COL = true
const isAddZh = true
const matrix = generateMatrix(isROW_concat_COL, isOrder)
const pinyin_Zh = getPinyinZhMap()

fillMatrix(matrix, pinyin_Zh, isROW_concat_COL, isAddZh)
matrix2excel(matrix, 'ROW_COL', 'csv')



/**
 * 与 flat() 相反，该函数用于将一维数组转换为二维数组
 * @param {*} array
 * @param {*} size
 * @returns
 */
function chunkArray(array, size = 1) {
    const result = []
    for (let i = 0; i < array.length; i += size) {
        result.push(array.slice(i, i + size))
    }
    return result
}

function generateMatrix(isROW_concat_COL = true, isOrder = false) {
    const initial = INITIAL.sort((a, b) => isOrder ? a.charCodeAt() - b.charCodeAt() : 0)
    const vowel = VOWEL.sort((a, b) => isOrder ? a.charCodeAt() - b.charCodeAt() : 0)

    let matrix = []
    let topRow = null
    let topCol = null

    if (isROW_concat_COL) {
        topRow = ['', ...initial]
        topCol = chunkArray(vowel, 1)
    } else {
        topRow = ['', ...vowel]
        topCol = chunkArray(initial, 1)
    }

    matrix.push(topRow, ...topCol)

    return matrix
}

function getPinyinZhMap() {
    const jsonFlat = json.flat().flat()
    const pinyinMap = {}
    jsonFlat.forEach(v => {
        const t = v.split(' ')
        pinyinMap[t[0]] = t[1]
    })
    return pinyinMap
}

function fillMatrix(matrix, pinyin_Zh, isROW_concat_COL, isAddZh) {
    for (let r = 1; r < matrix.length; r++) {
        for (let c = 1; c < matrix[0].length; c++) {
            const pinyin = concatPinyin(isROW_concat_COL, matrix[0][c], matrix[r][0])
            if (pinyin in pinyin_Zh) {
                matrix[r][c] = `${isAddZh ? pinyin_Zh[pinyin] + ' ' : ''}${pinyin}`
            }
        }
    }
    return matrix
}
function concatPinyin(isROW_COL, row, col) {
    return isROW_COL
        ? row + col
        : col + row
}

function matrix2excel(matrix, filenameNoExt, extension = 'xlsx') {
    // 创建一个新的工作簿
    const workbook = new ExcelJS.Workbook()

    // 添加一个工作表
    const worksheet = workbook.addWorksheet('Sheet1')

    // 将数组数据写入工作表
    matrix.forEach(row => {
        worksheet.addRow(row)
    })

    const filename = `${filenameNoExt}.${extension}`
    // 保存工作簿为表格文件
    workbook[extension].writeFile(filename)
        .then(function () {
            console.log('Excel 文件已保存')
            extension === 'csv' && toUtf8bom(filename)
        })
        .catch(function (error) {
            console.log('保存 Excel 文件时出错：', error)
        })
}

function toUtf8bom(filename) {
    try {
        // excel 打开 csv 文件时，utf8bom 才不会导致中文乱码
        const csvData = fs.readFileSync(filename, {encoding: 'utf8'})
        fs.writeFileSync('bom-'+filename, '\ufeff' + csvData, { encoding: 'utf8' })
    } catch (error) {
        console.log('无法转换成 utf8bom 文件', error);
    }
}
