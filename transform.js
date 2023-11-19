const Docxtemplater = require('docxtemplater');
const XLSX = require('xlsx');
const PizZip = require("pizzip");
const fs = require("fs");
const path = require("path");

const wordType = ['n.', 'adj.', 'v.', 'vt.', 'vi.', '②', 'adv.']


const folderPath = './'; // 当前文件夹路径

const workbook = XLSX.utils.book_new();

fs.readdir(folderPath, (err, files) => {
    if (err) {
        console.error('Error reading folder:', err);
        return;
    }

    const docxFiles = files.filter(file => path.extname(file) === '.docx');
    const promises = [];
    docxFiles.forEach(file => {
        const filePath = path.join(folderPath, file);
        promises.push(processDocxFile(filePath));
    });

    Promise.all(promises)
        .then(() => {
            console.log('所有文档遍历完成');
            XLSX.writeFile(workbook, 'output.xlsx');
        })
        .catch(error => {
            console.error('处理文档时出错:', error);
        });
});



function isLetter(char) {
    let pattern = /^[a-zA-Z]+$/;
    return pattern.test(char);
}

function processDocxFile(filePath) {
    return new Promise((resolve, reject) => {
        // 读取文件内容
        // Load the docx file as binary content
        const content = fs.readFileSync(filePath, "binary");

        const zip = new PizZip(content);

        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });


        let s = doc.getFullText();

        let unitArray = s.split(/(Unit\s\d+)/);
        const units = unitArray.map((text, index) => {
            let unit = {}

            if (text.length > 100 && index - 1 >= 0) {
                unit.title = unitArray[index - 1];
                unit.content = text;
                return unit
            }

        }).filter((unit) => unit)


        units.forEach((unit, unitNumber) => {
            const regex = /([a-zA-Z\s]+\[.{1,35}\])/;

            // 获取所有段落
            const paragraphs = unit.content.split(regex).filter((paragraph) => paragraph.trim() !== '');

            let rows = [];
            paragraphs.forEach((value, index) => {
                if (value && typeof (value) === "string" && regex.test(value)) {
                    rows.push(value + paragraphs[index + 1]);
                }
            })

            let table = [];
            rows.forEach((value, index) => {
                let cells = []
                let R = "";
                let L = "";
                let E = "";
                if (value && typeof (value) === "string") {
                    rows[index] = value.split(/(\[.{1,35}\])/).filter((cell) => cell.trim() !== '');
                    cells.push(rows[index][0]);
                    cells.push(rows[index][1]);

                    let indexR = rows[index][2].indexOf('记');
                    let indexL = rows[index][2].indexOf('联');

                    if (indexR > 0) {
                        R = rows[index][2].slice(indexR);
                    }
                    if (indexL > 0) {
                        L = rows[index][2].slice(indexL);
                    }
                    if (indexR > 0 && indexL > 0) {
                        R = rows[index][2].slice(indexR, indexL);
                        if (indexR < indexL) {
                            E = rows[index][2].slice(0, indexR)
                        } else {
                            E = rows[index][2].slice(0, indexL)
                        }
                    } else if (indexL > 0) {

                        E = rows[index][2].slice(0, indexL)
                    } else if (indexR > 0) {

                        E = rows[index][2].slice(0, indexR)
                    } else {

                        E = rows[index][2]
                    }
                    //cells.push(E);

                    let indexArray = [];
                    wordType.forEach((value) => {

                        if (E.indexOf(value) === 0) {
                            indexArray.push(E.indexOf(value));
                            return
                        }

                        if (E[E.indexOf(value) - 1] === "／") {
                            return;

                        }

                        if (E.indexOf(value) > 0 && !isLetter(E[E.indexOf(value) - 1])) {

                            indexArray.push(E.indexOf(value));

                        }

                    })
                    indexArray.sort((a, b) => {
                        return a - b
                    });

                    let counter = 0;
                    indexArray.forEach((value, index) => {

                        if (value >= 0) {
                            if (counter < 1) {
                                if (indexArray[index + 1]) {
                                    cells.push.apply(cells, E.slice(value, indexArray[index + 1]).split(/例/))
                                } else {
                                    cells.push.apply(cells, E.slice(value).split('例'))
                                }
                                table.push(cells)
                                counter++;
                            } else {
                                if (indexArray[index + 1]) {
                                    table.push(['', '', E.slice(value, indexArray[index + 1]).split('例')[0], E.slice(value, indexArray[index + 1]).split('例')[1]])
                                } else {
                                    table.push(['', '', E.slice(value).split('例')[0], E.slice(value).split('例')[1]])
                                }
                            }
                        }
                    })

                    if (R) {
                        table.push([" ", " ", " ", R])
                    }
                    if (L) {
                        table.push([" ", " ", " ", L])
                    }
                }
            });

            const worksheet = XLSX.utils.aoa_to_sheet(table);
            XLSX.utils.book_append_sheet(workbook, worksheet, unit.title);
        })

        resolve();
    });
}






