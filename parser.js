const path = require('path');
const XLSX = require('xlsx');
const debug = require('debug');
const cls2Department = require('./json-data/class2department.json');
const place2Part = require('./json-data/place2part.json');
const range = {
    s: {
        c: 2,
        r: 2
    },
    e: {
        c: 8,
        r: 12
    }
}

function parseSchedule(clsName) {
    const dLog = debug(clsName);
    const xlsFile = path.join(__dirname, 'excel-data', `${clsName}.xls`);
    const workBox = XLSX.readFile(xlsFile);

    const result = {}
    for(let R = range.s.r; R <= range.e.r; ++R) {
        for(let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = {c:C, r:R};
            const cellRef = XLSX.utils.encode_cell(cellAddress);
            const cell = workBox.Sheets.Sheet0[cellRef];
            if(!cell) {
                continue;
            }
            for (const cellText of cell.v.split('\r\n')) {
                const sep = cellText.split('/');
                let courseName, courseTime,coursePlace,teacher,size, _;
                // debug('123')(sep.length)
                if(sep.length===5) {
                    [courseName, courseTime,coursePlace,teacher,size] = sep;
                } else if(sep.length === 6) {
                    [_, courseName, courseTime,coursePlace,teacher,size] = sep;
                    courseName=`${_}/${courseName}`
                } else {
                    dLog('no enoguh info', cellText)
                    continue;
                }
                const key = `${courseName}/${coursePlace}/${courseTime}/${teacher}`
                if(!result[key]) {
                    const [campus, place] = coursePlace.split(' ',2)
                    result[key] = {
                        size: parseInt(size),
                        teacher,
                        campus,
                        week: [],
                        department: cls2Department[clsName],
                        place,
                        courseName
                    }
                    result[key].part = place2Part[place];
                    // (1-3节,5-6节)11-13周(单),14-18周
                    // 暂忽略单双
                    const weekRe = /((\d+)-)?(\d+)周/;
                    courseTime.split(',').forEach(text => {
                        const match = text.match(weekRe);
                        if(!match) {
                            if(place !== '未排地点') {
                                dLog(text, cellText);
                            }
                        } else {
                            const [_, __,start, end] = match;
                            result[key].week.push([start, end]);
                        }
                    })
                }
                if(result[key].section) {
                    // [...result[key], {order:R-1, day: C-1}] : [{order:R-1, day: C-1}]
                    result[key].section.push({section:R-1, day: C-1})
                } else {
                    result[key].section = [{section:R-1, day: C-1}]
                }
            }
        }
    }
    return result;
}

module.exports = parseSchedule;
