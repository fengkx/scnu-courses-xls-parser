const path = require('path');
const XLSX = require('xlsx');
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
    const xlsFile = path.join(__dirname, 'excel-data', `${clsName}.xls`);
    const workBox = XLSX.readFile(xlsFile);

    const result = {}
    for(var R = range.s.r; R <= range.e.r; ++R) {
        for(var C = range.s.c; C <= range.e.c; ++C) {
            var cell_address = {c:C, r:R};
            var cell_ref = XLSX.utils.encode_cell(cell_address);
            const cell = workBox.Sheets.Sheet0[cell_ref];
            if(cell) {
                // no empty cell
                cell.v.split('\r\n').forEach(cellText => {
                    const [courseName, courseTime,coursePlace,teacher,size] = cellText.split('/');
                    const key = `${courseName}/${coursePlace}/${courseTime}/${teacher}`
                    const match = courseTime.match(/((\d+)-)?(\d+)周/);
                    // console.log(courseTime, match)
                    const [_, __,start, end] = match;
                    if(!result[key]) {
                        const [campus, place] = coursePlace.split(' ',2)
                        result[key] = {
                            size: parseInt(size),
                            teacher,
                            campus,
                            week: [Number(start) || 1, Number(end)],
                            department: cls2Department[clsName],
                            place,
                            courseName
                        }
                        result[key].part = place2Part[place]
                    }
                    if(result[key].order) {
                        // [...result[key], {order:R-1, day: C-1}] : [{order:R-1, day: C-1}]
                        result[key].order.push({order:R-1, day: C-1})
                    } else {
                        result[key].order = [{order:R-1, day: C-1}]
                    }
                })
            }
        }
    }
    return result;
}

module.exports = parseSchedule;
