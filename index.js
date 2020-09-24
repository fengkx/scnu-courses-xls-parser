const path = require('path');
const fs = require('fs').promises;
const parseSchedule = require('./parser');

;(async () => {
    const files = await fs.readdir(path.join(__dirname, 'excel-data'));
    const clsNames = files.map(item => item.substring(0,item.length-4));
    const data = clsNames.map(clsName => parseSchedule(clsName)).filter(item => Object.keys(item).length > 0);
    console.log(JSON.stringify(data, null, 2))
})()
