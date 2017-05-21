if (typeof require !== 'undefined') XLSX = require('xlsx');
const maps = require('./gmaps.js');
const _ = require('lodash');

module.exports = {transform};

const isTruthy = value => !!value;
const get = cell => cell ? cell.v : undefined;

const setDistance = val => {
    if (val) {
        let cell = {};
        val = Math.round((val / 1000) * 10) / 10;
        cell.v = val;
        cell.w = val + '';
        cell.t = 'n';
        return cell;
    }
    return undefined;
};

const setDuration = val => {
    if (val) {
        let cell = {};
        cell.v = val / (60 * 60 * 24);
        cell.w = Math.floor(val / (60 * 60)) + ':' + Math.round(val / 60);
        cell.t = 'n';
        return cell;
    }
    return undefined;

};


function transform(files = []) {

    return Promise.all(files.map(file => {
        let workbook = XLSX.readFile(file);
        let splitted = file.split('.');

        return transformSheets(workbook.Sheets)
            .then(() => {
                XLSX.writeFile(workbook, splitted[0] + '_COPY.' + splitted[1], {bookType: 'xlsx'})
            })

    }));
}

function transformSheets(sheets) {
    return Promise.all(Object.keys(sheets)
        .map(sheetName => sheets[sheetName])
        .map(transformSheet)
        .filter(isTruthy))
}

function transformSheet(sheet) {
    // e.g.: !ref = A1:Q23
    const meta = sheet['!ref'].split(':')[1];
    const maxRow = parseInt(meta.match(/\d+/g), 10);

    const placeCol = 'B';
    const hotelCol = 'C';
    const distanceCol = 'D';
    const durationCol = 'E';

    return Promise.all(_.range(0, maxRow).map(async i => {

        let orig = get(sheet[hotelCol + i]) + ', ' + get(sheet[placeCol + i]);
        let dest = get(sheet[hotelCol + (i + 1)]) + ', ' + get(sheet[placeCol + (i + 1)]);
        let res;

        try {
            res = await maps.durAndDis({origin: orig, destination: dest});
        } catch (e) {
            res = {distance: {value: 0}, duration: {value: 0}}
        }

        let dis = setDistance(res.distance.value);
        let dur = setDuration(res.duration.value);

        if (dis) {
            sheet[distanceCol + (i + 1)] = dis;
        }
        if (dur) {
            sheet[durationCol + (i + 1)] = dur;
        }


    }));
}