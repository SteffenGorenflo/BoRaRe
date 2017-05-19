if (typeof require !== 'undefined') XLSX = require('xlsx');
const maps = require('./gmaps.js');
const _ = require('lodash');

module.exports = {transform};

const isTruthy = value => !!value;
const get = cell => cell ? cell.v : undefined;

const setDistance = (cell, val) => {
    if (cell && val) {
        val = Math.round((val / 1000) * 10) / 10;
        cell.v = val;
        cell.w = val + '';
    }
};

const setDuration = (cell, val) => {
    if (cell && val) {
        cell.v = val / (60 * 60 * 24);
        cell.w = Math.floor(val / (60 * 60)) + ':' + Math.round(val / 60);
    }

};


function transform(files = []) {
    return new Promise((resolve, reject) => {

        files.forEach(file => {
            let workbook = XLSX.readFile(file);

            transformSheets(workbook.Sheets);

            let splitted = file.split('.');
            XLSX.writeFile(workbook, splitted[0] + '_COPY.' + splitted[1], {bookType: 'xlsx'});
        });

        return resolve('fertiiiiiiig');
    })
}

function transformSheets(sheets) {
    return Object.keys(sheets)
        .map(sheetName => ({sheetName, sheet: sheets[sheetName]}))
        .map(transformSheet)
        .filter(isTruthy)
}

function transformSheet({sheetName, sheet}) {
    // e.g.: !ref = A1:Q23
    const meta = sheet['!ref'].split(':')[1];
    const maxRow = parseInt(meta.match(/\d+/g), 10);

    // find x
    const tours = findToursInSheet(sheet, maxRow);
    if (!tours) {
        return null;
    }
    return tours.map(tour => transformTour(sheet, tour));
}

function transformTour(sheet, tour) {

    const placeCol = 'B';
    const hotelCol = 'C';
    const distanceCol = 'D';
    const durationCol = 'E';

    _.range(tour.start, tour.end - 1).map(async i => {

        let orig = get(sheet[hotelCol + i]) + ', ' + get(sheet[placeCol + i]);
        let dest = get(sheet[hotelCol + (i + 1)]) + ', ' + get(sheet[placeCol + (i + 1)]);
        let res;

        try {
            res = await maps.durAndDis({origin: orig, destination: dest});
        } catch (e) {
            res = {distance: {value: 0}, duration: {value: 0}}
        }

        setDistance(sheet[distanceCol + i], res.distance.value);
        setDuration(sheet[durationCol + i], res.duration.value);

    });

    return sheet;
}

function findToursInSheet(sheet, maxRow) {

    const col = 'A';
    let tours = [];
    let tour = {};

    // fuzzy logic ¯\_(ツ)_/¯
    for (let i = 0; i <= maxRow; i++) {

        let x = (sheet[col + i] ? sheet[col + i].v : undefined);

        if (!x) {
            if (tour.start && tour.pause) {
                tour.end = i - 1;
                tours.push(tour);
                tour = {};
            }
        }

        if (x && x === '#') {

            if (!tour.start) {
                tour.start = i + 1;
            } else {
                tour.pause = i;
            }
        }


        if (i === maxRow) {
            tour.end = i;
            tours.push(tour);
        }
    }
    return tours;
}
