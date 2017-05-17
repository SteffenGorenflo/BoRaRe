if (typeof require !== 'undefined') XLSX = require('xlsx');
const maps = require('./gmaps.js');

module.exports = { transform };

const isTruthy = value => !!value;

function transform(files = []) {
    return files.map(file => XLSX.readFile(file))
        .map(workbook => workbook.Sheets)
        .map(transformSheets)
        .filter(isTruthy)
}

function transformSheets(sheets) {
    return Object.keys(sheets)
        .map(sheetName => ({ sheetName, sheet: sheets[sheetName] }))
        .map(transformSheet)
}

function transformSheet({ sheetName, sheet }) {
    // e.g.: !ref = A1:Q23
    const [meta] = sheet['!ref'].split(':');
    const maxRow = parseInt(meta.match(/\d+/g), 10);

    // find x
    const tours = findToursInSheet(sheet, maxRow);
    if (!tours) {
        return null;
    }
    return tours.map(tour => transformTour(sheet, tour));
}

function transformTour(sheet, tour) {

    console.log(tour);

    const firstCol = 'A';
    const placeCol = 'B';
    const hotelCol = 'C';
    const distanceCol = 'D';
    const durationCol = 'E';

    return Promise.resolve("Good job")

}

function findToursInSheet(sheet, maxRow) {

    const col = 'A';
    const tours = [];
    const tour = {};

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
