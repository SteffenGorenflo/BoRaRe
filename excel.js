if (typeof require !== 'undefined') XLSX = require('xlsx');
const maps = require('./gmaps.js');

module.exports = {transform};

function transform(files) {

    files.forEach(file => {

        let workbook = XLSX.readFile(file);

        Promise.all(Object.entries(workbook.Sheets)
            .map(transformSheet))
            .then(console.log)
            .catch(console.log);


    })
}

function transformSheet(metaSheet) {
    return new Promise((resolve, reject) => {

        let sheetName = metaSheet[0];
        let sheet = metaSheet[1];

        // e.g.: !ref = A1:Q23
        let meta = sheet['!ref'].split(':');
        let maxRow = parseInt(meta[1].match(/\d+/g), 10);

        // find x
        let tours = findToursInSheet(sheet, maxRow);
        if (!tours) {
            return reject("no tours were found");
        } else {
            return Promise.all(tours.map(tour => transformTour(sheet, tour)));
        }
    })
}

function transformTour(sheet, tour) {
    return new Promise((resolve, reject) => {

        console.log(tour);

        let firstCol = 'A';
        let placeCol = 'B';
        let hotelCol = 'C';
        let distanceCol = 'D';
        let durationCol = 'E';



        return resolve("Good job")
    })
}

function findToursInSheet(sheet, maxRow) {

    let col = 'A';
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
