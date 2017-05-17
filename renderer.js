// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const {dialog} = require('electron').remote;

const button = document.getElementById('openFile');

const excel = require('./excel.js');

excel.transform(['/Users/steffen/Google Drive/Eigene Dateien/Kostenplanung 2017-05-01 - FERTIG.xlsx']);


button.addEventListener('click', function () {
    dialog.showOpenDialog(
        {
            filters: [
                {name: 'Excel', extensions: ['xlsx', 'xlx']},
                {name: 'Images', extensions: ['jpg', 'png', 'gif']}
            ],
            properties: ['openFile', 'multiSelections']
        }, files => {
            if (files) {
                files.forEach(console.log);
                excel.transform(files);
            }

        })
});