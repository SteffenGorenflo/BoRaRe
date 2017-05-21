// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const {dialog} = require('electron').remote;

const button = document.getElementById('openFile');

const excel = require('./excel.js');


button.addEventListener('click', function () {
    dialog.showOpenDialog(
        {
            filters: [
                {name: 'Excel', extensions: ['xlsx', 'xlx']}
            ],
            properties: ['openFile', 'multiSelections']
        }, files => {
            if (files) {
                excel.transform(files)
                    .then(() => alert("Jetzt wirklich fertig. Ganz wirklich"));
            }

        })
});