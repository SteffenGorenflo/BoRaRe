const googleMapsClient = require('@google/maps').createClient({
    key: 'AIzaSyDrQjznY__r5X0g91pdOhtJISy7QUzuIPY',
    Promise: require('q').Promise
});

module.exports = {durAndDis};


function durAndDis(places) {
    return new Promise((resolve, reject) => {
        googleMapsClient.directions(places)
            .asPromise()
            .catch(reject)
            .then(response => {

                if (response.json.status === 'OK') {
                    return resolve({
                        distance: response.json.routes[0].legs[0].distance,
                        duration: response.json.routes[0].legs[0].duration
                    })
                } else {
                    return reject('Nothing found')
                }
            });

    })

}