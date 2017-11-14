var fs = require('fs');
const parseString = require('xml2js').parseString;
var request = require('request');

// var urls = new Array("http://www.yahoo.com","http://www.bing.com");
var urls = new Array("http://extranet.dpd.de/cgi-bin/delistrack?typ=2&lang=en&pknr=05228940834041&var=internalNewSearch&x=0&y=0");

for (var i = 0; i < urls.length; i++) {

    var file = 'log'+[i]+'.txt';
    var url = urls[i];

    // console.log(url);
    // console.log(file);

    new Promise(function (resolve, reject) {
        console.log('come in');
        request.get({
            url: url,
            // body : wsdlXml,
            // headers : {
            //     'Content-type': 'application/x-www-form-urlencoded',
            //     'Content-Length' : wsdlXml.length
            // }
        }, function(err, response, body) {
            console.log('jinlail');
            if (err) {
                console.log('err');
                reject(err);
            }
            try {
                if (err) {
                    console.log('2err');
                    throw err;
                }
                console.log('jinjin');
                console.log(response.statusCode);
                !err && response.statusCode === 200 && parseString(body, function(err1, result) {
                    console.log('result:'+result);
                    fs.writeFile(file, result);
                    resolve(result);
                    // let mid = result['soap:Envelope']['soap:Body'][0]['consultaExpedicionesStrResponse'][0]['out'][0]['_'];
                    // if (err1) {
                    //     throw err1;
                    // }

                    // !err1 && parseString(mid, function (err2, resData) {
                    //     if (!err2 && resData['EXPEDICIONES'] !== undefined) {
                    //         let situaciones = resData['EXPEDICIONES']['EXPEDICION'][0]['SITUACIONES'][0]['SITUACION'];
                    //         resolve(situaciones);
                    //     } else {
                    //         reject(err2);
                    //     }
                    // })
                })
            }catch (err) {
                reject(err);
            }
        });
    });

    // request('http://www.baidu.com', function (error, response, body) {
    //     if (!error && response.statusCode == 200) {
    //         console.log('request url: '+url);
    //         console.log('request f' +
    //             'ile: '+file);
    //         console.log(body.title)
    //         fs.writeFile(file, body);
    //     }
    // });

}
