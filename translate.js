var fs = require('fs');
var XLSX = require('xlsx');

fs.readdir('test', function(err,files){
    if(err){
        console.log(err);
    }
    files.forEach(file => {
        let choseKey = '';
        let workbook = XLSX.readFile('./test/' + file);
        let data = [];
        let sheet = workbook.SheetNames;
        sheet.forEach(function(y) {
            var worksheet = workbook.Sheets[y];
            var headers = {};
            for(z in worksheet) {
                if(z[0] === '!') continue;
                // parse out the column, row, and value
                var tt = 0;
                for (var i = 0; i < z.length; i++) {
                    if (!isNaN(z[i])) {
                        tt = i;
                        break;
                    }
                }
                var col = z.substring(0, tt);
                var row = parseInt(z.substring(tt));
                var value = worksheet[z].v;
                if(row == 1 && value) {
                    headers[col] = value;
                    continue;
                }

                if(!data[row]) data[row]={};
                data[row][headers[col]] = value;
            }
            data.shift();
            data.shift();
        });
        let translations = {};
        data.forEach(tra => {
            if (typeof (tra.EN) != 'undefined') {
                let lastOne = Object.keys(tra);
                choseKey = lastOne[lastOne.length - 1];
                translations[tra.EN] = tra[choseKey];
            }
        });
        let lowKey = choseKey.toLowerCase();
        let test = JSON.stringify(translations);
        // let path = 'resources/locale-'+lowKey+'.json';
        let path = 'result/locale-'+lowKey+'.json';
        fs.appendFile(path, test);
    })
});