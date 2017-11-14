var fs = require('fs');
const moment = require('moment');
const _ = require('lodash');
const cheerio = require('cheerio');
var request = require('request');
var XLSX = require('xlsx');
var workbook = XLSX.readFile('trackingDPD.ods');
// console.log(workbook);
var data = [];
var sheet_name_list = workbook.SheetNames;
// function getDatas() {
    sheet_name_list.forEach(function(y) {
        var worksheet = workbook.Sheets[y];
        var headers = {};
        for(z in worksheet) {
            if(z[0] === '!') continue;
            //parse out the column, row, and value
            var tt = 0;
            for (var i = 0; i < z.length; i++) {
                if (!isNaN(z[i])) {
                    tt = i;
                    break;
                }
            };
            var col = z.substring(0,tt);
            // console.log('col   '+col);
            var row = parseInt(z.substring(tt));
            // console.log('row  '+row);
            var value = worksheet[z].v;
            // console.log('value  '+value);

            //store header names
            if(row == 1 && value) {
                headers[col] = value;
                continue;
            }

            if(!data[row]) data[row]={};
            data[row][headers[col]] = value;
        }
        // console.log(data);
        //drop those first two rows which are empty
        data.shift();
        data.shift();
        // return data;
        // console.log(data);
    });
// }
let items = [];
// let urls=[
//     { 'ID ': 1, NAME: 'alan', tra: '05228940834041' },
//     { 'ID ': 2, NAME: 'aaaa', tra: '05228940834045' }];

// function scrape(url,file){
//
//     request(url, function (error, response, body) {
//         if (!error && response.statusCode == 200) {
//             console.log('request url: '+url);
//             console.log('request file: '+file);
//             // console.log(body);
//             console.log(body['soap:table']);
//             body('.contentInner:first-child .alternatingTable tr:not(:has(th))').each(function (idx, element) {
//                 console.log('fafafa');
//             })
//         }
//             // fs.writeFile(file, body);
//     });
//
// }
function getFromHtml(options, trackingCode) {
    let defaultOption = {
        uri: '',
        json: false
    };
    let cheerioOption = options.cheerio || {normalizeWhitespace: true};

    if(typeof options === 'string') {
        defaultOption.uri = options;
    } else if(typeof options === 'object') {
        defaultOption = _.extend(defaultOption, options.request);
    }
    return new Promise(function (resolve, reject) {
        request(defaultOption, function (error, response, body) {
            if (error) {
                reject(error);
            } else {
                if (response.statusCode !== 200) {
                    typeof defaultOption.strategy !== 'undefined' ? resolve(defaultOption.strategy) : reject(new Error('error status code:' + response.statusCode));
                } else {
                    let $ = cheerio.load(body, cheerioOption);
                    $('.contentInner:first-child .alternatingTable tr:not(:has(th))').each(function (idx, element) {
                        let $element = $(element);

                        let childTd = $($element.children('td'));

                        if(childTd != null) {
                            let statusCodes, codes;
                            statusCodes = $(childTd['5']).html();
                            codes = [];
                            if(statusCodes !== null) {
                                statusCodes = _.trim(statusCodes);
                                if(statusCodes.indexOf('<br>') !== -1) {
                                    codes = statusCodes.split(/<br>|, /);
                                } else {
                                    codes.push(statusCodes);
                                }
                            }
                            let status = _.trim($(childTd['2']).text());
                            if(status !== 0) {
                                var dateFormat = [
                                    'MM-DD-YYYY h:mm'
                                ];
                                var dateTime = moment(_.trim($(childTd['0']).text()), dateFormat, 'en').format();

                                items.unshift({
                                    trackingCode: trackingCode,
                                    dateTime : dateTime,
                                    status : _.trim($(childTd['2']).text())
                                    // status : status
                                });
                            }
                        }
                    });
                    var _headers = ['trackingCode', 'dateTime', 'status'];
                    var _data = items;
                    var headers = _headers
                    // 为 _headers 添加对应的单元格位置
                    // [ { v: 'id', position: 'A1' },
                    //   { v: 'name', position: 'B1' },
                    //   { v: 'age', position: 'C1' },
                    //   { v: 'country', position: 'D1' },
                    //   { v: 'remark', position: 'E1' } ]
                        .map((v, i) => Object.assign({}, {v: v, position: String.fromCharCode(65+i) + 1 }))
                        // 转换成 worksheet 需要的结构
                        // { A1: { v: 'id' },
                        //   B1: { v: 'name' },
                        //   C1: { v: 'age' },
                        //   D1: { v: 'country' },
                        //   E1: { v: 'remark' } }
                        .reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
                    var data = _data
                    // 匹配 headers 的位置，生成对应的单元格数据
                    // [ [ { v: '1', position: 'A2' },
                    //     { v: 'test1', position: 'B2' },
                    //     { v: '30', position: 'C2' },
                    //     { v: 'China', position: 'D2' },
                    //     { v: 'hello', position: 'E2' } ],
                    //   [ { v: '2', position: 'A3' },
                    //     { v: 'test2', position: 'B3' },
                    //     { v: '20', position: 'C3' },
                    //     { v: 'America', position: 'D3' },
                    //     { v: 'world', position: 'E3' } ],
                    //   [ { v: '3', position: 'A4' },
                    //     { v: 'test3', position: 'B4' },
                    //     { v: '18', position: 'C4' },
                    //     { v: 'Unkonw', position: 'D4' },
                    //     { v: '???', position: 'E4' } ] ]
                        .map((v, i) => _headers.map((k, j) => Object.assign({}, { v: v[k], position: String.fromCharCode(65+j) + (i+2) })))
                        // 对刚才的结果进行降维处理（二维数组变成一维数组）
                        // [ { v: '1', position: 'A2' },
                        //   { v: 'test1', position: 'B2' },
                        //   { v: '30', position: 'C2' },
                        //   { v: 'China', position: 'D2' },
                        //   { v: 'hello', position: 'E2' },
                        //   { v: '2', position: 'A3' },
                        //   { v: 'test2', position: 'B3' },
                        //   { v: '20', position: 'C3' },
                        //   { v: 'America', position: 'D3' },
                        //   { v: 'world', position: 'E3' },
                        //   { v: '3', position: 'A4' },
                        //   { v: 'test3', position: 'B4' },
                        //   { v: '18', position: 'C4' },
                        //   { v: 'Unkonw', position: 'D4' },
                        //   { v: '???', position: 'E4' } ]
                        .reduce((prev, next) => prev.concat(next))
                        // 转换成 worksheet 需要的结构
                        //   { A2: { v: '1' },
                        //     B2: { v: 'test1' },
                        //     C2: { v: '30' },
                        //     D2: { v: 'China' },
                        //     E2: { v: 'hello' },
                        //     A3: { v: '2' },
                        //     B3: { v: 'test2' },
                        //     C3: { v: '20' },
                        //     D3: { v: 'America' },
                        //     E3: { v: 'world' },
                        //     A4: { v: '3' },
                        //     B4: { v: 'test3' },
                        //     C4: { v: '18' },
                        //     D4: { v: 'Unkonw' },
                        //     E4: { v: '???' } }
                        .reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
// 合并 headers 和 data
                    var output = Object.assign({}, headers, data);
// 获取所有单元格的位置
                    var outputPos = Object.keys(output);
// 计算出范围
                    var ref = outputPos[0] + ':' + outputPos[outputPos.length - 1];
// 构建 workbook 对象
                    var wb = {
                        SheetNames: ['mySheet'],
                        Sheets: {
                            'mySheet': Object.assign({}, output, { '!ref': ref })
                        }
                    };
// 导出 Excel
                    XLSX.writeFile(wb, 'outoutput.xlsx');

                    console.log('trackingCode:'+trackingCode);
                    console.log('items:'+items);
                    // trackingInfo.status = items;
                    // return trackingInfo;
                    // resolve($);
                }
            }
        });
    });
};
// new Promise(function (resolve, reject) {
//     for (var i = 0; i < urls.length; i++) {
//
//         var file = 'log'+[i]+'.txt';
//         var url = urls[i];
//
//         console.log(url);
//         console.log(file);
//
//         getFromHtml({
//             request: {
//                 url: 'http://extranet.dpd.de/cgi-bin/delistrack?typ=2&lang=en&pknr=05228940834041&var=internalNewSearch&x=0&y=0'
//             }
//         }, '05228940834041');
//     }
//     console.log('items:'+items);
// });
new Promise(function (resolve, reject) {
    // console.log('dataaaaa:'+data);
    // return;
    data.forEach(dat => {
        // console.log('youmeiyouo:'+dat.tracking_numbers);
        // return;
        let trackingCode = '0'+dat.tracking_numbers;
        // console.log('youmeiyouo:'+trackingCode);
        // return;
        getFromHtml({
            request: {
                url: 'http://extranet.dpd.de/cgi-bin/delistrack?typ=2&lang=en&pknr=' + trackingCode + '&var=internalNewSearch&x=0&y=0'
            }
        }, trackingCode);
    });
    // console.log('iiiiiiiiiii:'+items);
});
