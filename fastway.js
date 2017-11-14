var fs = require('fs');
const moment = require('moment');
const _ = require('lodash');
const cheerio = require('cheerio');
var request = require('request');
var XLSX = require('xlsx');
var workbook = XLSX.readFile('trackingFastway.ods');
// var workbook = XLSX.readFile('test.ods');
var data = [];
var sheet_name_list = workbook.SheetNames;
// function getDatas() {
sheet_name_list.forEach(function(y) {
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
let items = [];
function getFromHtml(options, trackingCode) {
    let defaultOption = {
        uri: ''
    };
    if(typeof options === 'string') {
        defaultOption.uri = options;
    } else if(typeof options === 'object') {
        defaultOption = _.extend(defaultOption, options);
    }
    return new Promise(function (resolve, reject) {
        request(defaultOption, function (error, response, body) {
            if(error) {
                reject(error);
            } else {
                if(response.statusCode !== 200) {
                    typeof defaultOption.strategy !== 'undefined' ? resolve(defaultOption.strategy) : reject(new Error('11error status code:' + response.statusCode));
                } else {
                    let dataa = body;
                    // console.log(body);
                    let aa = dataa.split('6(');
                    if (typeof aa[1] !== 'undefined') {
                        let bb = aa[1].substr(0, aa[1].length - 2);
                        let messages = JSON.parse(bb);
                        // console.log('messages:', messages);
                        let cc = decodeURIComponent(escape(messages.result));
                        let ddd = cheerio.load(cc, {normalizeWhitespace: true});
                        let $ = ddd;
                        if (typeof $('.track_row') !== 'undefined') {
                            console.log('ok');
                            $('.track_row').each(function (idx, element) {
                                var $element = $(element);
                                var childTd = $($element.children('div'));
                                if(typeof $(childTd['1']).text() !== 'undefined') {
                                    var dateTime = $(childTd['2']).text();
                                    var status = $(childTd['3']).text();
                                    if(status !== 0 || true) {
                                        items.push({
                                            trackingCode: trackingCode,
                                            dateTime : dateTime,
                                            status : $(childTd['1']).text()
                                        });
                                    }
                                } else {
                                    console.log('fail', trackingCode);
                                    items.push({
                                        trackingCode: trackingCode,
                                        dateTime : null,
                                        status : null
                                    });
                                }
                            });
                        } else {
                            console.log('fail', trackingCode);
                            items.push({
                                trackingCode: trackingCode,
                                dateTime : null,
                                status : null
                            });
                        }
                    } else {
                        console.log('fail', trackingCode);
                        items.push({
                            trackingCode: trackingCode,
                            dateTime : null,
                            status : null
                        });
                    }
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
                    XLSX.writeFile(wb, 'trackingFW3.xlsx');
                }
            }
        });
    });
}
new Promise(function (resolve, reject) {
    // return;
    // console.log(data);
    data.forEach(dat => {
        // return;
        let trackingCode = dat.tracking_numbers;
        // console.log(trackingCode);
        // return;
        getFromHtml({
            url: 'http://fastway.com.au/track.php?callback=jQuery11240984190144918168_1493263685156&LabelNo='+trackingCode+'&_=1493263685157'
        }, trackingCode);
    });
});
