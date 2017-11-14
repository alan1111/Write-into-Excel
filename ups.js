var fs = require('fs');
const moment = require('moment');
const _ = require('lodash');
const cheerio = require('cheerio');
var request = require('request');
var XLSX = require('xlsx');
var workbook = XLSX.readFile('trackingUps.ods');
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
                    $('.dataTable tr').each(function (idx, element) {
                        var $element = $(element);
                        var childTd = $($element.children('td'));
                        if ($(childTd['1']).text() !== '') {
                            // var dateFormat = [
                            //     'DD.MM.YYYY h.mm',
                            //     'DD.MM.YYYY'
                            // ];
                            // var dateTime = moment(_.trim($(childTd['1']).text()) + ' ' +_.trim($(childTd['2']).text()), dateFormat).format();
                            let dateTime = $(childTd['1']).text() + ' ' +$(childTd['2']).text();
                            // var status = mapping(_.trim($(childTd['3']).text()));
                            var status = $(childTd['3']).text();
                            if (status !== 0 || true) {
                                items.push({
                                    trackingCode : trackingCode,
                                    dateTime : dateTime,
                                    status : $(childTd['3']).text()
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
                }
            }
        });
    });
}
new Promise(function (resolve, reject) {
    // return;
    data.forEach(dat => {
        // return;
        let trackingCode = dat.tracking_numbers;
        // return;
        getFromHtml({
            request: {
                url: 'http://wwwapps.ups.com/WebTracking/track?HTMLVersion=5.0&loc=_&Requester=UPSHome&WBPM_lid=homepage%2Fct1.html_pnl_trk&trackNums='+trackingCode+'&track.x=Traceren'
            }
        }, trackingCode);
    });
});
