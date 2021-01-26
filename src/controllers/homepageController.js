import excel from "excel4node";
import fetch from "node-fetch";

let getData = async (name) => {
    let data = [];
    let d1 = [], d2 = [], d3 = [], d4 = [], d5 = [], d6 = [];

    await fetch(`https://api.npmjs.org/downloads/range/2015-01-01:2015-12-31/${name}`)
        .then(res => res.json())
        .then(d => d1 = d.downloads);

    await fetch(`https://api.npmjs.org/downloads/range/2016-01-01:2016-12-31/${name}`)
        .then(res => res.json())
        .then(d => d2 = d.downloads);

    await fetch(`https://api.npmjs.org/downloads/range/2017-01-01:2017-12-31/${name}`)
        .then(res => res.json())
        .then(d => d3 = d.downloads);

    await fetch(`https://api.npmjs.org/downloads/range/2018-01-01:2018-12-31/${name}`)
        .then(res => res.json())
        .then(d => d4 = d.downloads);

    await fetch(`https://api.npmjs.org/downloads/range/2019-01-01:2019-12-31/${name}`)
        .then(res => res.json())
        .then(d => d5 = d.downloads);

    await fetch(`https://api.npmjs.org/downloads/range/2020-01-01:2021-12-31/${name}`)
        .then(res => res.json())
        .then(d => d6 = d.downloads);

    data = d1.concat(d2).concat(d3).concat(d4).concat(d5).concat(d6);
    return data;
}

let getHomepage = async (req, res) => {
    const PACHKAGE = "vue";
    const FILENAME = "vue5";
    let dataToWrite = await getData(PACHKAGE);

    const workbook = new excel.Workbook();
    const style = workbook.createStyle({
        font: { color: "#0101FF", size: 11 }
    });

    const worksheet = workbook.addWorksheet("Sheet 1");


    const arrayToWrite = Array.from({ length: 2 }, (v, k) => [`Row ${k + 1}, Col 1`, `Row ${k + 1}, Col 2`]);
    // arrayToWrite.forEach((row, rowIndex) => {
    //     row.forEach((entry, colIndex) => {
    //         worksheet.cell(rowIndex + 1, colIndex + 1).string(entry).style(style);
    //     })
    // })

    let arrDate = [
        '2015-01-01', '2015-02-01', '2015-03-01', '2015-04-01', '2015-05-01', '2015-06-01', '2015-07-01',
        '2015-08-01', '2015-09-01', '2015-10-01', '2015-11-01', '2015-12-01',

        '2016-01-01', '2016-02-01', '2016-03-01', '2016-04-01', '2016-05-01', '2016-06-01', '2016-07-01',
        '2016-08-01', '2016-09-01', '2016-10-01', '2016-11-01', '2016-12-01',

        '2017-01-01', '2017-02-01', '2017-03-01', '2017-04-01', '2017-05-01', '2017-06-01', '2017-07-01',
        '2017-08-01', '2017-09-01', '2017-10-01', '2017-11-01', '2017-12-01',

        '2018-01-01', '2018-02-01', '2018-03-01', '2018-04-01', '2018-05-01', '2018-06-01', '2018-07-01',
        '2018-08-01', '2018-09-01', '2018-10-01', '2018-11-01', '2018-12-01',

        '2019-01-01', '2019-02-01', '2019-03-01', '2019-04-01', '2019-05-01', '2019-06-01', '2019-07-01',
        '2019-08-01', '2019-09-01', '2019-10-01', '2019-11-01', '2019-12-01',

        '2020-01-01', '2020-02-01', '2020-03-01', '2020-04-01', '2020-05-01', '2020-06-01', '2020-07-01',
        '2020-08-01', '2020-09-01', '2020-10-01', '2020-11-01', '2020-12-01',

        '2021-01-01', '2021-02-01', '2021-03-01', '2021-04-01', '2021-05-01',

    ]
    arrayToWrite.forEach((row, rowIndex) => {
        if (rowIndex === 0) {
            dataToWrite.map((item, index) => {
                worksheet.cell(1, index + 1).string(item.day).style(style);
                if (!arrDate.includes(item.day)) {
                    worksheet.cell(4, index + 1).string('delete').style(style);
                }
            })
        }
        if (rowIndex === 1) {
            dataToWrite.map((item, index) => {
                worksheet.cell(2, index + 1).number(item.downloads).style(style);
            })
        }
    })


    await workbook.write(`${FILENAME}.xlsx`);

    return res.render("homepage.ejs");
};


module.exports = {
    getHomepage: getHomepage
};
