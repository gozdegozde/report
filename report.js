import { MongoClient } from 'mongodb';
import Excel from 'exceljs';

async function createReport()
{
    const client = new MongoClient('mongodb://127.0.0.1:27017');
    const db = client.db('bin');
    let collection = db.collection('restaurants_data');

    const workbook = new Excel.Workbook();
    const sheet = workbook.addWorksheet('restaurant data');
    addHeader(sheet);
    await addCells(collection, sheet);
    await workbook.xlsx.writeFile('Restaurant report.xlsx');
    console.log('Report is done!')



}
async function addCells(collection, worksheet)
{
    const objects = await collection.find({});
    for await (let record of objects)
    {
        let row = [];
        row = [
            record.restaurant_id,
            record.name,
            record.cuisine,
            record.address.street + ' ' + record.address.building + ', ' + record.address.zipcode,
            averageScore(record.grades)
        ];
        worksheet.addRow(row).commit;
    }
    console.log('All rows are added!');

}
function averageScore(scores)
{
    let total = 0;
    let scoreList = [];
    let average;
    for (let s of scores)
    {
        scoreList.push(s.score);
        total += s.score;
    }

    average = total/scoreList.length;
    return average
}

function addHeader(worksheet)
{
    worksheet.getCell('A1').value = "Restaurant ID";
    worksheet.getCell('B1').value = "Restaurant Name";
    worksheet.getCell('C1').value = "Restaurant Cuisine";
    worksheet.getCell('D1').value = "Restaurant Address";
    worksheet.getCell('E1').value = "Restaurant Average Score";
    console.log('All headers are added!');

}


createReport();



