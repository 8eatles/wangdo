const xlsx = require("xlsx");
const Workbook = require("xlsx-workbook").Workbook;

const args = process.argv.slice(2);

const filename = args[0] || "./data.xlsx";

const book = xlsx.readFile(filename);
const [dataSheetName] = book.SheetNames;
const dataSheet = book.Sheets[dataSheetName];

const data = {};
let i = 2;

while (true) {
  const no = dataSheet[`B${i}`].v;
  const row = {
    startTime: dataSheet[`A${i}`].v,
    veh: dataSheet[`C${i}`].v,
    capa: dataSheet[`D${i}`].v,
    rev: dataSheet[`E${i}`].v,
    tq: dataSheet[`F${i}`].v,
    vel: dataSheet[`G${i}`].v,
    toil: dataSheet[`H${i}`].v,
    triplength: dataSheet[`I${i}`].v,
    idle: dataSheet[`J${i}`].v,
    maxMile: dataSheet[`K${i}`].v,
    minMile: dataSheet[`L${i}`].v,
    x: dataSheet[`M${i}`].v,
  };

  if (data.hasOwnProperty(no)) {
    data[no].startTime =
      data[no].startTime <= row.startTime ? data[no].startTime : row.startTime;
    data[no].endTime =
      data[no].startTime >= row.startTime
        ? data[no].endTime
        : row.startTime + row.triplength / 3600 / 24;
    data[no].startODO =
      data[no].startODO <= row.minMile ? data[no].startODO : row.minMile;
    data[no].endODO =
      data[no].endODO >= row.maxMile ? data[no].endODO : row.maxMile;
    data[no].rev += row.rev * row.triplength;
    data[no].tq += row.tq * row.triplength;
    data[no].vel += row.vel * row.triplength;
    data[no].toil += row.toil * row.triplength;
    data[no].idle += row.idle * row.triplength;
    data[no].totalTriplength += row.triplength;
    data[no].sumX += row.x;
    data[no].tripCnt += 1;
  } else {
    data[no] = {
      startTime: row.startTime,
      endTime: row.startTime + row.triplength / 3600 / 24,
      startODO: row.minMile,
      endODO: row.maxMile,
      rev: row.rev * row.triplength,
      tq: row.tq * row.triplength,
      vel: row.vel * row.triplength,
      toil: row.toil * row.triplength,
      idle: row.idle * row.triplength,
      totalTriplength: row.triplength,
      sumX: row.x,
      tripCnt: 1,
    };
  }

  i += 1;
  if (!dataSheet[`B${i}`]) break;
}

const workbook = new Workbook();
const results = workbook.add("results");

results[0][0] = "No";
results[0][1] = "Start Time";
results[0][2] = "End Time";
results[0][3] = "Start ODO";
results[0][4] = "End ODO";
results[0][5] = "REV";
results[0][6] = "TQ";
results[0][7] = "VEL";
results[0][8] = "TOIL";
results[0][9] = "IDLE";
results[0][10] = "Total Triplength";
results[0][11] = "SUM_X";
results[0][12] = "Trip Cnt";

Object.entries(data).forEach(
  (
    [
      key,
      {
        no,
        startTime,
        endTime,
        startODO,
        endODO,
        rev,
        tq,
        vel,
        toil,
        idle,
        totalTriplength,
        sumX,
        tripCnt,
      },
    ],
    i
  ) => {
    results[i + 1][0] = key;
    results[i + 1][1] = startTime;
    results[i + 1][2] = endTime;
    results[i + 1][3] = startODO;
    results[i + 1][4] = endODO;
    results[i + 1][5] = rev / totalTriplength;
    results[i + 1][6] = tq / totalTriplength;
    results[i + 1][7] = vel / totalTriplength;
    results[i + 1][8] = toil / totalTriplength;
    results[i + 1][9] = idle / totalTriplength;
    results[i + 1][10] = totalTriplength;
    results[i + 1][11] = sumX;
    results[i + 1][12] = tripCnt;
  }
);

workbook.save(`${new Date().valueOf()}_result.xlsx`);
