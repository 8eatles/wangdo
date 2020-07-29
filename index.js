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
    data[no].mile += row.maxMile - row.minMile;
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
      mile: row.maxMile - row.minMile,
    };
  }

  i += 1;
  if (!dataSheet[`B${i}`]) break;
}

const workbook = new Workbook();
const results = workbook.add("results");

results[0] = [
  "No",
  "Start Time",
  "End Time",
  "Start ODO",
  "End ODO",
  "REV",
  "TQ",
  "VEL",
  "TOIL",
  "IDLE",
  "Total Triplength",
  "SUM_X",
  "Trip Cnt",
  "Mile",
];

Object.entries(data).forEach(
  (
    [
      key,
      {
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
        mile,
      },
    ],
    i
  ) => {
    results[i + 1] = [
      key,
      startTime,
      endTime,
      startODO,
      endODO,
      rev / totalTriplength,
      tq / totalTriplength,
      vel / totalTriplength,
      toil / totalTriplength,
      idle / totalTriplength,
      totalTriplength,
      sumX,
      tripCnt,
      mile,
    ];
  }
);

workbook.save(`${new Date().valueOf()}_result.xlsx`);
