const fs = require('fs');
const xml2js = require('xml2js');
const xl = require('excel4node');

interface Task {
  id: string;
  parent_id: string;
  activity_id?: string;
  name: string;
  start_date?: string;
  end_date?: string;
}

function readXml() {
  try {
    return fs.readFileSync('./sample.xml', 'utf8');
  } catch (err) {
    console.error(err);
  }
}

function xmlToJson(xml: string) {
  const parser = new xml2js.Parser();
  return new Promise((resolve, reject) => {
    parser.parseString(xml, (err: any, result: any) => {
      if (err) throw err;
      resolve(result);
    });
  });
}

function makeTasks(json: any): Task[] {
  const projects = Object.values(json).map((it: any) => it.Project).flat();

  const WBSArr = projects.map((it: any) => it.WBS).filter(Boolean).flat();
  const ActivityArr = projects.map((it: any) => it.Activity).filter(Boolean).flat();

  const tasks: Task[] = [];


  ActivityArr.forEach((activity: any) => {
    const task = {
      id: activity.ObjectId?.toString(),
      parent_id: activity.WBSObjectId?.toString(),
      name: activity.Name?.toString(),
      start_date: activity.StartDate?.toString(),
      end_date: activity.FinishDate?.toString(),
      activity_id: activity.Id?.toString()
    };
    tasks.push(task);
  });

  WBSArr.forEach((wbs: any) => {
    const task = {
      id: wbs.ObjectId?.toString(),
      parent_id: wbs.ParentObjectId?.toString(),
      name: wbs.Name?.toString(),
      start_date: new Date().toString(),
      end_date: new Date().toString()
    };
    tasks.push(task);
  });


  return tasks;
}

function saveToXlsx(tasks: Task[]): void {
  const wb = new xl.Workbook();

  const ws = wb.addWorksheet('Sheet 1');

  addHead(ws);
  tasks.forEach((task, index) => {
    addRow(ws, task, index + 2);
  });

  wb.write('MyExcelFile.xlsx');

}

function addHead(ws: any) {
  ws.cell(1, 1).string('Activity ID');
  ws.cell(1, 2).string('Task Name');
  ws.cell(1, 3).string('Start');
  ws.cell(1, 4).string('End');
  ws.cell(1, 5).string('ID');
  ws.cell(1, 6).string('Parent ID');
  return ws;
}

/**
 * @param ws
 * @param task
 * @param row 2부터 시작
 */
function addRow(ws: any, task: Task, row: number) {
  ws.cell(row, 1).string(task.activity_id ?? '(not set)');
  ws.cell(row, 2).string(task.name);
  ws.cell(row, 3).string(task.start_date);
  ws.cell(row, 4).string(task.end_date);
  ws.cell(row, 5).string(task.id);
  if (!isNaN(Number(task.parent_id))) ws.cell(row, 6).string(task.parent_id);
  return ws;
}

async function main() {
  const xml = readXml();
  const result: any = await xmlToJson(xml);
  const tasks = makeTasks(result);
  saveToXlsx(tasks);
}

main();