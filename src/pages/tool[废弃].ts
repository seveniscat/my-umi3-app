import * as XLSX from 'xlsx';
import moment, { Moment } from 'moment-timezone';

function handleJson(json: any[]) {
  let results: any = {};

  json.forEach((e) => {
    const name = e.__EMPTY;
    if (name) {
      const person = getPerson(e);
      if (person) {
        if (!results[name]) {
          results[name] = []; //{马小娟: []}
        }
        results[name].push(person);
      }
    }
  });
  // console.log(json)
  console.log(results);
}

function getPerson(e: any) {
  const applicationMoments = getApplicationMoment(e.__EMPTY_12);
  if (!applicationMoments.length) return;

  const person: any = {};
  person.name = e.__EMPTY;
  person.date = e.概况统计与打卡明细;
  person.account = e.__EMPTY_1;
  person.department = e.__EMPTY_2;
  person.rule = e.__EMPTY_5;
  person.earliest = e.__EMPTY_7;
  person.latest = e.__EMPTY_8;
  person.times = e.__EMPTY_9;
  person.stdworktime = e.__EMPTY_10;
  person.actworktime = e.__EMPTY_11;
  person.application = e.__EMPTY_12;

  // [
  //   "2024-11-01T00:30:00.000Z",
  //   "2024-11-05T10:38:00.000Z"
  // ]
  person.application = applicationMoments;

  const dateStr = person.date.split(' ')[0]; // 取出打卡日期
  if (dateStr) {
    // 生成打卡时间
    const dateFormat = 'YYYY/MM/DD HH:mm';
    if (!['--', '未打卡', ''].includes(person.earliest)) {
      const m = moment.tz(
        `${dateStr} ${person.earliest}`,
        dateFormat,
        'Asia/Shanghai',
      );
      if (m.isValid()) person.startMoment = m;
    }
    if (!['--', '未打卡'].includes(person.latest)) {
      const m = moment.tz(
        `${dateStr} ${person.latest}`,
        dateFormat,
        'Asia/Shanghai',
      );
      if (m.isValid()) person.endMoment = m;
    }
  }

  return person;
}

// application 转换为标准日期
function getApplicationMoment(application: string): Moment[] {
  const dateStr = extractBracketContent(application);
  if (!dateStr) return [];
  if (dateStr.indexOf('-') > 2) {
    // 11/5 上午 - 11/5 下午
    // 11/5 09:00 - 11/5 18:00
    // 11/8 14:00 - 11/15 18:00
    let [startStr, endStr] = dateStr.split('-');
    if (startStr.includes('上午') && endStr.includes('下午')) {
      startStr = startStr.replace('上午', '08:30');
      endStr = endStr.replace('下午', '17:00');
    } else {
      startStr = startStr.replace('上午', '08:30');
      endStr = endStr.replace('上午', '12:00');
      startStr = startStr.replace('下午', '13:00');
      endStr = endStr.replace('下午', '17:00');
    }
    // console.log(startStr, endStar)
    // 定义日期时间格式
    const dateFormat = 'MM/DD HH:mm';
    const start = moment.tz(startStr, dateFormat, 'Asia/Shanghai');
    const end = moment.tz(endStr, dateFormat, 'Asia/Shanghai');
    // console.log(startTime, endTime)

    if (start.isValid() && end.isValid()) {
      return [start, end];
    }
  } else {
    // 11-05 19:00
    // 定义日期时间格式
    // console.log(dateStr)

    const dateFormat = 'MM-DD HH:mm';
    const m = moment.tz(dateStr, dateFormat, 'Asia/Shanghai');
    console.log(m);
    if (m.isValid()) return [m];
  }

  return [];
}

function extractBracketContent(input: string): string | undefined {
  // 使用正则表达式匹配括号内的内容
  const regex = /[$\（][^$\）]+[\)\）]/g;
  let matches: string[] = [];
  let match;
  while ((match = regex.exec(input)) !== null) {
    // 去掉括号
    matches.push(match[0].slice(1, -1));
  }
  return matches[0];
}

export function test() {
  getApplicationMoment('--');

  getApplicationMoment('出差1.0天（11/5 上午 - 11/5 下午）');

  getApplicationMoment('补卡申请（11-05 19:00）');

  getApplicationMoment('外出9.0小时（11/5 09:00 - 11/5 18:00）');

  getApplicationMoment('事假1天19.0小时（11/8 14:00 - 11/15 18:00）');
}

export function handleFile(file: File | undefined) {
  const fr = new FileReader();

  fr.onload = (x) => {
    console.log(x);

    if (x.target && x.target.result) {
      const data = new Uint8Array(x.target.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });

      const sheetName = workbook.SheetNames?.[0];
      const sheet = workbook.Sheets[sheetName];
      const dataJson = XLSX.utils.sheet_to_json(sheet);
      handleJson(dataJson);
      console.log(dataJson);
    }
  };
  fr.readAsArrayBuffer(file!);
}
