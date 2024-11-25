import * as XLSX from 'xlsx';
import moment from 'moment-timezone';
moment.tz.setDefault('Asia/Shanghai');

function handleJson(json: any[]) {
  const obj = Object.assign({}, json[1], json[2]);
  // [标题, 表头, [数据源]]
  let results: any[] = [json[0], obj];

  json.forEach((e) => {
    const name = e.__EMPTY;
    if (name && name != '姓名') {
      const person: any = e;

      // 异常次数
      const exceptionTimes = Number(person.__EMPTY_14);
      if (!isNaN(exceptionTimes) && exceptionTimes > 0) {
        const [dateStr] = person.概况统计与打卡明细?.split(' '); //2024/11/17
        const format = 'YYYY/MM/DD HH:mm';
        const shouldStart = moment(`${dateStr} 8:30`, format);
        const shouldEnd = moment(`${dateStr} 17:00`, format);
        // 最早时间
        const earliest = moment(`${dateStr} ${person.__EMPTY_7}`, format);
        // 最晚时间
        const latest = moment(`${dateStr} ${person.__EMPTY_8}`, format);
        // 弹性时间逻辑
        if (earliest.isValid() && latest.isValid()) {
          const deltaStart = earliest.diff(shouldStart, 'minutes');
          const deltaEnd = latest.diff(shouldEnd, 'minutes');
          console.log(deltaStart, deltaEnd);
          if (
            (deltaStart >= -30 || deltaStart <= 30) &&
            deltaEnd > deltaStart
          ) {
            // 异常次数 -1
            person.__EMPTY_14 = exceptionTimes - 1;
          }
        }
      }
      // TODO 删除不必要的字段

      let current = results.find((res: any) => res.__EMPTY === name);
      if (!current) {
        results.push(person);
      } else {
        // 所有数字字段  求和
        Object.entries(current).forEach(([key, value]) => {
          let count = Number(value);
          let addCount = Number(person[key]);
          if (!isNaN(addCount)) {
            if (isNaN(count)) {
              count = 0;
            }
            count += addCount;
            if (!Number.isInteger(count)) {
              count = Number.parseFloat(count.toFixed(1));
            }
            current[key] = count;
          }
        });
      }
    }
  });
  console.log(results);
  return results;
}

export function handleFile(
  file: File | undefined,
  callBack: (data: any[]) => void,
) {
  if (!file) return;
  const fr = new FileReader();

  fr.onload = (x) => {
    if (x.target && x.target.result) {
      const fileData = new Uint8Array(x.target.result as ArrayBuffer);
      const workbook = XLSX.read(fileData, { type: 'array' });

      const sheetName = workbook.SheetNames?.[0];
      const worksheet = workbook.Sheets[sheetName];
      const dataJson = XLSX.utils.sheet_to_json<any>(worksheet);
      // 处理数据
      const newData = handleJson(dataJson);
      callBack(newData);
    } else {
      callBack([]);
    }
  };
  fr.onerror = (e) => {
    console.log(e);
    callBack([]);
  };
  fr.readAsArrayBuffer(file!);
}
