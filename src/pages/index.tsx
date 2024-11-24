import { useState } from 'react';
import { handleFile } from './utils';
import { Button, Card, Divider, PageHeader, Table, Upload } from 'antd';
// import result from './resut.json';
import './index.less';
import * as XLSX from 'xlsx';

export default function IndexPage() {
  const [columns, setColumns] = useState<any[]>([]);
  const [dataSource, setDataSource] = useState<any[]>([]);
  const [sortedObject, setSortedObject] = useState<any>();
  const [titleColumn, setTitleColumn] = useState<any>();

  const formatColumns = (columnsData: any) => {
    const keys = Object.keys(columnsData);
    // 自定义排序函数
    const sortedKeys = keys.sort((a, b) => {
      const numA = parseInt(a.replace(/__EMPTY_?/, ''), 10);
      const numB = parseInt(b.replace(/__EMPTY_?/, ''), 10);
      return numA - numB;
    });
    const sortedObject = sortedKeys.reduce((acc, key) => {
      acc[key] = columnsData[key];
      return acc;
    }, {} as typeof columnsData);
    setSortedObject(sortedObject);
    console.log(sortedObject);

    const formattedColumns = Object.entries(sortedObject).map(
      ([key, value]) => {
        return {
          title: value,
          dataIndex: key,
          key: key,
          render: (text: any) => (
            <span
              title={text}
              style={{
                whiteSpace: 'nowrap',
                overflow: 'hidden',
                textOverflow: 'ellipsis',
              }}
            >
              {text}
            </span>
          ),
        };
      },
    );
    setColumns(formattedColumns);
  };

  const exportExcel = () => {
    const newData = [titleColumn, sortedObject, ...dataSource];
    const wb = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.json_to_sheet(newData, {
      skipHeader: true,
    });
    XLSX.utils.book_append_sheet(wb, newWorksheet, 'Sheet1');
    XLSX.writeFile(wb, 'output.xlsx');
  };

  return (
    <div>
      <h1>Hello World</h1>
      <input
        type="file"
        accept=".xls,.xlsx"
        onChange={async (k) => {
          handleFile(k.target.files?.[0], (data) => {
            if (data) {
              // [标题, 表头, [数据源]]
              setTitleColumn(data[0]);
              formatColumns(data[1]);
              setDataSource(data.slice(2));
              console.log('月度总结', data);
            } else {
              setDataSource([]);
            }
          });
        }}
      ></input>
      <Divider />
      <Button onClick={exportExcel} type="primary" style={{ marginBottom: 16 }}>
        导出文件
      </Button>
      <Table
        title={() => '月度总结'}
        pagination={{ pageSize: 50 }}
        dataSource={dataSource}
        columns={columns}
      ></Table>
    </div>
  );
}
