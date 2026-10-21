import { read, utils, SSF } from 'xlsx';

import { t } from '../i18n';

// 英文月份缩写，用于英文日期格式
const MONTH_SHORT_NAMES = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];

// 数据行从第3行开始（0基索引为2）
const DATA_START_ROW = 2;

// Excel日期序列号转Date（1900日期系统）
const serialToDate = serial => new Date((serial - 25569) * 86400000);

// 按当前语言格式化日期
const formatDate = date =>
  t('format.date', {
    year: date.getFullYear(),
    month: (date.getMonth() + 1).toString().padStart(2, '0'),
    day: date.getDate().toString().padStart(2, '0'),
    monthShort: MONTH_SHORT_NAMES[date.getMonth()],
  });

// Excel时间小数（0~1）转 HH:MM
const formatTime = timeValue => {
  const hours = Math.floor(timeValue * 24);
  const minutes = Math.floor((timeValue * 24 * 60) % 60);
  return `${hours.toString().padStart(2, '0')}:${minutes
    .toString()
    .padStart(2, '0')}`;
};

export const parseExcel = file => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        // cellNF: 读取单元格数字格式（cell.z），用于判断日期/时间单元格
        const workbook = read(data, { type: 'array', cellNF: true });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = utils.sheet_to_json(worksheet, { header: 1 });

        if (jsonData.length < 3) {
          reject(new Error(t('error.excelFormat')));
          return;
        }

        const headers = jsonData[1];
        // 过滤空行时保留工作表行号，便于按单元格地址回查格式
        const dataRows = [];
        for (let r = DATA_START_ROW; r < jsonData.length; r++) {
          const row = jsonData[r];
          if (
            row &&
            row.some(cell => cell !== undefined && cell !== null && cell !== '')
          ) {
            dataRows.push({ row, sheetRow: r });
          }
        }

        if (dataRows.length === 0) {
          reject(new Error(t('error.excelNoData')));
          return;
        }

        const formattedRows = dataRows.map(({ row, sheetRow }) => {
          const obj = {};
          headers.forEach((header, index) => {
            let value = row[index] ?? '';

            if (value instanceof Date) {
              value = formatDate(value);
            } else if (typeof value === 'number') {
              // 仅当单元格本身是日期/时间格式时才转换，
              // 避免把普通大数字（重量、编号等）误判为日期
              const cell =
                worksheet[utils.encode_cell({ r: sheetRow, c: index })];
              const cellFormat = cell && cell.z;

              if (cellFormat && SSF.is_date(cellFormat)) {
                if (value >= 0 && value < 1) {
                  value = formatTime(value);
                } else if (value >= 1) {
                  value = formatDate(serialToDate(value));
                }
              }
            }

            obj[header] = value;
          });

          return obj;
        });

        resolve(formattedRows);
      } catch (error) {
        reject(new Error(t('error.excelParseFailed')));
      }
    };
    reader.onerror = () => reject(new Error(t('error.fileReadFailed')));
    reader.readAsArrayBuffer(file);
  });
};

export const processCheckboxOptions = (content, rowData) => {
  let updatedContent = content;

  const checkboxPattern = /\{\{(check\d+)(.*?)\}\}/g;

  updatedContent = updatedContent.replace(
    checkboxPattern,
    (_, key, optionsPart) => {
      const excelValue = rowData[key];

      if (!excelValue) {
        return optionsPart;
      }

      const selectedValues = excelValue
        .toString()
        .split(/[，,、]/)
        .map(v => v.trim());

      if (selectedValues.length === 0) {
        return optionsPart;
      }

      const textTagPattern = /<w:t[^>]*>([^<]*)<\/w:t>/g;
      const allMatches = [];
      let matchResult;

      while ((matchResult = textTagPattern.exec(optionsPart)) !== null) {
        allMatches.push({
          fullMatch: matchResult[0],
          text: matchResult[1],
          index: matchResult.index,
        });
      }

      if (allMatches.length === 0) {
        return optionsPart;
      }

      const result = optionsPart.split('');

      for (let i = 0; i < allMatches.length; i++) {
        const current = allMatches[i];

        if (current.text.trim() === '□' && i + 1 < allMatches.length) {
          const next = allMatches[i + 1];
          const nextText = next.text.trim();

          const isSelected = selectedValues.some(
            selectedValue =>
              nextText.includes(selectedValue) ||
              selectedValue.includes(nextText),
          );

          if (isSelected) {
            const checkboxIndex = current.fullMatch.indexOf('□');
            const absoluteIndex = current.index + checkboxIndex;
            result[absoluteIndex] = '☑';
          }
        }
      }

      return result.join('');
    },
  );

  return updatedContent;
};

export const readWordTemplate = file => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const content = e.target.result;
        resolve(content);
      } catch (error) {
        reject(new Error(t('error.wordReadFailed')));
      }
    };
    reader.onerror = () => reject(new Error(t('error.fileReadFailed')));
    reader.readAsArrayBuffer(file);
  });
};
