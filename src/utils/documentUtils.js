import { read, utils } from 'xlsx';

export const parseExcel = file => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = utils.sheet_to_json(worksheet, { header: 1 });

        if (jsonData.length < 3) {
          reject(
            new Error(
              'Excel表格格式不正确，需要包含描述行、表头行和至少一行数据'
            )
          );
          return;
        }

        const headers = jsonData[1];
        const rows = jsonData
          .slice(2)
          .filter(row =>
            row.some(cell => cell !== undefined && cell !== null && cell !== '')
          );

        if (rows.length === 0) {
          reject(new Error('Excel表格中未检测到有效数据行'));
          return;
        }

        const formattedRows = rows.map(row => {
          const obj = {};
          headers.forEach((header, index) => {
            let value = row[index] || '';

            if (typeof value === 'number') {
              if (value > 25569) {
                const date = new Date((value - 25569) * 86400000);
                value = `${date.getFullYear()}年${(date.getMonth() + 1)
                  .toString()
                  .padStart(2, '0')}月${date
                  .getDate()
                  .toString()
                  .padStart(2, '0')}日`;
              } else if (value > 0 && value < 1) {
                const hours = Math.floor(value * 24);
                const minutes = Math.floor((value * 24 * 60) % 60);
                value = `${hours.toString().padStart(2, '0')}:${minutes
                  .toString()
                  .padStart(2, '0')}`;
              }
            }

            obj[header] = value;
          });

          return obj;
        });

        resolve(formattedRows);
      } catch (error) {
        reject(new Error('Excel文件解析失败'));
      }
    };
    reader.onerror = () => reject(new Error('文件读取失败'));
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
              selectedValue.includes(nextText)
          );

          if (isSelected) {
            const checkboxIndex = current.fullMatch.indexOf('□');
            const absoluteIndex = current.index + checkboxIndex;
            result[absoluteIndex] = '☑';
          }
        }
      }

      return result.join('');
    }
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
        reject(new Error('Word模板读取失败'));
      }
    };
    reader.onerror = () => reject(new Error('文件读取失败'));
    reader.readAsArrayBuffer(file);
  });
};
