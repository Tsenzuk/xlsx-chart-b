const {
  getColName,
} = require('../services');
const CONTENT_TYPES = require('../ContentTypes');
const RELATION_TYPES = require('../relationTypes');

let worksheetCounter = 1;

class Worksheet {
  constructor(name) {
    this.contentType = CONTENT_TYPES.worksheet;
    this.relationType = RELATION_TYPES.worksheet;
    this.id = worksheetCounter;
    this.fileName = `sheet${worksheetCounter}.xml`;
    this.name = `${+new Date()}`;
    this.content = JSON.parse(JSON.stringify(require('../../template/xl/worksheets/sheet1.xml.json')));
    this.name = name;

    worksheetCounter++;
  }

  /**
   *
   * @param {string} name
   */
  setWorksheetName(name) {
    this.name = name;
  }

  /**
   *
   * @param {{ titles: string[], fields: string[], data: Object.<string, Object.<string, number|string>> }} chartConfig
   */
  setWorksheetData(chartConfig, rowOffset = 0, columnOffset = 0) {
    const data = chartConfig.data;
    this._data = data;

    this.content.worksheet.sheetData.row = this.content.worksheet.sheetData.row || [];

    this.content.worksheet.sheetData.row.push({
      $: {
        r: 1 + rowOffset,
      },
      c: chartConfig[chartConfig.titlesField].map((columnName, columnIndex) => ({
        $: {
          r: `${getColName(columnIndex + 2 + columnOffset)}${1 + rowOffset}`,
          t: 'str',
        },
        v: [
          columnName,
        ],
      })),
    });

    chartConfig[chartConfig.fieldsField].forEach((rowName, rowIndex) => this.content.worksheet.sheetData.row.push({
      $: {
        r: rowIndex + 2 + rowOffset,
      },
      c: [
        {
          $: {
            r: `${getColName(1 + columnOffset)}${rowIndex + 2 + rowOffset}`,
            t: 'str',
          },
          v: [
            rowName,
          ],
        },
        ...chartConfig[chartConfig.titlesField].map((columnName, columnIndex) => ({
          $: {
            r: `${getColName(columnIndex + 2 + columnOffset)}${rowIndex + 2 + rowOffset}`,
          },
          v: [
            data[columnName][rowName].value || '',
          ],
        })),
      ],
    }));
  }

  getRelationship(relationshipId) {
    return {
      $: {
        'sheetId': `${this.id}`,
        'name': this.name,
        'state': 'visible',
        'r:id': relationshipId,
      },
    };
  }
}

module.exports = Worksheet;
