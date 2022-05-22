const JSZip = require('jszip');
const xml2js = require('xml2js');

const { getRowNames } = require('./services');

const CONTENT_TYPES = require('./ContentTypes');
const RELATION_TYPES = require('./relationTypes');

const Worksheet = require('./components/Worksheet');
const Chart = require('./components/Chart');
const Drawing = require('./components/Drawing');
const Relationship = require('./components/Relationship');

/**
 * @typedef {Object} ChartOptions
 * @property {string} type
 */

class ExcelFile {
  constructor(options) {
    this.zip = new JSZip();
    // this.xmlBuilder = new xml2js.Builder({ renderOpts: { indent: '', newline: '' } });
    this.xmlBuilder = new xml2js.Builder();

    this.document = {
      workbook: require('../template/xl/workbook.xml.json'),
      worksheets: [],
      drawings: [],
    };

    this.relationships = {
      contentTypes: require('../template/[Content_Types].xml.json'),
      root: require('../template/_rels/.rels.json'),
      workbook: require('../template/xl/_rels/workbook.xml.rels.json'),
      worksheets: [],
      drawings: [],
    };

    this.metadata = {
      app: require('../template/docProps/app.xml.json'),
      core: require('../template/docProps/core.xml.json'),
      styles: require('../template/xl/styles.xml.json'),
      theme: require('../template/xl/theme/theme1.xml.json'),
    };

    this.options = {
      type: 'nodebuffer',
      ...options,
    };
  }

  writeFile() {
    const { type } = this.options;

    this.fillWorksheets();

    this.addFiles();

    return this.zip.generateAsync({ type });
  }

  addItem(item) {
    switch (item.contentType) {
      case CONTENT_TYPES.worksheet: {
        const relationship = new Relationship('worksheet');

        this.relationships.contentTypes.Types.Override.push(relationship.getContentTypes(item.fileName));
        this.relationships.workbook.Relationships.Relationship.push(relationship.getRelationship(item.fileName));
        this.document.workbook.workbook.sheets.push(item.getRelationship(relationship.relationshipId));

        this.relationships.worksheets.push(relationship);
        break;
      }
      case CONTENT_TYPES.drawing: {
        const relationship = new Relationship('drawing');

        this.relationships.contentTypes.Types.Override.push(relationship.getContentTypes(item.fileName));
        this.relationships.worksheets[item.id - 1].content.Relationships.Relationship.push(relationship.getRelationship(item.fileName));
        this.document.worksheets[item.id - 1].content.worksheet.drawing.push(item.getRelationship(relationship.relationshipId));

        this.relationships.drawings.push(relationship);
        break;
      }
      case CONTENT_TYPES.chart: {
        const relationship = new Relationship('chart');

        this.relationships.contentTypes.Types.Override.push(relationship.getContentTypes(item.fileName));
        if (this.dataPerSheet) {
          this.relationships.drawings[item.id - 1].content.Relationships.Relationship.push(relationship.getRelationship(item.fileName));
          this.document.drawings[item.id - 1].content['xdr:wsDr']['xdr:twoCellAnchor'].push(item.getRelationship(relationship.relationshipId));
        } else {
          this.relationships.drawings[0].content.Relationships.Relationship.push(relationship.getRelationship(item.fileName));
          this.document.drawings[0].content['xdr:wsDr']['xdr:twoCellAnchor'].push(item.getRelationship(relationship.relationshipId));
        }
      }
    }
  }

  fillWorksheets() {
    let worksheet;
    let drawing;
    let rowOffset = 0;

    if (!this.dataPerSheet) {
      worksheet = new Worksheet();
      worksheet.setWorksheetName('Table');
      this.document.worksheets.push(worksheet);
      this.addItem(worksheet);

      drawing = new Drawing();
      this.document.drawings.push(drawing);
      this.addItem(drawing);
    }

    this.options.charts.forEach((chartConfig, chartIndex) => {
      const { data } = chartConfig;
      let { chartTitle } = chartConfig;
      chartTitle = chartTitle || `Chart ${chartIndex}`;

      if (this.dataPerSheet) {
        worksheet = new Worksheet();
        worksheet.setWorksheetName(chartTitle);
      }
      worksheet.setWorksheetData(data, rowOffset);
      if (this.dataPerSheet) {
        this.document.worksheets.push(worksheet);
        this.addItem(worksheet);

        drawing = new Drawing();
        this.document.drawings.push(drawing);
        this.addItem(drawing);
      }

      const chart = new Chart('Table', chartConfig);

      chart.setChartName(chartTitle);
      chart.setChartData(data, rowOffset);

      drawing.charts.push(chart);
      this.addItem(chart);

      if (!this.dataPerSheet) {
        rowOffset += getRowNames(data).length + 2;
      }

    });
  }

  fillWorkbook() {
    // write worksheets
    this.document.worksheets.forEach((worksheet, sheetIndex) => {
      worksheet.relationshipId = `rId${this.relationships.workbook.Relationships.Relationship.length + 1}`;

      this.relationships.contentTypes.Types.Override.push({
        $: {
          PartName: `/xl/worksheets/${worksheet.fileName}`,
          ContentType: CONTENT_TYPES.worksheet,
        },
      });

      const relationshipId = `rId${this.relationships.workbook.Relationships.Relationship.length + 1}`;

      this.relationships.workbook.Relationships.Relationship.push({
        $: {
          Id: relationshipId,
          Type: RELATION_TYPES.worksheet,
          Target: `worksheets/${worksheet.fileName}`,
        },
      });

      this.document.workbook.workbook.sheets.push({
        sheet: [
          {
            $: {
              'sheetId': `${worksheet.id}`,
              'name': this.document.worksheets[sheetIndex].name,
              'state': 'visible',
              'r:id': relationshipId,
            },
          },
        ],
      });
    });
  }

  addFiles() {
    this.document.worksheets.forEach(({ fileName, content }) => this.zip
      .folder('xl')
      .folder('worksheets')
      .file(fileName, this.xmlBuilder.buildObject(content)));

    this.document.drawings.forEach(({ fileName, content, charts }) => {
      this.zip
        .folder('xl')
        .folder('drawings')
        .file(fileName, this.xmlBuilder.buildObject(content));

      charts.forEach(({ fileName, content }) => this.zip
        .folder('xl')
        .folder('charts')
        .file(fileName, (console.log(content), this.xmlBuilder.buildObject(content))));
    });

    this.zip
      .folder('xl')
      .file('workbook.xml', this.xmlBuilder.buildObject(this.document.workbook));

    this.zip
      .folder('_rels')
      .file('.rels', this.xmlBuilder.buildObject(this.relationships.root));

    this.zip
      .folder('xl')
      .folder('_rels')
      .file('workbook.xml.rels', this.xmlBuilder.buildObject(this.relationships.workbook));

    this.relationships.worksheets.forEach(({ content }, index) => {
      const fileName = `sheet${index + 1}.xml.rels`;
      this.zip
        .folder('xl')
        .folder('worksheets')
        .folder('_rels')
        .file(fileName, this.xmlBuilder.buildObject(content));
    });

    this.relationships.drawings.forEach(({ content }, index) => {
      const fileName = `drawing${index + 1}.xml.rels`;
      this.zip
        .folder('xl')
        .folder('drawings')
        .folder('_rels')
        .file(fileName, this.xmlBuilder.buildObject(content));
    });

    this.zip.file('[Content_Types].xml', this.xmlBuilder.buildObject(this.relationships.contentTypes));

    this.zip.folder('docProps').file('app.xml', this.xmlBuilder.buildObject(this.metadata.app));
    this.zip.folder('docProps').file('core.xml', this.xmlBuilder.buildObject(this.metadata.core));

    this.zip.folder('xl').file('styles.xml', this.xmlBuilder.buildObject(this.metadata.styles));

    this.zip.folder('xl').folder('theme').file('theme1.xml', this.xmlBuilder.buildObject(this.metadata.theme));
  }

  validateOptions() {
    const {
      type,
      charts,
    } = this.options;

    const errors = [];

    if (typeof type !== 'string') {
      errors.push(new Error(`options.type should be string or missing, passed: ${type}`));
    }
    if (!Array.isArray(charts)) {
      errors.push(new Error(`options.charts should be array, passed: ${charts}`));
    }

    charts?.forEach((chart, index) => {
      if (!chart.data || typeof chart.data !== 'object') {
        errors.push(new Error(`each options.charts[${index}].data should contain object, passed: ${chart.data}`));
      }
      if (chart.data && Object.values(chart.data).some((subData) => typeof subData !== 'object')) {
        errors.push(new Error(`each options.charts[${index}].data[title] should contain object`));
      }
    });

    return errors.length ? errors : null;
  }

  /**
   *
   * @param {ChartOptions} options
   */
  generate(options) {
    this.options = {
      type: 'nodebuffer',
      ...this.options,
      ...options,
    };

    const errors = this.validateOptions();

    if (errors) {
      return Promise.reject(errors);
    }

    return this.writeFile();
  }
}

module.exports = ExcelFile;
