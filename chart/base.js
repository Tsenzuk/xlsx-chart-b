const JSZip = require('jszip');
const xml2js = require('xml2js');

const {
  normalizeOptions,
  validateOptions,
} = require('./normalization');

const CONTENT_TYPES = require('./ContentTypes');

const Worksheet = require('./components/Worksheet');
const Chart = require('./components/Chart');
const Drawing = require('./components/Drawing');
const Relationship = require('./components/Relationship');

/**
 * @typedef {import('../chart').XLSXChartOptions} XLSXChartOptions
 */

class ExcelFile {
  /**
   * @param {XLSXChartOptions} options
   */
  constructor(options) {
    this.zip = new JSZip();
    // this.xmlBuilder = new xml2js.Builder({ renderOpts: { indent: '', newline: '' } }); // use this to disable xml formatting
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

    this.options = normalizeOptions(options);
  }

  /**
   * @param {XLSXChartOptions} options
   */
  writeFile(options) {
    const { type } = options;

    this.fillWorksheets(options);

    this.addFiles();

    return this.zip.generateAsync({ type });
  }

  addItem(item) {
    switch (item.contentType) {
      case CONTENT_TYPES.worksheet: {
        this.document.worksheets.push(item);
        const relationship = new Relationship('worksheet');

        this.relationships.contentTypes.Types.Override.push(relationship.getContentTypes(item.fileName));
        this.relationships.workbook.Relationships.Relationship.push(relationship.getRelationship(item.fileName));
        this.document.workbook.workbook.sheets.sheet.push(item.getRelationship(relationship.relationshipId));

        this.relationships.worksheets.push(relationship);
        break;
      }
      case CONTENT_TYPES.drawing: {
        this.document.drawings.push(item);
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
        if (this.options.dataPerSheet) {
          this.relationships.drawings[item.id - 1].content.Relationships.Relationship.push(relationship.getRelationship(item.fileName));
          this.document.drawings[item.id - 1].content['xdr:wsDr']['xdr:twoCellAnchor'].push(item.getRelationship(relationship.relationshipId));
        } else {
          this.relationships.drawings[0].content.Relationships.Relationship.push(relationship.getRelationship(item.fileName));
          this.document.drawings[0].content['xdr:wsDr']['xdr:twoCellAnchor'].push(item.getRelationship(relationship.relationshipId));
        }
      }
    }
  }

  /**
   * @param {XLSXChartOptions} options
   */
  fillWorksheets(options) {
    let rowOffset = 0;

    if (options.dataPerSheet) {
      options.charts.forEach((chartConfig) => {
        const { chartTitle } = chartConfig;

        const worksheet = new Worksheet(chartTitle);
        this.addItem(worksheet);

        const drawing = new Drawing();
        this.addItem(drawing);
      });
    } else {
      let worksheet = new Worksheet('Charts');
      this.addItem(worksheet);
      const drawing = new Drawing();
      this.addItem(drawing);

      worksheet = new Worksheet('Table');
      this.addItem(worksheet);
    }

    options.charts.forEach((chartConfig, chartIndex) => {
      let worksheet;
      let drawing;

      const {
        fields,
        chartTitle,
      } = chartConfig;

      if (options.dataPerSheet) {
        worksheet = this.document.worksheets[chartIndex];
        drawing = this.document.drawings[chartIndex];
      } else {
        worksheet = this.document.worksheets[1]; // the second worksheet is for data
        drawing = this.document.drawings[0]; // the single drawing is for all charts
      }

      worksheet.setWorksheetData(chartConfig, rowOffset);

      const chart = new Chart(worksheet.name, chartConfig);

      chart.setChartName(chartTitle);
      chart.setChartData(chartConfig, rowOffset);

      drawing.charts.push(chart);
      this.addItem(chart);

      if (!options.dataPerSheet) {
        rowOffset += fields.length + 2; // +2 for header and empty row
      }
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
        .file(fileName, this.xmlBuilder.buildObject(content)));
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

  /**
   *
   * @param {XLSXChartOptions} options
   */
  generate(options) {
    // option normalization, as it could be done in constructor
    if (options) {
      this.options = normalizeOptions(options);
    }

    const errors = validateOptions(this.options);

    if (errors) {
      return Promise.reject(errors);
    }

    return this.writeFile(this.options);
  }
}

module.exports = ExcelFile;
