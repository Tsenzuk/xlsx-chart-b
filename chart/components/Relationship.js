const CONTENT_TYPES = require('../ContentTypes');
const RELATION_TYPES = require('../relationTypes');

let relationshipCounter = 3;

class Relationship {
  constructor(type) {
    this.id = relationshipCounter;
    this.relationshipId = `rId${this.id}`;
    this.type = type;
    this.contentType = CONTENT_TYPES[type];
    this.relationType = RELATION_TYPES[type];

    switch (this.contentType) {
      case CONTENT_TYPES.worksheet:
        this.content = JSON.parse(JSON.stringify(require('../../template/xl/worksheets/_rels/sheet1.xml.rels.json')));
        this.partNameBase = '/xl/worksheets/';
        break;
      case CONTENT_TYPES.drawing:
        this.content = JSON.parse(JSON.stringify(require('../../template/xl/drawings/_rels/drawing1.xml.rels.json')));
        this.partNameBase = '/xl/drawings/';
        break;
      case CONTENT_TYPES.chart:
        this.content = JSON.parse(JSON.stringify(require('../../template/xl/charts/_rels/chart1.xml.rels.json')));
        this.partNameBase = '/xl/charts/';
        break;
      default:
        console.error(`Relationship ${this.type} is not supported`);
    }
    relationshipCounter++;
  }

  getContentTypes(fileName) {
    return {
      $: {
        PartName: `${this.partNameBase}${fileName}`,
        ContentType: this.contentType,
      },
    };
  }

  getRelationship(fileName) {
    return {
      $: {
        Id: this.relationshipId,
        Type: this.relationType,
        Target: `${this.partNameBase}${fileName}`, // usually Excel creates relative paths, but absolute works too
      },
    };
  }
}

module.exports = Relationship;
