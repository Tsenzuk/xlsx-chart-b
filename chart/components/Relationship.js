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
        break;
      case CONTENT_TYPES.drawing:
        this.content = JSON.parse(JSON.stringify(require('../../template/xl/drawings/_rels/drawing1.xml.rels.json')));
        break;
      default:
        console.error(`Relationship ${this.type} is not supported`);
    }
    relationshipCounter++;
  }

  getContentTypes(fileName) {
    switch (this.contentType) {
      case CONTENT_TYPES.worksheet:
        return {
          $: {
            PartName: `/xl/worksheets/${fileName}`,
            ContentType: this.contentType,
          },
        };
      case CONTENT_TYPES.drawing:
        return {
          $: {
            PartName: `/xl/drawings/${fileName}`,
            ContentType: this.contentType,
          },
        };
      case CONTENT_TYPES.chart:
        return {
          $: {
            PartName: `/xl/charts/${fileName}`,
            ContentType: this.contentType,
          },
        };
      default:
        console.error(`Content ${this.type} is not supported`);
    }
  }

  getRelationship(fileName) {
    switch (this.contentType) {
      case CONTENT_TYPES.worksheet:
        return {
          $: {
            Id: this.relationshipId,
            Type: this.relationType,
            Target: `worksheets/${fileName}`,
          },
        };
      case CONTENT_TYPES.drawing:
        return {
          $: {
            Id: this.relationshipId,
            Type: this.relationType,
            Target: `../drawings/${fileName}`,
          },
        };
      case CONTENT_TYPES.chart:
        return {
          $: {
            Id: this.relationshipId,
            Type: this.relationType,
            Target: `../charts/${fileName}`,
          },
        };
      default:
        console.error(`Relationship for ${this.type} is not supported`);
    }
  }
}

module.exports = Relationship;
