const CONTENT_TYPES = require('../ContentTypes');
const RELATION_TYPES = require('../relationTypes');

let drawingCounter = 1;

class Drawing {
  constructor() {
    this.contentType = CONTENT_TYPES.drawing;
    this.relationType = RELATION_TYPES.drawing;
    this.id = drawingCounter;
    this.fileName = `drawing${drawingCounter}.xml`;
    this.charts = [];
    this.content = JSON.parse(JSON.stringify(require('../../template/xl/drawings/drawing1.xml.json')));

    drawingCounter++;
  }

  getRelationship(relationshipId) {
    return {
      $: {
        'r:id': relationshipId,
      },
    };
  }
}

module.exports = Drawing;
