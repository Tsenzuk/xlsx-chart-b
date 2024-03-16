const {
  getColName,
} = require('../services');
const CONTENT_TYPES = require('../ContentTypes');
const RELATION_TYPES = require('../relationTypes');
const {
  DEFAULT_TITLES_FIELD,
  DEFAULT_FIELDS_FIELD,
  DEFAULT_LINE_COLOR_FIELD,
  DEFAULT_FILL_COLOR_FIELD,
  DEFAULT_MARKER_COLOR_FIELD,
} = require('../normalization');

const CHART_TAG_BY_CHART_NAME = {
  bar: 'c:barChart',
  column: 'c:barChart',
  line: 'c:lineChart',
  radar: 'c:radarChart',
  area: 'c:areaChart',
  scatter: 'c:scatterChart',
  pie: 'c:pieChart',
  doughnut: 'c:doughnutChart',
};

const CHART_GROUPING_BY_CHART_NAME = {
  bar: 'clustered',
  column: 'clustered',
  line: 'standard',
  radar: undefined, // radar, scatter, pie and doughnut charts should not have "c:grouping" tag, otherwise rest of xml is ignored
  area: 'standard',
  scatter: undefined,
  pie: undefined,
  doughnut: undefined,
};

// const CHART_TYPES = ['bar', 'column', 'line', 'radar', 'area', 'scatter', 'pie', 'doughnut'];

const DEFAULT_CHART_OPTIONS = {
  chart: 'bar',
  grouping: null,
  customColors: null,
  firstSliceAng: null,
  holeSize: null,
  showLabels: null,
  legendPos: null,
  manualLayout: null,
};

const DELIMITER = '\r\r';

let chartCounter = 1;

class Chart {
  constructor(dataWorksheetName = 'Table', chartOptions) {
    this.contentType = CONTENT_TYPES.chart;
    this.relationType = RELATION_TYPES.chart;
    this.id = chartCounter;
    this.fileName = `chart${chartCounter}.xml`;
    this.dataWorksheetName = dataWorksheetName;
    this.chartOptions = Object.assign({}, DEFAULT_CHART_OPTIONS, chartOptions);
    this.titlesField = this.chartOptions.titlesField || DEFAULT_TITLES_FIELD;
    this.fieldsField = this.chartOptions.fieldsField || DEFAULT_FIELDS_FIELD;
    this.lineColorField = this.chartOptions.lineColorField || DEFAULT_LINE_COLOR_FIELD;
    this.fillColorField = this.chartOptions.fillColorField || DEFAULT_FILL_COLOR_FIELD;
    this.markerColorField = this.chartOptions.markerColorField || DEFAULT_MARKER_COLOR_FIELD;

    this.name = `${+new Date()}`;
    this.content = JSON.parse(JSON.stringify(require('../../template/xl/charts/chart1.xml.json')));

    chartCounter++;
  }

  /**
   *
   * @param {string} title
   * @param {{ x: number, y: number }} [titlePosition]
   */
  setChartName(title, titlePosition = null) {
    const layout = {};

    if (titlePosition) {
      layout['c:manualLayout'] = {
        'c:xMode': { $: { val: 'edge' } },
        'c:yMode': { $: { val: 'edge' } },
        'c:x': { $: { val: titlePosition.x } },
        'c:y': { $: { val: titlePosition.y } },
      };
    }

    this.content['c:chartSpace']['c:chart']['c:title'] = {
      'c:tx': {
        'c:rich': {
          'a:bodyPr': {},
          'a:lstStyle': {},
          'a:p': {
            'a:pPr': {
              'a:defRPr': {},
            },
            'a:r': {
              'a:rPr': { $: { lang: 'en-US' } },
              'a:t': title,
            },
          },
        },
      },
      'c:layout': layout,
      'c:overlay': { $: { val: '0' } },
    };
  }

  setChartData(chartConfig, rowOffset = 0, columnOffset = 0) {
    // Group series by chart type and grouping
    // it is needed to gather all series of the same type and grouping together as single chart layer
    const seriesByChart = chartConfig[this.titlesField].reduce((seriesByChart, columnName, columnIndex) => {
      const columnData = chartConfig.data[columnName];
      const chart = columnData.chart || this.chartOptions.chart;
      const grouping = columnData.grouping || this.chartOptions.grouping || CHART_GROUPING_BY_CHART_NAME[chart];

      const customColorsPoints = {
        'c:dPt': [],
      };
      const customColorsSeries = {};

      customColorsPoints['c:dPt'] = chartConfig[this.fieldsField].map((field, i) => {
        const fillColor = chartConfig.data[columnName][field][this.fillColorField];
        const lineColor = chartConfig.data[columnName][field][this.lineColorField];
        const markerColor = chartConfig.data[columnName][field][this.markerColorField];

        const dataPointConfig = {
          'c:idx': { $: { val: i } },
        };

        if (!fillColor || !lineColor || !markerColor) {
          return null;
        }


        if (fillColor === 'noFill' && lineColor === 'noFill' && markerColor === 'noFill') {
          dataPointConfig['c:spPr'] = {
            'a:noFill': '',
          };

          return dataPointConfig;
        }

        dataPointConfig['c:spPr'] = {
          'a:solidFill': {
            'a:srgbClr': { $: { val: fillColor } },
          },
          'a:ln': {
            'a:solidFill': {
              'a:srgbClr': { $: { val: lineColor } },
            },
          },
          'c:marker': {
            'c:spPr': {
              'a:solidFill': {
                'a:srgbClr': { $: { val: markerColor } },
              },
            },
          },
        };

        return dataPointConfig;
      }).filter(Boolean);

      if (chartConfig.data[columnName][this.fillColorField]
        || chartConfig.data[columnName][this.lineColorField]
        || chartConfig.data[columnName][this.markerColorField]
      ) {
        const fillColor = chartConfig.data[columnName][this.fillColorField];
        const lineColor = chartConfig.data[columnName][this.lineColorField];
        const markerColor = chartConfig.data[columnName][this.markerColorField];

        customColorsSeries['c:spPr'] = {
          'a:solidFill': {
            'a:srgbClr': { $: { val: fillColor } },
          },
          'a:ln': {
            'a:solidFill': {
              'a:srgbClr': { $: { val: lineColor } },
            },
          },
          'c:marker': {
            'c:spPr': {
              'a:solidFill': {
                'a:srgbClr': { $: { val: markerColor } },
              },
            },
          },
        };
      }

      const ser = {
        'c:idx': { $: { val: columnIndex } },
        'c:order': { $: { val: columnIndex } },
        'c:tx': {
          'c:strRef': {
            'c:f': `'${this.dataWorksheetName}'!$${getColName(columnIndex + 2)}$${rowOffset + 1}`,
            'c:strCache': {
              'c:ptCount': { $: { val: 1 } },
              'c:pt': {
                '$': {
                  idx: 0,
                },
                'c:v': columnName,
              },
            },
          },
        },
        ...customColorsPoints,
        ...customColorsSeries,
        'c:cat': {
          'c:strRef': {
            'c:f': `'${this.dataWorksheetName}'!$${getColName(1 + columnOffset)}$${rowOffset + 2}:$${getColName(1 + columnOffset)}$${chartConfig[this.fieldsField].length + rowOffset + 1}`,
            'c:strCache': {
              'c:ptCount': { $: { val: chartConfig[this.fieldsField].length } },
              'c:pt': chartConfig[this.fieldsField].map((rowName, rowIndex) => {
                return {
                  '$': {
                    idx: rowIndex,
                  },
                  'c:v': rowName,
                };
              }),
            },
          },
        },
        'c:val': {
          'c:numRef': {
            'c:f': `'${this.dataWorksheetName}'!$${getColName(columnIndex + 2)}$${rowOffset + 2}:$${getColName(columnIndex + 2)}$${chartConfig[this.fieldsField].length + rowOffset + 1}`,
            'c:numCache': {
              'c:formatCode': 'General',
              'c:ptCount': { $: { val: chartConfig[this.fieldsField].length } },
              'c:pt': chartConfig[this.fieldsField].map((rowName, rowIndex) => {
                return {
                  '$': {
                    idx: rowIndex,
                  },
                  'c:v': columnData[rowName].value,
                };
              }),
            },
          },
        },
      };

      if (chart == 'scatter') {
        ser['c:xVal'] = ser['c:cat'];
        delete ser['c:cat'];
        ser['c:yVal'] = ser['c:val'];
        delete ser['c:val'];
        ser['c:spPr'] = {
          'a:ln': {
            '$': {
              w: 28575,
            },
            'a:noFill': '',
          },
        };
      }

      const seriesKey = `${chart}${DELIMITER}${grouping}`;
      seriesByChart[seriesKey] = seriesByChart[seriesKey] || [];
      seriesByChart[seriesKey].push(ser);

      return seriesByChart;
    }, {});


    Object.entries(seriesByChart).forEach(([chartAndGrouping, ser]) => {
      const [chart, grouping] = chartAndGrouping.split(DELIMITER);

      const chartTagName = CHART_TAG_BY_CHART_NAME[chart];

      if (!chartTagName) {
        throw new Error(`Chart type '${chart}' is not supported`);
      }

      // this.content['c:chartSpace']['c:chart']['c:plotArea'] = this.content['c:chartSpace']['c:chart']['c:plotArea'] || {};
      this.content['c:chartSpace']['c:chart']['c:plotArea'][chartTagName] = this.content['c:chartSpace']['c:chart']['c:plotArea'][chartTagName] || [];
      // minimal chart config
      const newChart = {
        'c:dLbls': {
          'c:showLegendKey': { $: { val: 0 } },
          'c:showVal': { $: { val: 0 } },
          'c:showCatName': { $: { val: 0 } },
          'c:showSerName': { $: { val: 0 } },
          'c:showPercent': { $: { val: 0 } },
          'c:showBubbleSize': { $: { val: 0 } },
          'c:showLeaderLines': { $: { val: 0 } },
        },
      };

      if (grouping != 'undefined') {
        newChart['c:grouping'] = {
          $: {
            val: grouping || CHART_GROUPING_BY_CHART_NAME[chart],
          },
        };
        if (grouping == 'stacked') {
          newChart['c:overlap'] = {
            $: {
              val: 100, // usually stacked expected to be seen with overlap 100%
            },
          };
        }
      }

      if (chart === 'column' || chart === 'bar') {
        // clone barChart from template
        // newChart = _.clone (me.chartTemplate ["c:chartSpace"]["c:chart"]["c:plotArea"][templateChartName]);
        newChart['c:barDir'] = {
          $: {
            val: chart.substr(0, 3),
          },
        };

      } else if (chart == 'line' || chart == 'area' || chart == 'radar' || chart == 'scatter') {
        // newChart = _.clone (me.chartTemplate ['c:chartSpace']['c:chart']['c:plotArea'][templateChartName]);
        // delete newChart['c:barDir'];
      } else {
        newChart['c:varyColors'] = {
          $: {
            val: 1,
          },
        };

        newChart['c:ser'] = ser;
        if (this.chartOptions.firstSliceAng) {
          newChart['c:firstSliceAng'] = {
            $: {
              val: this.chartOptions.firstSliceAng,
            },
          };
        }
        if (this.chartOptions.holeSize) {
          newChart['c:holeSize'] = {
            $: {
              val: this.chartOptions.holeSize,
            },
          };
        }

        if (this.chartOptions.showLabels) {
          newChart['c:dLbls']['c:showLegendKey'].$.val = 1;
        }
      }

      newChart['c:ser'] = ser;

      if (chart === 'column' || chart === 'bar' || chart === 'line' || chart === 'area' || chart === 'scatter') {
        if (!this.xAxis) {
          this.xAxis = {
            'c:delete': [{ $: { val: '0' } }], // delete axis, value not set or set to 1, means deleted
            'c:axId': {
              $: {
                val: Math.floor(Math.random() * 100000000),
              },
            },
            'c:scaling': {
              'c:orientation': { $: { val: 'minMax' } },
            },
            'c:axPos': { $: { val: 'b' } },
            'c:tickLblPos': { $: { val: 'nextTo' } },
            'c:crossAx': [
              // array is needed here to have multiple crossAx tags
              { $: { val: Math.floor(Math.random() * 100000000) } },
            ],
            'c:crosses': { $: { val: 'autoZero' } },
            'c:auto': { $: { val: '1' } },
            'c:lblAlgn': { $: { val: 'ctr' } },
            'c:lblOffset': { $: { val: 100 } },
            // 'c:numFmt': [{ $: { formatCode: 'General', sourceLinked: '0' } }],
            // 'c:majorTickMark': [{ $: { val: 'none' } }],
            // 'c:minorTickMark': [{ $: { val: 'none' } }],
          };

          this.content['c:chartSpace']['c:chart']['c:plotArea']['c:catAx'] = this.content['c:chartSpace']['c:chart']['c:plotArea']['c:catAx'] || [];
          this.content['c:chartSpace']['c:chart']['c:plotArea']['c:catAx'].push(this.xAxis);
        }

        if (!this.y1Axis) {
          this.y1Axis = {
            'c:delete': [{ $: { val: '0' } }], // delete axis, value not set or set to 1, means deleted
            'c:axId': { $: { val: this.xAxis['c:crossAx'][0]['$'].val } },
            'c:scaling': {
              'c:orientation': { $: { val: 'minMax' } },
            },
            'c:axPos': { $: { val: 'l' } },
            // 'c:majorGridlines': {},
            // 'c:minorGridlines': {},
            'c:numFmt': {
              $: {
                formatCode: 'General',
                sourceLinked: 1,
              },
            },
            'c:tickLblPos': { $: { val: 'nextTo' } },
            'c:crossAx': { $: { val: this.xAxis['c:axId']['$'].val } },
            'c:crosses': { $: { val: 'autoZero' } },
            'c:crossBetween': { $: { val: 'between' } },
          };

          this.content['c:chartSpace']['c:chart']['c:plotArea']['c:valAx'] = this.content['c:chartSpace']['c:chart']['c:plotArea']['c:valAx'] || [];
          this.content['c:chartSpace']['c:chart']['c:plotArea']['c:valAx'].push(this.y1Axis);
        }

        newChart['c:axId'] = [
          { $: { val: this.xAxis['c:axId']['$'].val } },
          { $: { val: this.y1Axis['c:axId']['$'].val } },
        ];
      }

      this.content['c:chartSpace']['c:chart']['c:plotArea'][chartTagName].push(newChart);

      if (this.chartOptions.legendPos === undefined || this.chartOptions.legendPos) {
        this.content['c:chartSpace']['c:chart']['c:legend'] = {
          'c:legendPos': {
            $: {
              val: this.legendPos || 'r',
            },
          },
        };
      } else if (this.chartOptions.legendPos === null) {
        delete this.content['c:chartSpace']['c:chart']['c:legend'];
      }

      if (this.chartOptions.manualLayout && this.chartOptions.manualLayout.plotArea) {
        this.content['c:chartSpace']['c:chart']['c:plotArea']['c:layout'] = this.content['c:chartSpace']['c:chart']['c:plotArea']['c:layout'] || {};
        this.content['c:chartSpace']['c:chart']['c:plotArea']['c:layout']['c:manualLayout'] = {
          'c:xMode': { $: { val: 'edge' } },
          'c:yMode': { $: { val: 'edge' } },
          'c:x': { $: { val: this.chartOptions.manualLayout.plotArea.x } },
          'c:y': { $: { val: this.chartOptions.manualLayout.plotArea.y } },
          'c:w': { $: { val: this.chartOptions.manualLayout.plotArea.w } },
          'c:h': { $: { val: this.chartOptions.manualLayout.plotArea.h } },
        };
      }
    });
  }

  getRelationship(relationshipId) {
    const position = Object.assign({
      fromColumn: 0,
      fromColumnOffset: 0,
      fromRow: this.id * 20,
      fromRowOffset: 0,
      toColumn: 10,
      toColumnOffset: 0,
      toRow: (this.id + 1) * 20,
      toRowOffset: 0,
    },
      this.chartOptions.position
    );

    return {
      'xdr:from': {
        'xdr:col': position.fromColumn,
        'xdr:colOff': position.fromColumnOffset,
        'xdr:row': position.fromRow,
        'xdr:rowOff': position.fromRowOffset,
      },
      'xdr:to': {
        'xdr:col': position.toColumn,
        'xdr:colOff': position.toColumnOffset,
        'xdr:row': position.toRow,
        'xdr:rowOff': position.toRowOffset,
      },
      'xdr:graphicFrame': {
        '$': {
          macro: '',
        },
        'xdr:nvGraphicFramePr': {
          'xdr:cNvPr': {
            $: {
              id: `${this.id + 1}`,
              name: `Diagram ${this.id}`,
            },
          },
          'xdr:cNvGraphicFramePr': {},
        },
        'xdr:xfrm': {
          'a:off': {
            $: {
              x: '0',
              y: '0',
            },
          },
          'a:ext': {
            $: {
              cx: '0',
              cy: '0',
            },
          },
        },
        'a:graphic': {
          'a:graphicData': {
            '$': {
              uri: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
            },
            'c:chart': {
              $: {
                'r:id': relationshipId,
                'xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'xmlns:r': this.relationType,
              },
            },
          },
        },
      },
      'xdr:clientData': {},
    };
  }
}

module.exports = Chart;
