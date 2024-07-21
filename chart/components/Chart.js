const {
  getColName,
} = require('../services');
const CONTENT_TYPES = require('../ContentTypes');
const RELATION_TYPES = require('../relationTypes');
const {
  DEFAULT_TITLES_FIELD,
  DEFAULT_FIELDS_FIELD,
  DEFAULT_CHART_TYPE_FIELD,
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
    this.chartTypeField = this.chartOptions.chartTypeField || DEFAULT_CHART_TYPE_FIELD;
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
      'c:layout': this.getLayoutProperties(titlePosition),
      'c:overlay': { $: { val: 0 } },
    };
  }

  getLayoutProperties(position) {
    if (!position) {
      return null;
    }

    const layoutProperties = {
      'c:manualLayout': {},
    };

    ['xMode', 'yMode', 'x', 'y', 'w', 'h'].forEach((field) => {
      if (field in position) {
        layoutProperties['c:manualLayout'][`c:${field}`] = { $: { val: position[field] } };
      }
    });

    return layoutProperties;
  }

  getChartShapeProperties(chartShape) {
    const chartShapeProperties = {};

    if (chartShape.fill === 'noFill' && chartShape.line === 'noFill' && chartShape.marker === 'noFill') {
      chartShapeProperties['a:noFill'] = '';
      return chartShapeProperties;
    }

    if (chartShape.fill) {
      if (chartShape.fill.solidFill) {
        chartShapeProperties['a:solidFill'] = {
          'a:srgbClr': { $: { val: chartShape.fill.solidFill } },
        };
      }
    }

    if (chartShape.line) {
      if (chartShape.line.solidFill) {
        chartShapeProperties['a:ln'] = {
          'a:solidFill': {
            'a:srgbClr': { $: { val: chartShape.line.solidFill } },
          },
        };
      }
    }

    if (chartShape.marker) {
      if (chartShape.marker.solidFill) {
        chartShapeProperties['a:marker'] = {
          'c:spPr': {
            'a:solidFill': {
              'a:srgbClr': { $: { val: chartShape.marker.solidFill } },
            },
          },
        };
      }
    }

    return chartShapeProperties;
  }

  getAxisProperties(axisConfig) {
    if (!axisConfig) {
      return null;
    }

    const returnAxis = {};
    [
      'delete',
      'axId',
      'crossAx',
      'auto',
      'lblOffset',
    ].forEach((field) => {
      if (field in axisConfig) {
        returnAxis[`c:${field}`] = { $: { val: +axisConfig[field] } };
      }
    });

    [
      'axPos',
      'tickLblPos',
      'crosses',
      'crossBetween',
      // 'majorTickMark',
      // 'minorTickMark',
    ].forEach((field) => {
      if (field in axisConfig) {
        returnAxis[`c:${field}`] = { $: { val: axisConfig[field] } };
      }
    });

    if (axisConfig.scaling) {
      returnAxis['c:scaling'] = {};
      [
        'orientation',
        // 'min',
        // 'max',
        // 'logBase',
      ].forEach((field) => {
        returnAxis['c:scaling'][`c:${field}`] = { $: { val: axisConfig.scaling[field] } };
      });
    }

    if (axisConfig.numFmt) {
      returnAxis['c:numFmt'] = {
        $: {
          formatCode: axisConfig.numFmt.formatCode,
          sourceLinked: +axisConfig.numFmt.sourceLinked,
        },
      };
    }

    console.log('axisConfig', returnAxis);

    return returnAxis;
  }

  setChartData(chartConfig, rowOffset = 0, columnOffset = 0) {
    // Group series by chart type and grouping
    // it is needed to gather all series of the same type and grouping together as single chart layer

    console.log('dataLayers', chartConfig.properties.dataLayers);

    this.xAxis = this.getAxisProperties(chartConfig.properties.axes.category);
    this.y1Axis = this.getAxisProperties(chartConfig.properties.axes.primary);
    this.y2Axis = this.getAxisProperties(chartConfig.properties.axes.secondary);

    this.content['c:chartSpace']['c:chart']['c:plotArea']['c:catAx'] = [];
    this.content['c:chartSpace']['c:chart']['c:plotArea']['c:valAx'] = [];

    if (this.xAxis) {
      this.content['c:chartSpace']['c:chart']['c:plotArea']['c:catAx'].push(this.xAxis);
    }

    if (this.y1Axis) {
      this.content['c:chartSpace']['c:chart']['c:plotArea']['c:valAx'].push(this.y1Axis);
    }

    if (this.y2Axis) {
      this.content['c:chartSpace']['c:chart']['c:plotArea']['c:valAx'].push(this.y2Axis);
    }

    const chartTags = chartConfig.properties.dataLayers.reduce((chartTags, dataLayer) => {
      const chart = dataLayer.chart || this.chartOptions.chart;
      const grouping = dataLayer.grouping || this.chartOptions.grouping || CHART_GROUPING_BY_CHART_NAME[chart];
      // const varyColors = dataLayer.varyColors || this.chartOptions.varyColors || 1;

      const chartTag = {
        'c:grouping': { $: { val: grouping } },
        // 'c:varyColors': { $: { val: +varyColors } },
        'c:dLbls': {
          'c:showLegendKey': { $: { val: 0 } },
          'c:showVal': { $: { val: 0 } },
          'c:showCatName': { $: { val: 0 } },
          'c:showSerName': { $: { val: 0 } },
          'c:showPercent': { $: { val: 0 } },
          'c:showBubbleSize': { $: { val: 0 } },
          'c:showLeaderLines': { $: { val: 0 } },
        },
        'c:axId': [
          { $: { val: this.xAxis?.['c:axId'].$.val || 0 } },
          { $: { val: this.y1Axis?.['c:axId'].$.val || 0 } },
        ],
      };

      if (grouping === 'stacked' || grouping === 'percentStacked') {
        chartTag['c:overlap'] = { $: { val: 100 } }; // usually stacked expected to be seen with overlap 100%
      }

      switch (chart) {
        case 'column':
        case 'bar':
          chartTag['c:barDir'] = { $: { val: chart.substring(0, 3) } };
          break;
        case 'line':
        case 'area':
        case 'radar':
        case 'scatter':
          chartTag['c:marker'] = { $: { val: 1 } }; // TODO: add marker support
          break;
        default:
        // do nothing
      }

      [
        'firstSliceAng',
        'holeSize',
        // 'invertIfNegative',
      ].forEach((field) => {
        if (field in dataLayer) {
          chartTag[`c:${field}`] = { $: { val: +dataLayer[field] } };
        }
      });

      if (dataLayer.labels) {
        chartTag['c:dLbls'] = Object.entries(dataLayer.labels).reduce((dLbls, [key, value]) => {
          dLbls[`c:${key}`] = { $: { val: +value } };
          return dLbls;
        }, chartTag['c:dLbls']);
      }

      console.log('dataLayer', dataLayer);

      chartTag['c:ser'] = Object.entries(dataLayer.series).map(([title, seriesData], columnIndex) => {
        const customColorsSeries = {};
        const customColorsPoints = {};

        if (seriesData.chartShape) {
          customColorsSeries['c:spPr'] = this.getChartShapeProperties(seriesData.chartShape);
        }

        if (seriesData.points) {
          customColorsPoints['c:dPt'] = Object.entries(seriesData.points).map(([field, point]) => {
            const fieldIndex = chartConfig.fields.indexOf(field);

            const dataPoint = {
              'c:idx': { $: { val: fieldIndex } },
            };

            ['bubble3D', 'explosion', 'invertIfNegative'].forEach((field) => {
              if (field in point) {
                dataPoint[`c:${field}`] = { $: { val: +point[field] } };
              }
            });

            if (point.chartShape) {
              dataPoint['c:spPr'] = this.getChartShapeProperties(point.chartShape);
            }

            return dataPoint;
          }).filter(Boolean);
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
                  'c:v': title,
                },
              },
            },
          },
          'c:cat': {
            'c:strRef': {
              'c:f': `'${this.dataWorksheetName}'!$${getColName(1 + columnOffset)}$${rowOffset + 2}:$${getColName(1 + columnOffset)}$${chartConfig[this.fieldsField].length + rowOffset + 1}`,
              'c:strCache': {
                'c:ptCount': { $: { val: chartConfig.fields.length } },
                'c:pt': chartConfig.fields.map((field, rowIndex) => {
                  return {
                    '$': {
                      idx: rowIndex,
                    },
                    'c:v': field,
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
                'c:ptCount': { $: { val: chartConfig.fields.length } },
                'c:pt': chartConfig.fields.map((field, rowIndex) => {
                  return {
                    '$': {
                      idx: rowIndex,
                    },
                    'c:v': chartConfig.data[title]?.[field] ?? 0,
                  };
                }),
              },
            },
          },
          // ...customColorsSeries,
          // ...customColorsPoints,
        };

        return ser;
      });

      const chartTagName = CHART_TAG_BY_CHART_NAME[chart];

      if (!chartTagName) {
        throw new Error(`Chart type '${chart}' is not supported`);
      }

      return {
        ...chartTags,
        [chartTagName]: chartTag,
      };
    }, {});

    this.content['c:chartSpace']['c:chart']['c:plotArea'] = {
      'c:layout': this.getLayoutProperties(chartConfig.properties.plotArea),
      ...this.content['c:chartSpace']['c:chart']['c:plotArea'],
      ...chartTags,
    };

    delete this.content['c:chartSpace']['c:chart']['c:legend'];


    // if (this.chartOptions.legendPos === undefined || this.chartOptions.legendPos) {
    //   this.content['c:chartSpace']['c:chart']['c:legend'] = {
    //     'c:legendPos': {
    //       $: {
    //         val: this.legendPos || 'r',
    //       },
    //     },
    //   };
    // } else if (this.chartOptions.legendPos === null) {
    //   delete this.content['c:chartSpace']['c:chart']['c:legend'];
    // }

    // if (this.chartOptions.manualLayout && this.chartOptions.manualLayout.plotArea) {
    //   this.content['c:chartSpace']['c:chart']['c:plotArea']['c:layout'] = this.content['c:chartSpace']['c:chart']['c:plotArea']['c:layout'] || {};
    //   this.content['c:chartSpace']['c:chart']['c:plotArea']['c:layout']['c:manualLayout'] = {
    //     'c:xMode': { $: { val: 'edge' } },
    //     'c:yMode': { $: { val: 'edge' } },
    //     'c:x': { $: { val: this.chartOptions.manualLayout.plotArea.x } },
    //     'c:y': { $: { val: this.chartOptions.manualLayout.plotArea.y } },
    //     'c:w': { $: { val: this.chartOptions.manualLayout.plotArea.w } },
    //     'c:h': { $: { val: this.chartOptions.manualLayout.plotArea.h } },
    //   };
    // }
  }

  getRelationship(relationshipId) {
    const position = this.chartOptions.properties.chartArea.twoCellAnchor;
    const name = this.chartOptions.properties.chartArea.name;

    return {
      'xdr:from': {
        'xdr:col': position.from.col,
        'xdr:colOff': position.from.colOff,
        'xdr:row': position.from.row,
        'xdr:rowOff': position.from.rowOff,
      },
      'xdr:to': {
        'xdr:col': position.to.col,
        'xdr:colOff': position.to.colOff,
        'xdr:row': position.to.row,
        'xdr:rowOff': position.to.rowOff,
      },
      'xdr:graphicFrame': {
        '$': {
          macro: '',
        },
        'xdr:nvGraphicFramePr': {
          'xdr:cNvPr': {
            $: {
              id: `${this.id + 1}`,
              name,
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
