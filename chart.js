const { chart } = require('./chart/ContentTypes');
const Chart = require('./chart/base');

const emptyCallBack = (...args) => args;

/**
 * @typedef {'noFill'} TransparentPointColor
 */

/**
 * @typedef OutputByType
 * @type {string | ArrayBuffer | Buffer | Uint8Array | Blob}
 */

/**
* @description Callback for chart generation.
* @callback OnReadyCallback
* @param {object?} error - Error object.
* @param {OutputByType} result - Base64 encoded string.
*/

/**
 * @description Name of one line in chart, or one columns group
 * @typedef {string} DataSeriesName
 */

/**
 * @description Commonly used for x-series
 * @typedef {string} DataPointName
 */

/**
 * @typedef {number} DataPointValue
 */

/**
 * @typedef {object} Position
 * @property {number} fromColumn - start column
 * @property {number} fromRow - start row
 * @property {number?} toColumn - end column
 * @property {number?} toRow - end row
 */

/**
 * @description Value of data point
 * @typedef {object} DataPoint
 * @property {DataPointValue} value - value of data point
 * @property {string | TransparentPointColor} [fillColor] - color of dot fill
 * @property {string | TransparentPointColor} [lineColor] - color of dot line
 * @property {string | TransparentPointColor} [markerColor] - color of dot marker
 */

/**
 * @description Data series of chart
 * @typedef ChartData
 * @type {object.<DataSeriesName, object.<DataPointName, DataPoint | DataPointValue>>}
 * @property {string?} [fillColor] - color of series fill
 * @property {string?} [lineColor] - color of series line
 * @property {string?} [markerColor] - color of series marker
 */

/**
 * @typedef {object} CustomColors - custom colors for chart
 * @property {object.<DataSeriesName, string | { fill: string, line: string, marker: string }>} series - custom colors for series
 * @property {object.<DataPointName, string | { fill: string, line: string, marker: string }>} points - custom colors for points
 */

/**
 * @typedef {object} ChartOptions
 * @property {string?} chartTitle - title of chart
 * @property {string} chart - type of chart
 * @property {ChartData} data - data for chart
 * @property {Position?} position - position of chart
 * @property {CustomColors?} customColors - custom colors for chart
 * @property {string[]?} titles - data series of chart
 * @property {string[]?} fields - data points of chart (commonly x-series)
 */


/**
 * @typedef {object} XLSXChartOptions
 * @property {string?} file - path to file for saving. Only to be called on server side.
 * @property {string?} [type="nodebuffer"] - encoding type for zip (xlsx) file
 * @property {Boolean?} dataPerSheet - if true, each chart will be placed on separate sheet
 * @property {ChartOptions[]?} charts - array of chart options
 */

const XLSXChart = {
  /**
   * @async
   * @param {XLSXChartOptions} options
   * @param {OnReadyCallback} callback
   * @returns {Promise.<OutputByType>}
   */
  generate(options, callback = emptyCallBack) {
    options = { type: 'nodebuffer', ...options }; // isolate passed options object, so local changes would not modify original object
    // the type is used by jszip library for correct zip generation and further for saving file

    const chart = new Chart();

    return chart.generate(options)
      .then((result) => {
        callback(null, result);

        return result;
      })
      .catch((error) => {
        callback(error, null);

        return Promise.reject(error);
      });
  },

  /**
   * @async
   * @param {XLSXChartOptions} options
   * @param {OnReadyCallback} callback
   * @returns {Promise.<OutputByType>}
   */
  writeFile(options, callback = emptyCallBack) {
    options = { type: 'base64', ...options }; // isolate passed options object, so local changes would not modify original object
    // the type is used by jszip library for correct zip generation and further for saving file

    const chart = new Chart();

    return chart.generate(options)
      .then((result) => {
        // this is a nodejs specific code, so it should be here
        const fs = require('fs');
        const { promisify } = require('util');

        return promisify(fs.writeFile)(options.file, result, 'base64');
      })
      .then((result) => {
        callback(null, result);

        return result;
      })
      .catch((error) => {
        callback(error, null);

        return Promise.reject(error);
      });
  },
};

module.exports = XLSXChart;





const newConfigExample = {
  file: 'path/to/file.xlsx',
  type: 'nodebuffer',
  dataPerSheet: true,
  charts: [
    {
      chartTitle: 'Chart 1',
      chart: 'line',
      properties: {
        lang: 'en-US',
        chartTile: {
          layout: {
            xMode: 'edge', // edge, factor, or edgeFactor
            yMode: 'edge', // edge, factor, or edgeFactor
            x: 0,
            y: 0,
          },
        },
        chartArea: {
          name: 'Diagram 1',
          twoCellAnchor: {
            from: {
              col: 0,
              colOff: 0,
              row: 0,
              rowOff: 0,
            },
            to: {
              col: 10,
              colOff: 0,
              row: 10,
              rowOff: 0,
            },
          },
          // transform2D: { // xdr:xfrm
          //   offset: {
          //     x: 0,
          //     y: 0,
          //   },
          //   ext: {
          //     cx: 0,
          //     cy: 0,
          //   },
          //   flipH: false,
          //   flipV: false,
          //   rotate: 0,
          // },
        },
        plotArea: {
          layout: {
            layoutTarget: 'inner',
            xMode: 'edge', // edge, factor, or edgeFactor
            yMode: 'edge', // edge, factor, or edgeFactor
            x: 0,
            y: 0,
            w: 1,
            h: 1,
          },
        },
        dataLayers: [
          {
            series: {
              'Series 1': {
                chartShape: {
                  smooth: false,
                  marker: {
                    // symbol: 'circle', // circle, dash, diamond, dot, none, picture, plus, square, star, triangle, x
                    // size: 5,
                    solidFill: 'FF0000',
                  },
                  line: {
                    solidFill: 'FF0000',
                    // width: 2,
                    // dashType: 'solid', // solid, dot, dash, lgDash, lgDashDot, lgDashDotDot, sysDash, sysDot
                  },
                  fill: {
                    solidFill: 'FF0000',
                  },
                },
                points: {
                  Jan: {
                    markers: {
                      symbol: 'circle', // circle, dash, diamond, dot, none, picture, plus, square, star, triangle, x
                      size: 5,
                    },
                    line: {
                      width: 2,
                      dashType: 'solid', // solid, dot, dash, lgDash, lgDashDot, lgDashDotDot, sysDash, sysDot
                    },
                    fill: {
                      solidFill: 'FF0000',
                    },
                  },
                },
              },
              'Series 2': {},
            },
            chart: 'line',
            grouping: 'clustered', // clustered, stacked, percentStacked
            axis: 'primary', // primary, secondary
            invertIfNegative: true,
            labels: {
              showLegendKey: false,
              showVal: true,
              showCatName: false,
              showSerName: false,
              showPercent: false,
              showBubbleSize: false,
            },
          },
        ],
        axes: {
          category: {
            scaling: {
              orientation: 'minMax', // minMax, min, max, or factor
              min: 0,
              max: 100,
              factor: 1,
            },
            position: 'b', // b, t, l, r, or nil
            crosses: 'autoZero', // autoZero, min, or max
            majorTickMark: 'cross', // cross, inside, outside, or none
            minorTickMark: 'cross', // cross, inside, outside, or none
            tickLblPos: 'nextTo', // nextTo, high, low, none
            numFmt: {
              formatCode: 'General',
              sourceLinked: true,
            },
            title: {
              layout: {
                xMode: 'edge', // edge, factor, or edgeFactor
                yMode: 'edge', // edge, factor, or edgeFactor
                x: 0,
                y: 0,
              },
            },
          },
          primary: {
            scaling: {
              orientation: 'minMax', // minMax, min, max, or factor
              min: 0,
              max: 100,
              factor: 1,
            },
            position: 'b', // b, t, l, r, or nil
            crosses: 'autoZero', // autoZero, min, or max
            majorTickMark: 'cross', // cross, inside, outside, or none
            minorTickMark: 'cross', // cross, inside, outside, or none
            tickLblPos: 'nextTo', // nextTo, high, low, none
            numFmt: {
              formatCode: 'General',
              sourceLinked: true,
            },
            title: {
              layout: {
                xMode: 'edge', // edge, factor, or edgeFactor
                yMode: 'edge', // edge, factor, or edgeFactor
                x: 0,
                y: 0,
              },
            },
          },
          secondary: {
            scaling: {
              orientation: 'minMax', // minMax, min, max, or factor
              min: 0,
              max: 100,
              factor: 1,
            },
            position: 'b', // b, t, l, r, or nil
            crosses: 'autoZero', // autoZero, min, or max
            majorTickMark: 'cross', // cross, inside, outside, or none
            minorTickMark: 'cross', // cross, inside, outside, or none
            tickLblPos: 'nextTo', // nextTo, high, low, none
            numFmt: {
              formatCode: 'General',
              sourceLinked: true,
            },
            title: {
              layout: {
                xMode: 'edge', // edge, factor, or edgeFactor
                yMode: 'edge', // edge, factor, or edgeFactor
                x: 0,
                y: 0,
              },
            },
          },
        },
        dataLabels: {},
        dataTable: {
          showHorzBorder: true,
          showVertBorder: true,
          showOutline: true,
          showKeys: true,
        },
        errorBars: {},
        gridLines: {},
        legend: {
          legendPos: 'r', // r, l, t, b, tr, tl, br, bl, or nil
          layout: {
            xMode: 'edge', // edge, factor, or edgeFactor
            yMode: 'edge', // edge, factor, or edgeFactor
            x: 0,
            y: 0,
            w: 1,
            h: 1,
          },
        },
        trendLines: [],
      },
      data: {
        'Series 1': {
          Jan: 10,
          Feb: 20,
          Mar: 30,
        },
        'Series 2': {
          Jan: 15,
          Feb: 25,
          Mar: 35,
        },
      },
      titles: ['Series 1', 'Series 2'],
      fields: ['Jan', 'Feb', 'Mar'],
    },
  ],
};
