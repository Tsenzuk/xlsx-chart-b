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
 * @type {object.<DataSeriesName, object.<DataPointName, DataPoint>>}
 */

/**
 * @typedef {object} CustomColors - custom colors for chart
 * @property {object.<DataSeriesName, string | { fill: string, line: string, marker: string }>} series - custom colors for series
 * @property {object.<DataPointName, string | { fill: string, line: string, marker: string }>} points - custom colors for points
 */

/**
 * @typedef {object} ChartOptions
 * @property {string} chartTitle - title of chart
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
 * @property {chartOptions[]?} charts - array of chart options
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
