const Chart = require('./chart/base');

const emptyCallBack = (...args) => args;

/**
 * @typedef options
 * @property {string} file
 */

const XLSXChart = {
  /**
   *
   * @async
   * @param {{}} options
   * @param {function} callback
   * @returns {Promise.<*>}
   */
  generate(options, callback = emptyCallBack) {
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
   *
   * @async
   * @param {options} options
   * @param {function} callback
   * @returns {Promise.<*>}
   */
  writeFile(options, callback = emptyCallBack) {
    options = { type: 'base64', ...options }; // isolate passed options object, so local changes would not modify original object
    // the type is used by jszip library for correct zip generation and further for saving file

    return this.generate(options)
      .then((result) => {
        const fs = require ('fs');
        const { promisify } = require('util');

        return promisify(fs.writeFile)(options.file, result, 'base64');
      })
      .then((result) => callback(null, result))
      .catch((error) => callback(error, null));
  },
};

module.exports = XLSXChart;
