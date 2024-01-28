const DEFAULT_EXPORT_TYPE = 'nodebuffer';
const DEFAULT_CHART_TYPE = 'column';

/**
 * @typedef {import('../chart').XLSXChartOptions} XLSXChartOptions
 * @typedef {import('../chart').ChartOptions} ChartOptions
 */

/**
 *
 * @param {ChartOptions} chartOptions
 * @returns {ChartOptions}
 */
const normalizeChartOptions = (chartOptions, chartIndex) => {
  const returnOptions = {};

  returnOptions.chartTitle = chartOptions.chartTitle || `Chart ${chartIndex}`;
  returnOptions.chart = chartOptions.chart || DEFAULT_CHART_TYPE;

  const titles = chartOptions.titles || [];
  const fields = chartOptions.fields || [];
  const data = {};

  // TODO: make the code asynchronous to avoid blocking if there are many charts or data points
  for (const title in chartOptions.data) {
    if (!titles.includes(title)) {
      titles.push(title);
    }

    data[title] = {};

    for (const field in chartOptions.data[title]) {
      if (!fields.includes(field)) {
        fields.push(field);
      }

      const point = chartOptions.data[title][field];

      if (typeof point === 'object') {
        data[title][field] = point;
      } else {
        data[title][field] = { value: point };
      }
    }
  }

  returnOptions.data = data;
  returnOptions.titles = titles;
  returnOptions.fields = fields;

  return returnOptions;
};

/**
 *
 * @param {XLSXChartOptions} options
 * @returns
 */
const normalizeOptions = (options) => {
  if (typeof options !== 'object') {
    return options;
  }

  const returnOptions = { charts: [] };

  returnOptions.type = options.type || DEFAULT_EXPORT_TYPE;

  // single-chart config
  if (options.data) {
    console.warn('@deprecated Single chart config using options.data is deprecated, use options.charts[] instead');
    returnOptions.charts.push({
      data: this.options.data,
      chartTitle: this.options.chartTitle,
      chart: this.options.chart,
    });
  }

  if (options.charts && options.charts.length) {
    returnOptions.charts = returnOptions.charts.concat(...options.charts);
  }

  returnOptions.charts = returnOptions.charts.map(normalizeChartOptions);

  returnOptions.dataPerSheet = options.dataPerSheet || false;

  return returnOptions;
};

const validateOptions = (options) => {
  const errors = [];

  if (typeof options !== 'object') {
    return errors.push(new Error(`options should be correct object, passed ${typeof options}`));
  }
  const {
    type,
    charts,
  } = options;


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
};

module.exports = {
  normalizeOptions,
  validateOptions,
};
