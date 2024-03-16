const DEFAULT_EXPORT_TYPE = 'nodebuffer';
const DEFAULT_CHART_TYPE = 'column';
const DEFAULT_TITLES_FIELD = 'titles';
const DEFAULT_FIELDS_FIELD = 'fields';
const DEFAULT_CHART_TYPE_FIELD = 'chart';
const DEFAULT_LINE_COLOR_FIELD = 'lineColor';
const DEFAULT_FILL_COLOR_FIELD = 'fillColor';
const DEFAULT_MARKER_COLOR_FIELD = 'markerColor';

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

  const titlesField = chartOptions.titlesField || DEFAULT_TITLES_FIELD;
  const fieldsField = chartOptions.fieldsField || DEFAULT_FIELDS_FIELD;
  const chartTypeField = chartOptions.chartTypeField || DEFAULT_CHART_TYPE_FIELD;
  const lineColorField = chartOptions.lineColorField || DEFAULT_LINE_COLOR_FIELD;
  const fillColorField = chartOptions.fillColorField || DEFAULT_FILL_COLOR_FIELD;
  const markerColorField = chartOptions.markerColorField || DEFAULT_MARKER_COLOR_FIELD;

  const titles = chartOptions[titlesField] || [];
  const fields = chartOptions[fieldsField] || [];
  const data = {};

  // TODO: make the code asynchronous to avoid blocking if there are many charts or data points
  for (const title in chartOptions.data) {
    if (chartOptions[titlesField] && !chartOptions[titlesField].includes(title)) {
      continue; // if user provided titles, skip the ones that are not in the list
    }

    if (!titles.includes(title)) {
      titles.push(title);
    }

    data[title] = {};

    for (const field in chartOptions.data[title]) {
      if (chartOptions[fieldsField] && !chartOptions[fieldsField].includes(field)) {
        continue; // if user provided fields, skip the ones that are not in the list
      }

      if (!fields.includes(field)) {
        fields.push(field);
      }

      const point = chartOptions.data[title][field];

      if (typeof point === 'object') {
        data[title][field] = point;
      } else {
        const colorConfig = chartOptions.customColors?.points?.[title]?.[field] || null;
        let fillColor = null;
        let lineColor = null;
        let markerColor = null;

        if (typeof colorConfig === 'string') {
          fillColor = colorConfig;
          lineColor = colorConfig;
          markerColor = colorConfig;
        } else if (colorConfig && typeof colorConfig === 'object') {
          fillColor = colorConfig.fill || fillColor;
          lineColor = colorConfig.line || lineColor;
          markerColor = colorConfig.marker || markerColor;
        }

        data[title][field] = {
          value: point,
          fillColor,
          lineColor,
          markerColor,
        };
      }
    }

    if (chartOptions[titlesField] && chartOptions.data[title][chartTypeField]) {
      data[title][chartTypeField] = chartOptions.data[title][chartTypeField];
    }

    if (chartOptions.customColors?.series?.[title]) {
      const colorConfig = chartOptions.customColors.series[title];

      if (!colorConfig) {
        continue;
      }

      if (typeof colorConfig === 'string') {
        data[title][fillColorField] = colorConfig;
        data[title][lineColorField] = colorConfig;
        data[title][markerColorField] = colorConfig;
        continue;
      }

      data[title][fillColorField] = colorConfig.fill || data[title][fillColorField];
      data[title][lineColorField] = colorConfig.line || data[title][lineColorField];
      data[title][markerColorField] = colorConfig.marker || data[title][markerColorField];
    }
  }

  returnOptions.data = data;
  returnOptions[titlesField] = titles;
  returnOptions[fieldsField] = fields;
  returnOptions.position = chartOptions.position;
  returnOptions.majorGridlines = chartOptions.majorGridlines ?? true;

  // service options, used for
  returnOptions.titlesField = titlesField;
  returnOptions.fieldsField = fieldsField;
  returnOptions.lineColorField = lineColorField;
  returnOptions.fillColorField = fillColorField;
  returnOptions.markerColorField = markerColorField;

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
    returnOptions.charts.push(normalizeChartOptions({
      data: this.options.data,
      chartTitle: this.options.chartTitle,
      chart: this.options.chart,
    }, 0));
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
  DEFAULT_TITLES_FIELD,
  DEFAULT_FIELDS_FIELD,
  DEFAULT_CHART_TYPE_FIELD,
  DEFAULT_LINE_COLOR_FIELD,
  DEFAULT_FILL_COLOR_FIELD,
  DEFAULT_MARKER_COLOR_FIELD,
};
