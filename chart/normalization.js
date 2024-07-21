const DEFAULT_EXPORT_TYPE = 'nodebuffer';
const DEFAULT_CHART_TYPE = 'column';
const DEFAULT_CHART_LANG = 'en-US';
const DEFAULT_CHART_TITLE = 'Chart';
const DEFAULT_LOCATION_X_MODE = 'edge';
const DEFAULT_LOCATION_Y_MODE = 'edge';
const DEFAULT_DIAGRAM_NAME = 'Diagram';
const DEFAULT_AXIS_ORIENTATION = 'minMax';
const DEFAULT_NUMBER_FORMAT = 'General';
const DEFAULT_TICK_LABEL_POSITION = 'nextTo';
const DEFAULT_TICK_MARK = 'cross';
const DEFAULT_LABEL_ALIGN = 'ctr';
const DEFAULT_LEGEND_POSITION = 'r';

const DEFAULT_TITLES_FIELD = 'titles';
const DEFAULT_FIELDS_FIELD = 'fields';
const DEFAULT_CHART_TYPE_FIELD = 'chart';
const DEFAULT_LINE_COLOR_FIELD = 'lineColor';
const DEFAULT_FILL_COLOR_FIELD = 'fillColor';
const DEFAULT_MARKER_COLOR_FIELD = 'markerColor';

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

/**
 * @typedef {import('../chart').XLSXChartOptions} XLSXChartOptions
 * @typedef {import('../chart').ChartOptions} ChartOptions
 */

const normalizeChartOptionsLocation = (location = {}) => {
  const returnLocation = {};

  returnLocation.x = location.x || 0;
  returnLocation.y = location.y || 0;
  returnLocation.xMode = location.xMode || DEFAULT_LOCATION_X_MODE;
  returnLocation.yMode = location.yMode || DEFAULT_LOCATION_Y_MODE;

  return returnLocation;
};

const normalizeChartOptionsTwoCellAnchor = (twoCellAnchor = {}, chartIndex) => {
  const returnTwoCellAnchor = {};

  returnTwoCellAnchor.from = twoCellAnchor.from || {};
  returnTwoCellAnchor.to = twoCellAnchor.to || {};

  returnTwoCellAnchor.from.col = returnTwoCellAnchor.from.col ?? 0;
  returnTwoCellAnchor.from.colOff = returnTwoCellAnchor.from.colOff ?? 0;
  returnTwoCellAnchor.from.row = returnTwoCellAnchor.from.row ?? 20 * chartIndex;
  returnTwoCellAnchor.from.rowOff = returnTwoCellAnchor.from.rowOff ?? 0;

  returnTwoCellAnchor.to.col = returnTwoCellAnchor.to.col ?? 10;
  returnTwoCellAnchor.to.colOff = returnTwoCellAnchor.to.colOff ?? 0;
  returnTwoCellAnchor.to.row = returnTwoCellAnchor.to.row ?? (20 * chartIndex + 20);
  returnTwoCellAnchor.to.rowOff = returnTwoCellAnchor.to.rowOff ?? 0;

  return returnTwoCellAnchor;
};

/**
 *
 * @link https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.categoryaxis?view=openxml-3.0.1
 * @link https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.valueaxis?view=openxml-3.0.1
 * @param {*} axis
 * @param {*} axisType
 * @param {*} chartType
 * @returns
 */
const normalizeChartOptionsAxis = (axis, axisType, chartType) => {
  if (axisType === 'secondary' && !axis) {
    return null;
  }

  axis = axis || {};

  const returnAxis = {
    axId: axis.axId || Math.floor(Math.random() * 100000000),
  };


  // position https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.axispositionvalues?view=openxml-3.0.1
  returnAxis.axPos = axis.axPos;
  if (chartType === 'bar') {
    const defaultAxisPosition = {
      category: 'l',
      primary: 'b',
      secondary: 't',
    };
    returnAxis.axPos = axis.axPos || defaultAxisPosition[axisType];
  } else {
    const defaultAxisPosition = {
      category: 'b',
      primary: 'l',
      secondary: 'r',
    };
    returnAxis.axPos = axis.axPos || defaultAxisPosition[axisType];
  }

  returnAxis.delete = axis.delete || false;
  returnAxis.crosses = axis.crosses || 'autoZero';
  returnAxis.auto = axis.auto || true;
  returnAxis.lblAlgn = axis.lblAlgn || 'ctr';

  // majorGridlines https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.chartlinestype?view=openxml-3.0.1
  returnAxis.majorGridlines = axis.majorGridlines || null;

  // majorTickMark https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.tickmarkvalues?view=openxml-3.0.1
  returnAxis.majorTickMark = axis.majorTickMark || DEFAULT_TICK_MARK;

  // minorGridlines https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.chartlinestype?view=openxml-3.0.1
  returnAxis.minorGridlines = axis.minorGridlines || null;

  // minorTickMark https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.tickmarkvalues?view=openxml-3.0.1
  returnAxis.minorTickMark = axis.minorTickMark || DEFAULT_TICK_MARK;

  // numberingFormat
  if (axisType !== 'category') {
    returnAxis.numFmt = axis.numFmt || {};
    returnAxis.numFmt.formatCode = returnAxis.numFmt.formatCode || 'General';
    returnAxis.numFmt.sourceLinked = returnAxis.numFmt.sourceLinked || true;
  }

  // scaling
  returnAxis.scaling = axis.scaling || {};
  returnAxis.scaling.min = returnAxis.scaling.min || null;
  returnAxis.scaling.max = returnAxis.scaling.max || null;
  returnAxis.scaling.orientation = returnAxis.scaling.orientation || DEFAULT_AXIS_ORIENTATION;
  returnAxis.scaling.logBase = returnAxis.scaling.logBase || null;

  // textProperties
  returnAxis.textProperties = axis.textProperties || null;

  // tickLabelPosition https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.ticklabelpositionvalues?view=openxml-3.0.1
  returnAxis.tickLblPos = axis.tickLblPos || DEFAULT_TICK_LABEL_POSITION;
  returnAxis.lblAlgn = axis.lblAlgn ||DEFAULT_LABEL_ALIGN;
  // returnAxis.lblOffset = axis.lblOffset || 100;
  returnAxis.noMultiLvlLbl = axis.noMultiLvlLbl || true;

  // title
  returnAxis.title = axis.title || {};
  returnAxis.title.layout = normalizeChartOptionsLocation(returnAxis.title.layout);

  return returnAxis;
};

/**
 *
 * @param {ChartOptions} chartOptions
 * @returns {ChartOptions}
 */
const normalizeChartOptions = (chartOptions, chartIndex) => {
  const returnOptions = {};

  returnOptions.chartTitle = chartOptions.chartTitle || `${DEFAULT_CHART_TITLE} ${chartIndex}`;
  returnOptions.chart = chartOptions.chart || DEFAULT_CHART_TYPE;

  returnOptions.titles = chartOptions.titles || [];
  returnOptions.fields = chartOptions.fields || [];

  returnOptions.data = chartOptions.data || {};

  if (!returnOptions.titles.length) {
    returnOptions.titles = Object.keys(returnOptions.data);
  }
  if (!returnOptions.fields.length) {
    returnOptions.fields = Array.from(new Set(returnOptions.titles.map((title) => Object.keys(returnOptions.data[title])).flat()));
  }

  const properties = chartOptions.properties || {};

  // properties
  returnOptions.properties = properties || {};

  // lang
  returnOptions.properties.lang = properties.lang || DEFAULT_CHART_LANG;

  // chartTitle
  returnOptions.properties.chartTitle = properties.chartTitle || {};
  returnOptions.properties.chartTitle.layout = normalizeChartOptionsLocation(properties.chartTitle.layout);

  // chartArea
  returnOptions.properties.chartArea = properties.chartArea || {};
  returnOptions.properties.chartArea.name = returnOptions.properties.chartArea.name || `${DEFAULT_DIAGRAM_NAME} ${chartIndex}`;
  let twoCellAnchor = returnOptions.properties.chartArea.twoCellAnchor || {};
  if (chartOptions.position) {
    console.warn('@deprecated chartOptions.position is deprecated, use properties.chartArea.twoCellAnchor instead');
    twoCellAnchor = {
      from: {
        col: chartOptions.position.fromColumn,
        colOff: chartOptions.position.fromColumnOffset,
        row: chartOptions.position.fromRow,
        rowOff: chartOptions.position.fromRowOffset,
      },
      to: {
        col: chartOptions.position.toColumn,
        colOff: chartOptions.position.toColumnOffset,
        row: chartOptions.position.toRow,
        rowOff: chartOptions.position.toRowOffset,
      },
    };
  }
  returnOptions.properties.chartArea.twoCellAnchor = normalizeChartOptionsTwoCellAnchor(twoCellAnchor, chartIndex);
  returnOptions.properties.chartArea.transform2D = properties.chartArea.transform2D || {};

  // plotArea
  returnOptions.properties.plotArea = properties.plotArea || {};
  returnOptions.properties.plotArea.layout = normalizeChartOptionsLocation(properties.plotArea.layout);

  // axes
  returnOptions.properties.axes = properties.axes || {};
  returnOptions.properties.axes.category = properties.axes.category || normalizeChartOptionsAxis(properties.axes.category, 'category', returnOptions.chart);
  returnOptions.properties.axes.primary = properties.axes.primary || normalizeChartOptionsAxis(properties.axes.primary, 'primary', returnOptions.chart);
  returnOptions.properties.axes.secondary = properties.axes.secondary || normalizeChartOptionsAxis(properties.axes.secondary, 'secondary', returnOptions.chart);
  returnOptions.properties.axes.category.crossAx = returnOptions.properties.axes.category.crossAx || [returnOptions.properties.axes.primary.axId];
  returnOptions.properties.axes.primary.crossAx = returnOptions.properties.axes.primary.crossAx || [returnOptions.properties.axes.category.axId];

  // legend https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.legendpositionvalues?view=openxml-3.0.1
  returnOptions.properties.legend = properties.legend || {};
  returnOptions.properties.legend.position = properties.legend.position || DEFAULT_LEGEND_POSITION;
  returnOptions.properties.legend.layout = normalizeChartOptionsLocation(properties.legend.layout);

  // dataLayers
  returnOptions.properties.dataLayers = properties.dataLayers;
  if (!returnOptions.properties.dataLayers) {
    returnOptions.properties.dataLayers = [];
    const dataLayer = {
      chart: returnOptions.chart,
      grouping: CHART_GROUPING_BY_CHART_NAME[returnOptions.chart],
      axis: 'primary',
      invertIfNegative: true,
      labels: {
        showLegendKey: false,
        showVal: false,
        showCatName: false,
        showSerName: false,
        showPercent: false,
        showBubbleSize: false,
      },
      series: returnOptions.titles.reduce((acc, title) => {
        acc[title] = {
          points: returnOptions.fields.reduce((acc, field) => {
            acc[field] = {
              value: returnOptions.data[title][field],
            };
            return acc;
          }, {}),
        };
        return acc;
      }, {}),
    };
    returnOptions.properties.dataLayers.push(dataLayer);

    if (chartOptions.customColors) {
      console.log('@deprecated chartOptions.customColors is deprecated, use properties.dataLayers[].series instead');
      if (chartOptions.customColors.series) {
        dataLayer.series = {};
        Object.keys(chartOptions.customColors.series).forEach((title) => {
          dataLayer.series[title] = {
            chartShape: {},
          };

          if (chartOptions.customColors.series[title].fill) {
            dataLayer.series[title].chartShape.fill = {
              solidFill: chartOptions.customColors.series[title].fill,
            };
          }
          if (chartOptions.customColors.series[title].line) {
            dataLayer.series[title].chartShape.line = {
              solidFill: chartOptions.customColors.series[title].line,
            };
          }
          if (chartOptions.customColors.series[title].marker) {
            dataLayer.series[title].chartShape.marker = {
              solidFill: chartOptions.customColors.series[title].marker,
            };
          }
        });
      }
    }
  } else {
    console.warn('data chartShape support TBD');
  }

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
