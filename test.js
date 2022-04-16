const chart = require('./chart');

// const options = {
//   file: './test.xlsx',
// };

const options1 = {
  file: './test.xlsx',
  charts: [
    {
      position: {
        fromColumn: 10,
        toColumn: 20,
        fromRow: 10,
        toRow: 20,
      },
      customColors: {
        points: {
          'Title 1': {
            'Field 1': 'FF0000',
            'Field 2': '00FF00',
          },
        },
        series: {
          'Title 2': {
            fill: 'aaaaaa',
            line: '00ffff',
          },
        },
      },
      chart: 'bar',
      data: {
        'Title 1': {
          'Field 1': 5,
          'Field 2': 10,
          'Field 3': 15,
          'Field 4': 20,
        },
        'Title 2': {
          'Field 1': 10,
          'Field 2': 5,
          'Field 3': 20,
          'Field 4': 15,
        },
        'Title 3': {
          'Field 1': 20,
          'Field 2': 15,
          'Field 3': 10,
          'Field 4': 5,
        },
      },
      chartTitle: 'Title 1',
    },
  ],
};

chart.writeFile(options1)
  .then(console.log, console.error);
