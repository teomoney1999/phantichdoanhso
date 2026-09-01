const fs = require('fs');
let code = fs.readFileSync('src/components/Dashboard.jsx', 'utf8');
code = code.replace(
  'setSummaryData(summaryParsed);',
  'console.log("TESTING TRENDLINE", summaryParsed.slice(0, 2)); setSummaryData(summaryParsed);'
);
fs.writeFileSync('src/components/Dashboard.jsx', code);
