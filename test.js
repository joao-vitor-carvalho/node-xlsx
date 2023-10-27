var xlsx2json = require('../index.js');

var rows = xlsx2json.load({
	filename: './test/countries.xlsx',
	minRow: 6,          // optional: 0 by default
	data: {             // mandatory: object json mapping
		country: 'A',
		region: 'B'
	}
});

exports.testParsing = function(test) {
	test.ok(true, rows[10].country.indexOf('Australia') === 0);
	test.done();
};


// cria o excel em Create-excel.js e converte em JSON a partir deste arquivo test.js