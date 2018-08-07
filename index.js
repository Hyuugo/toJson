const express = require('express'),
	path = require('path'),
	XLSX = require('xlsx'),
	formidable = require('express-formidable'),
	PORT = process.env.PORT || 5000,
	app = express();

app.set('port', PORT);
app.set('views', __dirname + '/views');
app.set('view engine', 'pug');

app.use(function (req, res, next) {
	res.header("Access-Control-Allow-Origin", "*");
	res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
	next();
});
app.use(formidable());

function xlsx_to_json(wb, listName) {
	let result = [];

	if(!listName) {
		for (i = 0; i <= wb.SheetNames.length; i++) {
			result.push(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[i]], { header: 1 }));
		}
		return result;
	}
	
	let ws = wb.Sheets[wb.SheetNames[wb.SheetNames.indexOf(listName)]],
		range = ws["!ref"],
		xlsx_alphabet = [];
	result.push(XLSX.utils.sheet_to_json(ws, { header: 1 }));

	for (i = 65; i <= 90; i++)
		xlsx_alphabet.push(String.fromCharCode(i));

	for (i = 65; i <= 90; i++) {
		j = i;
		for (k = 65; k <= 90; k++)
			xlsx_alphabet.push(String.fromCharCode(j) + String.fromCharCode(k + 1));
		xlsx_alphabet.pop();
	}

	let first_header = range.slice(0, range.indexOf(':', 0)),
		second_header = range.slice(range.indexOf(':', 0) + 1, range.length),
		first_header_letter = "",
		second_header_letter = "";

	for (let i of first_header) {
		if (i.charCodeAt(0) >= 65 && i.charCodeAt(0) <= 90)
			first_header_letter += i;
	}

	for (let i of second_header) {
		if (i.charCodeAt(0) >= 65 && i.charCodeAt(0) <= 90)
			second_header_letter += i;
	}

	let arr = {};

	for (i = xlsx_alphabet.indexOf(first_header_letter), k = 0; i <= xlsx_alphabet.indexOf(second_header_letter); i++ , k++) {
		arr[xlsx_alphabet[i]] = []
		for (j = 0; j < result[0].length; j++) {
			if (result[0][j][k] != null)
				arr[xlsx_alphabet[i]].push(result[0][j][k]);
		}
		if (arr[xlsx_alphabet[i]].length == 0)
			delete arr[xlsx_alphabet[i]];
	}

	return arr;
}

app.get('/', function (req, res, next) {
	res.render('index');
});

app.post('/xlsx', function (req, res, next) {

	let f = req.files[Object.keys(req.files)[0]];

	if (f.size == 0 || f.name.length == 0)
		res.status(400).send("XLSX file required.");

	let wb = XLSX.readFile(f.path),
		listName = req.fields[Object.keys(req.fields)[0]];

	if (!listName)
		res.json(xlsx_to_json(wb));
	else
		res.json(xlsx_to_json(wb, listName));
});

app.listen(app.get('port'), function () {
	console.log('Node app is running on port', app.get('port'));
});