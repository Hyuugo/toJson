const express = require('express'),
    path = require('path'),
    XLSX = require('xlsx'),
    formidable = require('express-formidable'),
    URL = require('url');
    PORT = process.env.PORT || 5000,
	app = express();

app.use(formidable());
app.set('port', PORT);
app.set('views', __dirname + '/views');
app.set('view engine', 'pug');

app.use(function(req, res, next) {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
    next();
  });

function make_book() {
	var ws = XLSX.utils.aoa_to_sheet(data);
	var wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
	return wb;
}

function get_data(req, res, type) {
	var wb = make_book();
	res.status(200).send(XLSX.write(wb, {type:'buffer', bookType:type}));
}

function get_file(req, res, file) {
	var wb = make_book();
	XLSX.writeFile(wb, file);
	res.status(200).send("wrote to " + file + "\n");
}

function load_data(file, sheetName) {
	let wb = XLSX.readFile(file),
		result = [],
		ws = wb.Sheets[wb.SheetNames[wb.SheetNames.indexOf(sheetName)]];
        result.push(XLSX.utils.sheet_to_json(ws, {header:1})),
        range = ws["!ref"],
        xlsx_alphabet = [];

	for (i = 65; i <= 90; i++)
		xlsx_alphabet.push(String.fromCharCode(i));

	for (i = 65; i <= 90; i++) {
		j = i;
		for (k = 65; k <= 90; k++)
			xlsx_alphabet.push(String.fromCharCode(j) + String.fromCharCode(k+1));
		xlsx_alphabet.pop();
	}

	let first_header = range.slice(0, range.indexOf(':', 0)),
        second_header = range.slice(range.indexOf(':', 0) + 1, range.length),
        first_header_letter = "",
        second_header_letter = "";

	for (let i of first_header) {
		if(i.charCodeAt(0) >= 65 && i.charCodeAt(0) <= 90)
			first_header_letter += i;
	}

	for (let i of second_header) {
		if(i.charCodeAt(0) >= 65 && i.charCodeAt(0) <= 90)
			second_header_letter += i;
	}

	let arr = {};

	for (i = xlsx_alphabet.indexOf(first_header_letter), k = 0; i <= xlsx_alphabet.indexOf(second_header_letter); i++, k++) {
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

function post_data(req, res) {
	var keys = Object.keys(req.files), 
		k = keys[0];
	res.json(load_data(req.files[k].path, req.fields.listName));
}

function post_file(req, res, file) {
	res.json(load_data(file));
}

app.get('/', function (req, res, next) {
	res.render('toJson');
});

app.post('/xlsx', function (req, res, next) {
	var url = URL.parse(req.url, true);
	if(url.query.f && url.query.listName) 
		return post_file(req, res, url.query.f, url.query.listName);
	return post_data(req, res);
});


app.listen(app.get('port'), function () {
	console.log('Node app is running on port', app.get('port'));
});