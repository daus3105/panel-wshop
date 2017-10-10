var fs = require('fs'),
		Excel = require('exceljs'),
		filename = process.argv[2],
		dest = process.argv[3] || './output/panel-workshop.json',
		states = {
			Johor: {
				name: 'Johor',
				places: []
			},
			Kedah: {
				name: 'Kedah',
				places: []
			},
			Kelantan: {
				name: 'Kelantan',
				places: []
			},
			Melaka: {
				name: 'Melaka',
				places: []
			},
			NegeriSembilan: {
				name: 'Negeri Sembilan',
				places: []
			},
			Perak: {
				name: 'Perak',
				places: []
			},
			Pahang: {
				name: 'Pahang',
				places:  []
			},
			Perlis: {
				name: 'Perlis',
				places: []
			},
			PulauPinang: {
				name: 'Pulau Pinang',
				places: []
			},
			Sabah: {
				name: 'Sabah',
				places: []
			},
			Sarawak: {
				name: 'Sarawak',
				places: []
			},
			Selangor: {
				name: 'Selangor',
				places: []
			},
			Terengganu: {
				name: 'Terengganu',
				places: []
			},
			WilayahPersekutuan: {
				name: 'Wilayah Persekutuan',
				places: []
			},
		},
		dataMap = {
			name: 2,
			address: 3,
			lat: 4,
			long: 5,
			phone: 7,
			fax: 8,
			email: 9,
			state: 12
		};

var workbook = new Excel.Workbook();
workbook.xlsx.readFile(filename)
  .then(function(worksheet) {
    // use workbook 
    // console.log(worksheet);
    getSheet(1);
  });

function getSheet(id) {
	var worksheet = workbook.getWorksheet(id);
	worksheet.eachRow(function(row, rowNumber) {
		if (rowNumber === 1) return;
    var values = row.values;
    var state = values[dataMap.state];
    if (state === 'Negeri Sembilan')
    	state = 'NegeriSembilan';
   	else if (state === 'Pulau Pinang')
    	state = 'PulauPinang';
    else if (state === 'Wilayah Persekutuan')
    	state = 'WilayahPersekutuan';

    var email = values[dataMap.email];
    if (typeof email === 'object') {
    	if (email.text)
    		email = email.text;
    	else 
    		email.richText.map(function(a, b){
    			b ? email+= '/' + a.text : email = a.text.split(';')[0]; 
    		});
    }

    states[state]['places'].push({
    	name: values[dataMap.name],
    	address: values[dataMap.address],
    	phone: values[dataMap.phone] !== undefined ? values[dataMap.phone].toString() : '',
    	fax: values[dataMap.fax] !== undefined ? values[dataMap.fax].toString() : '',
    	email: email,
    	lat: values[dataMap.lat] !== undefined ? values[dataMap.lat].toString() : '',
    	lon: values[dataMap.long] !== undefined ? values[dataMap.long].toString() : ''
    });
    if ((rowNumber + 1) === worksheet.rowCount) {
			createJSON();
    }
	});
}

function createJSON() {
	var result = {states: []};
	for (var key in states) {
		if (states.hasOwnProperty(key)) {
			result.states.push(states[key]);
		}
	}
	// console.log(JSON.stringify(result));
	fs.writeFile(dest, JSON.stringify(result, null, 2), function(err) {
    if(err) {
        return console.log(err);
    }
    console.log("The file was saved!");
	}); 
}

// console.log(process.argv);
