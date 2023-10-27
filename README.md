node-xlsx2json
=========

simple XLSX reader to JSON

Install
-------
```
npm install
```

Dependency
----------
[node-xlsx](https://github.com/mgcrea/node-xlsx)

Usage
-----
```
var rows = xlsx2json.load({
	filename: filename,		/* mandatory: xlsx filename to read */
	sheetNumber: 0,			/* optional: 0 by default */
	minRow: 3,			/* optional: 0 by default */
	max: 250,  			/* optional: to limit max row, without no row limit */
	data: {				/* mandatory: object json mapping */
		lastname: 'H',	
		firstname: 'G',
		email: 'P'
	}
});
```

Looping on each row
-------------------
```
for(var i=0; i<rows.length; i++) {
	var row = rows[i];
	var fullname = row.lastname + ' ' + row.firstname + ' ' + row.email;
	console.log(fullname);
}
```
