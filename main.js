var xlsx = require('node-xlsx');
var fs = require('fs');
var path = require('path');

var SRC_PATH = path.join(__dirname, 'src');
var DIST_PATH = path.join(__dirname, 'dist');
var DIST_FILENAME = path.join(DIST_PATH, 'build.xlsx');
var IGNORE_INDEXES = [0];

function loadXls(fileList) {
	fileList.forEach(function(filename, index) {
		console.log('读取文件: ' + filename);
		var obj = xlsx.parse(filename);
		if (!obj) {
			console.log('打开文件出错');
			return;
		}
		obj.forEach(function(sheet, i) {
			if (sheet.data && IGNORE_INDEXES.indexOf(i) == -1) {
				if (!result[i])
					result[i] = {
						'name': sheet.name,
						'data': []
					};
				sheet.data.forEach(function(item) {
					if (item && !isNaN(item[0])) {
						result[i].data.push(item);
					}
				});
			}
		});
	});
}

// fs.writeFile(__dirname + '/data.json', JSON.stringify(result), 'utf-8');

function saveResult(result) {
	console.log('正在合并...');
	var file = xlsx.build(result);
	fs.writeFileSync(path.join(DIST_PATH, 'result.xlsx'), file, 'binary');

	// result.forEach(function(sheet, index) {
	// 	if (!sheet)
	// 		return;
	// 	var file = xlsx.build([result[index]]);
	// 	console.log('正在合并' + result[index].name + '...');
	// 	fs.writeFile(path.join(DIST_PATH, result[index].name + '.xlsx'), file, 'binary', function(e) {
	// 		console.log(result[index].name + '合并完毕');
	// 	});
	// });
}

// Main process
if (!fs.existsSync(SRC_PATH)) {
	fs.mkdirSync(SRC_PATH);
}
if (fs.existsSync(DIST_PATH)) {
	var dstDir = fs.readdirSync(DIST_PATH);
	dstDir.forEach(function(item) {
		if (/^[^\.~]*.\.xlsx$/.test(item)) {
			fs.unlinkSync(path.join(DIST_PATH, item));
		}
	});
} else {
	fs.mkdirSync(DIST_PATH);
}

var srcDir = fs.readdirSync(SRC_PATH);
var fileList = [],
	result = [];
srcDir.forEach(function(item) {
	if (/^[^\.~]*.\.xlsx$/.test(item)) {
		fileList.push(path.join(SRC_PATH, item));
	}
});

console.log('正在处理，请稍后...');
loadXls(fileList);
saveResult(result);
console.log('合并完毕');


