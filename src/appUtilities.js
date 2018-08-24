/* global Word:true, OfficeExtension:true */

'use strict';

export function errHandler(error) {
	console.log("Error: " + error);

	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
}

export function createRexFromString(string, flags) {
	var escapedString = string.replace(/[|\\{}()[\]^$+*?.]/g, '\\$&');
	return new RegExp(escapedString, flags);
}

export function sortByAlphabet(A, B) {
	var a = A.toLowerCase();
	var b = B.toLowerCase();

	if (a < b) {
		return -1;
	}
	if (a > b) {
		return 1;
	}
	return 0; //default return value (no sorting)
}

export function sortByLongerLength(A, B) {
	var a = A.length;
	var b = B.length;

	if (a > b) {
		return -1;
	}
	if (a < b) {
		return 1;
	}
	return 0; //default return value (no sorting)
}

export function sortObject(src, comparator) {
	var out = Object.create(null);

	Object.keys(src).sort(comparator).forEach(function (key) {
		if (typeof src[key] == 'object' &&
			!Array.isArray(src[key]) &&
			!(src[key] instanceof RegExp)
		) {
			out[key] = sortObject(src[key], comparator); //run function again
			return;
		} else {
			out[key] = src[key];
		}
	});

	return out;
}

export function mergeObjects(target, src) {
	var a = target || Object.create(null);
	var b = src || Object.create(null);

	// merge b into a
	Object.keys(b).forEach(function (key) {
		a[key] = (a[key] || 0) + (b[key] || 0);
	});
}

export function mergeWithinObject(a, retainWord) {
	// helper function
	function mergeEntries(subObject, key) {
		subObject[retainWord] = (subObject[retainWord] || 0) + subObject[key];
		delete subObject[key];
	}

	Object.keys(a).forEach(function (mainKey) {
		if (mainKey !== 'defined') {
			var subObject = a[mainKey];
			// console.log('subObject', subObject);

			Object.keys(subObject).forEach(function (key) {
				if (/s$/.test(retainWord)) {
					// retainWord is plural, so merge singular key into plural
					if (retainWord === key + 's') {
						mergeEntries(subObject, key);
					}
				} else {
					// retainWord is singular, so merge plural key into singular
					if (key === retainWord + 's') {
						mergeEntries(subObject, key);
					}
				}
			});
		}
	});
}

export function addBullet(strOrObj) {
	var string = typeof strOrObj === 'object' ? strOrObj[0] : strOrObj;
	return string.replace(/^/, '• ');
}

export function createFirstTable(pojo) {
	var tableArray = [
		['May be Circular', 'Used But Not Defined in Selection'] //header row
	];
	var circularTerms = pojo.circular.length ? pojo.circular.map(function (pathArray) {
		return pathArray.join(' ->\r\n').replace(/^/, '• ');
	}).join('\r\n') : '';
	var notDefinedTerms = pojo.notDefined ? pojo.notDefined.map(addBullet).join('\r\n') : '';
	var rowArray = [];
	rowArray.push(circularTerms);
	rowArray.push(notDefinedTerms);
	tableArray.push(rowArray);

	return tableArray;
}

export function createSecondTable(pojo) {
	var tableArray = [
		['Cross-Reference Definitions'] //header row
	];
	var crossRefs = pojo.crossRefs.length ? pojo.crossRefs.map(addBullet).join('\r\n') : '';
	var rowArray = [];
	rowArray.push(crossRefs);
	tableArray.push(rowArray);

	return tableArray;
}

export function createMainTable(pojo) {
	var tableArray = [
		['Term', 'Incorporates', 'Used By', 'Defined in Selection'] //header row
	];

	Object.keys(pojo).forEach(function (dt) {
		var incorpsObj = pojo[dt].incorps;
		var incorpsTerms = incorpsObj ? Object.keys(incorpsObj).map(addBullet).join('\r\n') : '';
		var usedByObj = pojo[dt].usedBy;
		var usedByTerms = usedByObj ? Object.keys(usedByObj).map(addBullet).join('\r\n') : '';

		var definedVal = pojo[dt].defined ? pojo[dt].defined : 0;
		var definedTerm = definedVal === 1 ? 'yes' : (definedVal === 2 ? 'yes per user' : '');

		var rowArray = [];
		rowArray.push(dt);
		rowArray.push(incorpsTerms);
		rowArray.push(usedByTerms);
		rowArray.push(definedTerm);
		tableArray.push(rowArray);
	});

	return tableArray;
}

export function insertTable(docBody, tableArray) {
	return docBody.insertTable(
		tableArray.length, //rowLength
		tableArray[0].length, //columnLength
		Word.InsertLocation.end, //insertPosition
		tableArray
	);
}
