/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/app.js");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/app.js":
/*!********************!*\
  !*** ./src/app.js ***!
  \********************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
eval("/* global fabric:true, Office:true, OfficeExtension:true, Word:true */\n\n\n\nvar _typeof = typeof Symbol === \"function\" && typeof Symbol.iterator === \"symbol\" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === \"function\" && obj.constructor === Symbol && obj !== Symbol.prototype ? \"symbol\" : typeof obj; };\n\n(function () {\n\tvar messageBanner;\n\n\tOffice.initialize = function () {\n\t\t$(document).ready(function () {\n\t\t\t// initialize FabricUI notification mechanism and hide it\n\t\t\tvar element = document.querySelector('.ms-MessageBanner');\n\t\t\tmessageBanner = new fabric.MessageBanner(element);\n\t\t\tmessageBanner.hideBanner();\n\n\t\t\t// check Office\n\t\t\tif (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {\n\t\t\t\tconsole.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');\n\t\t\t}\n\n\t\t\tvar docx = Office.context.document;\n\n\t\t\t// pull into 'live settings' the data (if any) that is stored in the file\n\t\t\tdocx.settings.refreshAsync(function () {\n\t\t\t\t// get userTerms from live settings and show them in ui\n\t\t\t\t['add', 'minus'].forEach(function (cmd) {\n\t\t\t\t\taddToShownUserTerms(cmd, docx.settings.get('userTerms-' + cmd) || []);\n\t\t\t\t});\n\t\t\t});\n\n\t\t\t$('#parse-button').on('click', parseParas);\n\t\t\t$('#button-text').text('Parse Selected Definitions');\n\n\t\t\t$('#user-term-add').on('keydown', function (e) {\n\t\t\t\tif (e.keyCode === 13) {\n\t\t\t\t\tkeydownHandler('add', $(this));\n\t\t\t\t}\n\t\t\t});\n\t\t\t$('#user-term-minus').on('keydown', function (e) {\n\t\t\t\tif (e.keyCode === 13) {\n\t\t\t\t\tkeydownHandler('minus', $(this));\n\t\t\t\t}\n\t\t\t});\n\n\t\t\t$('#user-terms-add-container').on('click', '.user-term', function () {\n\t\t\t\tremoveClickHandler('add', $(this));\n\t\t\t});\n\t\t\t$('#user-terms-minus-container').on('click', '.user-term', function () {\n\t\t\t\tremoveClickHandler('minus', $(this));\n\t\t\t});\n\t\t});\n\t};\n\n\t/* UI Functions */\n\tfunction keydownHandler(cmd, elm) {\n\t\tvar inpVal = elm.val().trim();\n\n\t\tif (!inpVal) {\n\t\t\treturn; //bail\n\t\t}\n\n\t\t// add to shown user terms if not a dupe\n\t\tif (getShownUserTerms(cmd).indexOf(inpVal) === -1) {\n\t\t\taddToShownUserTerms(cmd, [inpVal]);\n\t\t\telm.val(''); //clear input\n\t\t}\n\n\t\t// sync to settings if not a dupe\n\t\tvar docx = Office.context.document;\n\t\tvar userTerms = docx.settings.get('userTerms-' + cmd) || [];\n\t\tif (userTerms.indexOf(inpVal) === -1) {\n\t\t\tuserTerms.push(inpVal);\n\t\t\tuserTerms.sort(sortByAlphabet);\n\t\t\tdocx.settings.set('userTerms-' + cmd, userTerms);\n\t\t\tdocx.settings.saveAsync();\n\t\t}\n\t}\n\n\tfunction removeClickHandler(cmd, elm) {\n\t\tvar val = elm.text();\n\t\telm.remove();\n\n\t\t// sync to settings\n\t\tvar docx = Office.context.document;\n\t\tvar userTerms = docx.settings.get('userTerms-' + cmd);\n\t\tif (userTerms) {\n\t\t\tuserTerms.splice(userTerms.indexOf(val), 1);\n\t\t\tdocx.settings.set('userTerms-' + cmd, userTerms);\n\t\t\tdocx.settings.saveAsync();\n\t\t}\n\t}\n\n\tfunction getShownUserTerms(cmd) {\n\t\tvar userTerms = [];\n\n\t\t$('#user-terms-' + cmd + '-container .user-term').each(function () {\n\t\t\tuserTerms.push($(this).text());\n\t\t});\n\n\t\treturn userTerms;\n\t}\n\n\tfunction addToShownUserTerms(cmd, arrayOfTerms) {\n\t\tvar container = $('#user-terms-' + cmd + '-container');\n\t\tvar frag = document.createDocumentFragment();\n\n\t\tarrayOfTerms.forEach(function (term) {\n\t\t\tvar div = document.createElement('div');\n\t\t\tdiv.classList.add('user-term');\n\t\t\tdiv.textContent = term;\n\t\t\tfrag.appendChild(div);\n\t\t});\n\t\tcontainer.prepend(frag);\n\n\t\treturn container;\n\t}\n\n\tfunction showNotification(header, content) {\n\t\t$(\"#notification-header\").text(header);\n\t\t$(\"#notification-body\").text(content);\n\t\tmessageBanner.showBanner();\n\t\tmessageBanner.toggleExpansion();\n\t}\n\n\t/* Utility Functions */\n\tfunction errHandler(error) {\n\t\tconsole.log(\"Error: \" + error);\n\n\t\tif (error instanceof OfficeExtension.Error) {\n\t\t\tconsole.log(\"Debug info: \" + JSON.stringify(error.debugInfo));\n\t\t}\n\t}\n\n\tfunction createRexFromString(string, flags) {\n\t\tvar escapedString = string.replace(/[|\\\\{}()[\\]^$+*?.]/g, '\\\\$&');\n\t\treturn new RegExp(escapedString, flags);\n\t}\n\n\tfunction sortByAlphabet(A, B) {\n\t\tvar a = A.toLowerCase();\n\t\tvar b = B.toLowerCase();\n\n\t\tif (a < b) {\n\t\t\treturn -1;\n\t\t}\n\t\tif (a > b) {\n\t\t\treturn 1;\n\t\t}\n\t\treturn 0; //default return value (no sorting)\n\t}\n\n\tfunction sortByLongerLength(A, B) {\n\t\tvar a = A.length;\n\t\tvar b = B.length;\n\n\t\tif (a > b) {\n\t\t\treturn -1;\n\t\t}\n\t\tif (a < b) {\n\t\t\treturn 1;\n\t\t}\n\t\treturn 0; //default return value (no sorting)\n\t}\n\n\tfunction sortObject(src, comparator) {\n\t\tvar out = Object.create(null);\n\n\t\tObject.keys(src).sort(comparator).forEach(function (key) {\n\t\t\tif (_typeof(src[key]) == 'object' && !Array.isArray(src[key]) && !(src[key] instanceof RegExp)) {\n\t\t\t\tout[key] = sortObject(src[key], comparator); //run function again\n\t\t\t\treturn;\n\t\t\t} else {\n\t\t\t\tout[key] = src[key];\n\t\t\t}\n\t\t});\n\n\t\treturn out;\n\t}\n\n\tfunction mergeObjects(target, src) {\n\t\tvar a = target || Object.create(null);\n\t\tvar b = src || Object.create(null);\n\n\t\t// merge b into a\n\t\tObject.keys(b).forEach(function (key) {\n\t\t\ta[key] = (a[key] || 0) + (b[key] || 0);\n\t\t});\n\t}\n\n\tfunction mergeWithinObject(a, retainWord) {\n\t\t// helper function\n\t\tfunction mergeEntries(subObject, key) {\n\t\t\tsubObject[retainWord] = (subObject[retainWord] || 0) + subObject[key];\n\t\t\tdelete subObject[key];\n\t\t}\n\n\t\tObject.keys(a).forEach(function (mainKey) {\n\t\t\tif (mainKey !== 'defined') {\n\t\t\t\tvar subObject = a[mainKey];\n\t\t\t\t// console.log('subObject', subObject);\n\n\t\t\t\tObject.keys(subObject).forEach(function (key) {\n\t\t\t\t\tif (/s$/.test(retainWord)) {\n\t\t\t\t\t\t// retainWord is plural, so merge singular key into plural\n\t\t\t\t\t\tif (retainWord === key + 's') {\n\t\t\t\t\t\t\tmergeEntries(subObject, key);\n\t\t\t\t\t\t}\n\t\t\t\t\t} else {\n\t\t\t\t\t\t// retainWord is singular, so merge plural key into singular\n\t\t\t\t\t\tif (key === retainWord + 's') {\n\t\t\t\t\t\t\tmergeEntries(subObject, key);\n\t\t\t\t\t\t}\n\t\t\t\t\t}\n\t\t\t\t});\n\t\t\t}\n\t\t});\n\t}\n\n\tfunction addBullet(strOrObj) {\n\t\tvar string = (typeof strOrObj === 'undefined' ? 'undefined' : _typeof(strOrObj)) === 'object' ? strOrObj[0] : strOrObj;\n\t\treturn string.replace(/^/, '• ');\n\t}\n\n\tfunction createFirstTable(pojo) {\n\t\tvar tableArray = [['May be Circular', 'Used But Not Defined in Selection'] //header row\n\t\t];\n\t\tvar circularTerms = pojo.circular.length ? pojo.circular.map(function (pathArray) {\n\t\t\treturn pathArray.join(' ->\\r\\n').replace(/^/, '• ');\n\t\t}).join('\\r\\n') : '';\n\t\tvar notDefinedTerms = pojo.notDefined ? pojo.notDefined.map(addBullet).join('\\r\\n') : '';\n\t\tvar rowArray = [];\n\t\trowArray.push(circularTerms);\n\t\trowArray.push(notDefinedTerms);\n\t\ttableArray.push(rowArray);\n\n\t\treturn tableArray;\n\t}\n\n\tfunction createSecondTable(pojo) {\n\t\tvar tableArray = [['Cross-Reference Definitions'] //header row\n\t\t];\n\t\tvar crossRefs = pojo.crossRefs.length ? pojo.crossRefs.map(addBullet).join('\\r\\n') : '';\n\t\tvar rowArray = [];\n\t\trowArray.push(crossRefs);\n\t\ttableArray.push(rowArray);\n\n\t\treturn tableArray;\n\t}\n\n\tfunction createMainTable(pojo) {\n\t\tvar tableArray = [['Term', 'Incorporates', 'Used By', 'Defined in Selection'] //header row\n\t\t];\n\n\t\tObject.keys(pojo).forEach(function (dt) {\n\t\t\tvar incorpsObj = pojo[dt].incorps;\n\t\t\tvar incorpsTerms = incorpsObj ? Object.keys(incorpsObj).map(addBullet).join('\\r\\n') : '';\n\t\t\tvar usedByObj = pojo[dt].usedBy;\n\t\t\tvar usedByTerms = usedByObj ? Object.keys(usedByObj).map(addBullet).join('\\r\\n') : '';\n\n\t\t\tvar definedVal = pojo[dt].defined ? pojo[dt].defined : 0;\n\t\t\tvar definedTerm = definedVal === 1 ? 'yes' : definedVal === 2 ? 'yes per user' : '';\n\n\t\t\tvar rowArray = [];\n\t\t\trowArray.push(dt);\n\t\t\trowArray.push(incorpsTerms);\n\t\t\trowArray.push(usedByTerms);\n\t\t\trowArray.push(definedTerm);\n\t\t\ttableArray.push(rowArray);\n\t\t});\n\n\t\treturn tableArray;\n\t}\n\n\tfunction insertTable(docBody, tableArray) {\n\t\treturn docBody.insertTable(tableArray.length, //rowLength\n\t\ttableArray[0].length, //columnLength\n\t\tWord.InsertLocation.end, //insertPosition\n\t\ttableArray);\n\t}\n\n\t/* Operative Function */\n\tfunction parseParas() {\n\t\tWord.run(function (context) {\n\t\t\t/* var data = context.document.body.paragraphs; */\n\t\t\tvar data = context.document.getSelection().paragraphs;\n\t\t\tcontext.load(data, 'text');\n\n\t\t\treturn context.sync().then(function () {\n\t\t\t\tvar paras = [];\n\t\t\t\tdata.items.forEach(function (item) {\n\t\t\t\t\tparas.push(item.text.trim());\n\t\t\t\t});\n\t\t\t\t// console.log(paras);\n\n\t\t\t\t/* START HERE */\n\t\t\t\tvar rexPojo = Object.create(null);\n\t\t\t\tvar pojo = Object.create(null);\n\n\t\t\t\tvar rexqtPhrase = /^“[^“]+”([^“]{1,7}“[^“]+”)*/;\n\t\t\t\tvar rexqts = /“[^“]+”/g;\n\t\t\t\tvar rexInitCaps = /((([A-Z][\\w\\-]+|\\d{4})\\s?(of|and)?\\s?)(\\d{4}(\\-\\d{1,2})?\\s?)?)+/g;\n\t\t\t\tvar rexLeadArticles = /^(A|An|If|The|This|That|Each|Such|Every)\\s/;\n\t\t\t\tvar badLoneWords = ['for', 'with', 'each', 'if', 'the', 'this', 'none', 'such', 'every', 'in', 'on'];\n\n\t\t\t\t/* 'REXPOJO' PASS */\n\t\t\t\t// populate rexPojo with every quoted term appearing at the beginning of each para\n\t\t\t\tparas.forEach(function (p) {\n\t\t\t\t\tvar qtPhrase = p.match(rexqtPhrase);\n\t\t\t\t\tif (qtPhrase) {\n\t\t\t\t\t\t(qtPhrase[0].match(rexqts) || []).map(function (qt) {\n\t\t\t\t\t\t\treturn qt.replace(/[“”,]/g, '');\n\t\t\t\t\t\t}).forEach(function (dt) {\n\t\t\t\t\t\t\trexPojo[dt] = createRexFromString(dt, 'g'); //put in rexPojo\n\t\t\t\t\t\t});\n\t\t\t\t\t}\n\t\t\t\t});\n\t\t\t\t// console.log('rexPojo before adding userTerms', rexPojo);\n\n\t\t\t\t// add user specified terms (held in live settings) to rexPojo\n\t\t\t\t// also, store them in a variable for adjustments below\n\t\t\t\tvar userTermsAdded = Office.context.document.settings.get('userTerms-add') || [];\n\t\t\t\tuserTermsAdded.forEach(function (uta) {\n\t\t\t\t\trexPojo[uta] = createRexFromString(uta, 'g'); //put in rexPojo\n\t\t\t\t});\n\n\t\t\t\t// sort rexPojo by length (so longer ones get removed from para first per below, and\n\t\t\t\t// avoid creating fragments of defined terms that would be caught later by init caps)\n\t\t\t\tvar sortedRexPojo = sortObject(rexPojo, sortByLongerLength); /*key*/\n\t\t\t\t// console.log('sortedRexPojo', sortedRexPojo);\n\n\t\t\t\t/* 'INCORPS' PASS */\n\t\t\t\t// populate 'incorps'\n\t\t\t\tvar last_dts;\n\t\t\t\tparas.forEach(function (p) {\n\t\t\t\t\tvar dts;\n\t\t\t\t\tvar qtPhrase = p.match(rexqtPhrase);\n\t\t\t\t\tif (qtPhrase) {\n\t\t\t\t\t\tlast_dts = dts = qtPhrase[0].match(rexqts).map(function (qt) {\n\t\t\t\t\t\t\treturn qt.replace(/[“”\\,]/g, '');\n\t\t\t\t\t\t});\n\t\t\t\t\t\t// the above replicates the rexPojo Pass, except that, here, we track last_dts\n\t\t\t\t\t\t// to link dts to paras that don't have quoted defined terms at their beginnings\n\t\t\t\t\t} else {\n\t\t\t\t\t\tdts = last_dts; //use last_dts (since this para doesn't have its own dts)\n\t\t\t\t\t}\n\n\t\t\t\t\t(dts || []).forEach(function (t) {\n\t\t\t\t\t\tif (!pojo[t]) {\n\t\t\t\t\t\t\tpojo[t] = Object.create(null); //add defined term to pojo\n\t\t\t\t\t\t}\n\t\t\t\t\t\tpojo[t].defined = 1; //track if t is a \"defined term\"\n\n\t\t\t\t\t\t// apply sortedRexPojo\n\t\t\t\t\t\tObject.keys(sortedRexPojo).forEach(function (rex) {\n\t\t\t\t\t\t\t(p.match(rex) || []).filter(function (n) {\n\t\t\t\t\t\t\t\treturn dts.indexOf(n) === -1; //exclude any defined terms (i.e., itself)\n\t\t\t\t\t\t\t}).forEach(function (n) {\n\t\t\t\t\t\t\t\tif (!pojo[t].incorps) {\n\t\t\t\t\t\t\t\t\tpojo[t].incorps = Object.create(null);\n\t\t\t\t\t\t\t\t}\n\t\t\t\t\t\t\t\tpojo[t].incorps[n] = pojo[t].incorps[n] + 1 || 1;\n\t\t\t\t\t\t\t});\n\n\t\t\t\t\t\t\t// remove rex from para to avoid catching fragments later /*key*/\n\t\t\t\t\t\t\tp = p.replace(rex, '');\n\t\t\t\t\t\t});\n\n\t\t\t\t\t\t// apply init caps\n\t\t\t\t\t\t(p.match(/“[^“]+”/g) || []). //get all quoted terms contained in the p\n\t\t\t\t\t\tmap(function (qt) {\n\t\t\t\t\t\t\treturn qt.replace(/[“”\\,]/g, ''); //remove their quotation marks\n\t\t\t\t\t\t}).filter(function (dt) {\n\t\t\t\t\t\t\treturn (/^[a-z]/.test(dt)\n\t\t\t\t\t\t\t); //keep those whose first letter is lower case\n\t\t\t\t\t\t}).concat(p.match(rexInitCaps) || []) //CONCAT with new array of init caps\n\t\t\t\t\t\t.map(function (n) {\n\t\t\t\t\t\t\treturn n.trim() //trim leading and trailing spaces\n\t\t\t\t\t\t\t.replace(rexLeadArticles, '') //trim leading articles\n\t\t\t\t\t\t\t.replace(/\\s(of|and)$/, ''); //trim trailing of|and;\n\t\t\t\t\t\t}).filter(function (n) {\n\t\t\t\t\t\t\treturn n.length && dts.indexOf(n) === -1; //exclude any defined terms\n\t\t\t\t\t\t}).filter(function (n) {\n\t\t\t\t\t\t\treturn badLoneWords.indexOf(n.toLowerCase()) === -1; //exclude badLoneWords\n\t\t\t\t\t\t}).filter(function (n) {\n\t\t\t\t\t\t\treturn !/^\\d+$/.test(n); //exclude number-only strings\n\t\t\t\t\t\t}).forEach(function (n) {\n\t\t\t\t\t\t\tif (!pojo[t].incorps) {\n\t\t\t\t\t\t\t\tpojo[t].incorps = Object.create(null);\n\t\t\t\t\t\t\t}\n\t\t\t\t\t\t\tpojo[t].incorps[n] = pojo[t].incorps[n] + 1 || 1;\n\t\t\t\t\t\t});\n\t\t\t\t\t});\n\t\t\t\t});\n\n\t\t\t\t/* REMOVE PASS */\n\t\t\t\t(Office.context.document.settings.get('userTerms-minus') || []).forEach(function (utm) {\n\t\t\t\t\tObject.keys(pojo).forEach(function (key) {\n\t\t\t\t\t\tif (key === utm) {\n\t\t\t\t\t\t\tdelete pojo[key];\n\t\t\t\t\t\t} else {\n\t\t\t\t\t\t\tvar incorpsObj = pojo[key].incorps;\n\n\t\t\t\t\t\t\tif (incorpsObj) {\n\t\t\t\t\t\t\t\tObject.keys(incorpsObj).forEach(function (term) {\n\t\t\t\t\t\t\t\t\tif (term === utm) {\n\t\t\t\t\t\t\t\t\t\tdelete pojo[key].incorps[term];\n\t\t\t\t\t\t\t\t\t}\n\t\t\t\t\t\t\t\t});\n\t\t\t\t\t\t\t}\n\t\t\t\t\t\t}\n\t\t\t\t\t});\n\t\t\t\t});\n\n\t\t\t\t/* 'USEDBY' PASS */\n\t\t\t\t// use incorps data to populate 'usedBy'\n\t\t\t\tObject.keys(pojo).forEach(function (t) {\n\t\t\t\t\t// console.log(pojo[t].incorps);\n\t\t\t\t\tif (pojo[t].incorps) {\n\t\t\t\t\t\tObject.keys(pojo[t].incorps).forEach(function (n) {\n\t\t\t\t\t\t\t// console.log(n);\n\t\t\t\t\t\t\tif (!pojo[n]) {\n\t\t\t\t\t\t\t\tpojo[n] = Object.create(null);\n\t\t\t\t\t\t\t}\n\t\t\t\t\t\t\tif (!pojo[n].usedBy) {\n\t\t\t\t\t\t\t\tpojo[n].usedBy = Object.create(null);\n\t\t\t\t\t\t\t}\n\t\t\t\t\t\t\tvar val = pojo[t].incorps[n];\n\t\t\t\t\t\t\tpojo[n].usedBy[t] = pojo[n].usedBy[t] + val || val;\n\t\t\t\t\t\t});\n\t\t\t\t\t}\n\t\t\t\t});\n\n\t\t\t\tvar sortedPojo = sortObject(pojo, sortByAlphabet);\n\t\t\t\t// console.log('debug sortedPojo', sortedPojo);\n\n\t\t\t\t/* PLURAL PASS */\n\t\t\t\tvar retainWords = [];\n\t\t\t\tObject.keys(sortedPojo).forEach(function (plural, i, self) {\n\t\t\t\t\tif (i > 0) {\n\t\t\t\t\t\tvar singular = self[i - 1]; //previous key\n\n\t\t\t\t\t\tif (plural === singular + 's') {\n\t\t\t\t\t\t\t// console.log(singular, '+s ===', plural);\n\t\t\t\t\t\t\tif (sortedPojo[plural].defined && !sortedPojo[singular].defined) {\n\t\t\t\t\t\t\t\t// retain plural form (as target)\n\t\t\t\t\t\t\t\tretainWords.push(plural);\n\t\t\t\t\t\t\t\tmergeObjects(sortedPojo[plural].incorps, sortedPojo[singular].incorps);\n\t\t\t\t\t\t\t\tmergeObjects(sortedPojo[plural].usedBy, sortedPojo[singular].usedBy);\n\t\t\t\t\t\t\t\tdelete sortedPojo[singular];\n\t\t\t\t\t\t\t} else if (!sortedPojo[plural].defined) {\n\t\t\t\t\t\t\t\t// retain singular form (as target)\n\t\t\t\t\t\t\t\tretainWords.push(singular);\n\t\t\t\t\t\t\t\tmergeObjects(sortedPojo[singular].incorps, sortedPojo[plural].incorps);\n\t\t\t\t\t\t\t\tmergeObjects(sortedPojo[singular].usedBy, sortedPojo[plural].usedBy);\n\t\t\t\t\t\t\t\tdelete sortedPojo[plural];\n\t\t\t\t\t\t\t}\n\t\t\t\t\t\t}\n\t\t\t\t\t}\n\t\t\t\t});\n\n\t\t\t\t// merge plural/singular terms contained within each object in sortedPojo\n\t\t\t\tretainWords.forEach(function (word) {\n\t\t\t\t\tObject.keys(sortedPojo).forEach(function (term) {\n\t\t\t\t\t\tmergeWithinObject(sortedPojo[term], word);\n\t\t\t\t\t});\n\t\t\t\t});\n\n\t\t\t\t/* ANALYSIS PASS */\n\t\t\t\tvar analysisPojo = Object.create(null);\n\t\t\t\tvar sortedPojoKeys = Object.keys(sortedPojo);\n\n\t\t\t\t/* Pick out terms that are not defined in selection */\n\t\t\t\tsortedPojoKeys.forEach(function (term) {\n\t\t\t\t\tif (!sortedPojo[term].defined) {\n\t\t\t\t\t\tif (userTermsAdded.indexOf(term) !== -1) {\n\t\t\t\t\t\t\t// unless it is one of the userTermsAdded\n\t\t\t\t\t\t\tsortedPojo[term].defined = 2; //use 2 instead of 1 to distinguish\n\t\t\t\t\t\t} else {\n\t\t\t\t\t\t\tif (!analysisPojo.notDefined) {\n\t\t\t\t\t\t\t\t// use array (instead of another object) as value\n\t\t\t\t\t\t\t\tanalysisPojo.notDefined = [];\n\t\t\t\t\t\t\t}\n\t\t\t\t\t\t\tanalysisPojo.notDefined.push(term);\n\t\t\t\t\t\t}\n\t\t\t\t\t}\n\t\t\t\t});\n\n\t\t\t\t/* Find circular terms */\n\t\t\t\tvar circularPaths = [];\n\t\t\t\tfunction walker(caller, target, path, depth) {\n\t\t\t\t\tif (sortedPojo[caller].incorps) {\n\t\t\t\t\t\tObject.keys(sortedPojo[caller].incorps).forEach(function (n) {\n\t\t\t\t\t\t\t// using a deep clone of 'path' -- must do so when\n\t\t\t\t\t\t\t// recursively invoking walker function below\n\t\t\t\t\t\t\tvar clone = path.slice(0);\n\n\t\t\t\t\t\t\tif (n === target) {\n\t\t\t\t\t\t\t\t// clone.push(n); //can't push n b/c that screws up removal of dupes\n\t\t\t\t\t\t\t\tcircularPaths.push(clone);\n\t\t\t\t\t\t\t} else if (sortedPojo[n].incorps) {\n\t\t\t\t\t\t\t\tif (clone.length < depth && clone.indexOf(n) === -1) {\n\t\t\t\t\t\t\t\t\tclone.push(n);\n\t\t\t\t\t\t\t\t\twalker(n, target, clone, depth); //recursively invoke walker\n\t\t\t\t\t\t\t\t}\n\t\t\t\t\t\t\t}\n\t\t\t\t\t\t});\n\t\t\t\t\t}\n\t\t\t\t}\n\n\t\t\t\tsortedPojoKeys.forEach(function (term) {\n\t\t\t\t\twalker(term, term, [term], 6);\n\t\t\t\t});\n\n\t\t\t\tanalysisPojo.circular = circularPaths\n\t\t\t\t// remove dupe paths\n\t\t\t\t.filter(function (path, i, self) {\n\t\t\t\t\treturn i === self.findIndex(function (item) {\n\t\t\t\t\t\treturn item.slice(0).sort().join('') === path.slice(0).sort().join('');\n\t\t\t\t\t});\n\t\t\t\t})\n\t\t\t\t// add back in last path item\n\t\t\t\t.map(function (path) {\n\t\t\t\t\tpath.push(path[0]);\n\t\t\t\t\treturn path;\n\t\t\t\t});\n\n\t\t\t\t/* Pick cross-referenced definitions */\n\t\t\t\tvar rexFirstSentence = /^.+?\\.(?:\\s|$)/;\n\t\t\t\tanalysisPojo.crossRefs = paras.map(function (p) {\n\t\t\t\t\treturn p.match(rexFirstSentence);\n\t\t\t\t}).filter(function (sentence) {\n\t\t\t\t\treturn (/\\b(meaning|defined|definition)s*?\\b/.test(sentence)\n\t\t\t\t\t);\n\t\t\t\t}).filter(function (sentence) {\n\t\t\t\t\treturn (/^“/.test(sentence)\n\t\t\t\t\t);\n\t\t\t\t}).filter(function (sentence) {\n\t\t\t\t\treturn sentence[0].split(' ').length < 30;\n\t\t\t\t});\n\n\t\t\t\t// console.log(JSON.stringify(analysisPojo, null, 5));\n\t\t\t\t/* END HERE */\n\n\t\t\t\tif (!Object.keys(sortedPojo).length) {\n\t\t\t\t\tvar header = 'Error:';\n\t\t\t\t\tvar content = 'No definition paragraphs have been selected';\n\t\t\t\t\tshowNotification(header, content);\n\n\t\t\t\t\treturn context.sync(); //bail\n\t\t\t\t}\n\n\t\t\t\tvar firstTableArray = createFirstTable(analysisPojo);\n\t\t\t\tvar secondTableArray = createSecondTable(analysisPojo);\n\t\t\t\tvar mainTableArray = createMainTable(sortedPojo);\n\t\t\t\tvar newDoc = context.application.createDocument();\n\t\t\t\tcontext.load(newDoc);\n\n\t\t\t\treturn context.sync().then(function () {\n\t\t\t\t\t// console.log('newDoc', newDoc);\n\t\t\t\t\tvar firstTable = insertTable(newDoc.body, firstTableArray);\n\t\t\t\t\tfirstTable.headerRowCount = 1;\n\t\t\t\t\tfirstTable.style = 'List Table 4 - Accent 1';\n\t\t\t\t\tfirstTable.styleFirstColumn = false;\n\n\t\t\t\t\tvar secondTable = insertTable(newDoc.body, secondTableArray);\n\t\t\t\t\tsecondTable.headerRowCount = 1;\n\t\t\t\t\tsecondTable.style = 'List Table 4 - Accent 1';\n\t\t\t\t\tsecondTable.styleFirstColumn = false;\n\n\t\t\t\t\tvar mainTable = insertTable(newDoc.body, mainTableArray);\n\t\t\t\t\tmainTable.headerRowCount = 1;\n\t\t\t\t\tmainTable.style = 'List Table 4 - Accent 1';\n\n\t\t\t\t\treturn context.sync().then(function () {\n\t\t\t\t\t\tnewDoc.open();\n\n\t\t\t\t\t\treturn context.sync();\n\t\t\t\t\t}).catch(errHandler);\n\t\t\t\t}).catch(errHandler);\n\t\t\t}).catch(errHandler);\n\t\t}).catch(errHandler);\n\t}\n})();\n\n//# sourceURL=webpack:///./src/app.js?");

/***/ })

/******/ });