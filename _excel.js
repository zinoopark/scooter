"use strict";

function _typeof(obj) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (obj) { return typeof obj; } : function (obj) { return obj && "function" == typeof Symbol && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; }, _typeof(obj); }

var xlsx = _interopRequireWildcard(require("xlsx"));

var _path = _interopRequireDefault(require("path"));

var _makearr = require("./_makearr.js");

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { "default": obj }; }

function _getRequireWildcardCache(nodeInterop) { if (typeof WeakMap !== "function") return null; var cacheBabelInterop = new WeakMap(); var cacheNodeInterop = new WeakMap(); return (_getRequireWildcardCache = function _getRequireWildcardCache(nodeInterop) { return nodeInterop ? cacheNodeInterop : cacheBabelInterop; })(nodeInterop); }

function _interopRequireWildcard(obj, nodeInterop) { if (!nodeInterop && obj && obj.__esModule) { return obj; } if (obj === null || _typeof(obj) !== "object" && typeof obj !== "function") { return { "default": obj }; } var cache = _getRequireWildcardCache(nodeInterop); if (cache && cache.has(obj)) { return cache.get(obj); } var newObj = {}; var hasPropertyDescriptor = Object.defineProperty && Object.getOwnPropertyDescriptor; for (var key in obj) { if (key !== "default" && Object.prototype.hasOwnProperty.call(obj, key)) { var desc = hasPropertyDescriptor ? Object.getOwnPropertyDescriptor(obj, key) : null; if (desc && (desc.get || desc.set)) { Object.defineProperty(newObj, key, desc); } else { newObj[key] = obj[key]; } } } newObj["default"] = obj; if (cache) { cache.set(obj, newObj); } return newObj; }

function _toConsumableArray(arr) { return _arrayWithoutHoles(arr) || _iterableToArray(arr) || _unsupportedIterableToArray(arr) || _nonIterableSpread(); }

function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }

function _unsupportedIterableToArray(o, minLen) { if (!o) return; if (typeof o === "string") return _arrayLikeToArray(o, minLen); var n = Object.prototype.toString.call(o).slice(8, -1); if (n === "Object" && o.constructor) n = o.constructor.name; if (n === "Map" || n === "Set") return Array.from(o); if (n === "Arguments" || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)) return _arrayLikeToArray(o, minLen); }

function _iterableToArray(iter) { if (typeof Symbol !== "undefined" && iter[Symbol.iterator] != null || iter["@@iterator"] != null) return Array.from(iter); }

function _arrayWithoutHoles(arr) { if (Array.isArray(arr)) return _arrayLikeToArray(arr); }

function _arrayLikeToArray(arr, len) { if (len == null || len > arr.length) len = arr.length; for (var i = 0, arr2 = new Array(len); i < len; i++) { arr2[i] = arr[i]; } return arr2; }

var today = new Date();
var year = today.getFullYear();
var month = today.getMonth() + 1;
var day = today.getDate();
var hour = today.getHours();
var minute = today.getMinutes();
month = month > 9 ? month : "0".concat(month);
day = day > 9 ? day : "0".concat(day);
hour = hour > 9 ? hour : "0".concat(hour);
minute = minute > 9 ? minute : "0".concat(minute);
var json = (0, _makearr.getHtml)().then(function (html) {
  return (0, _makearr.makearr)(html);
}).then(function (res) {
  console.log(res);
  var json = [];
  var arr = ['전체', '우선순위', '법인', '배달용', '우선비대상'];
  var workbook = xlsx.utils.book_new();
  var worksheet = xlsx.utils.aoa_to_sheet([[]]);
  xlsx.utils.sheet_add_aoa(worksheet, [["시도", "지역구분", "민간공고대수", '', '', '', '', "접수대수", '', '', '', '', "출고대수", '', '', '', '', "접수잔여대수", '', '', '', '', "출고잔여대수", '', '', '', '']], {
    origin: 0
  });
  xlsx.utils.sheet_add_aoa(worksheet, [["", "", '전체', '우선순위', '법인', '배달용', '우선비대상', '전체', '우선순위', '법인', '배달용', '우선비대상', '전체', '우선순위', '법인', '배달용', '우선비대상', '전체', '우선순위', '법인', '배달용', '우선비대상', '전체', '우선순위', '법인', '배달용', '우선비대상']], {
    origin: -1
  });
  res.forEach(function (e, i) {
    xlsx.utils.sheet_add_aoa(worksheet, [[e.시도, e.지역구분].concat(_toConsumableArray(e.민간공고대수), _toConsumableArray(e.접수대수), _toConsumableArray(e.출고대수), _toConsumableArray(e.접수잔여대수), _toConsumableArray(e.출고잔여대수))], {
      origin: -1
    });
  });
  xlsx.utils.book_append_sheet(workbook, worksheet, 'sheet1');
  xlsx.writeFile(workbook, _path["default"].join(__dirname, "\uC794\uC5EC\uB300\uC218\uD604\uD669_".concat(year).concat(month).concat(day, "_").concat(hour, "\uC2DC").concat(minute, "\uBD84.xlsx")));
}); // const arrayData = [['kim', 23], ['park', 24]];
// const jsonData = [
//   {name: 'kim', age: 23},
//   {name: 'park', age: 24},
// ]
//   const workbook = xlsx.utils.book_new();
//   const worksheet = xlsx.utils.json_to_sheet(jsonData);
//   xlsx.utils.book_append_sheet(workbook, worksheet, 'sheet1');
//   xlsx.writeFile(workbook, path.join(__dirname, 'array_to_sheet_result.xlsx'));
