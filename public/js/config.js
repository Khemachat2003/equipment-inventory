// public/js/config.js
const IMAGE_CACHE_VERSION = "1.0.0";
const STOCK_SITE = 'Intranin';

// State
let stockData = [];
let assetData = [];
let allChartData = {};
let currentChartRange = 14;
let historyChart = null;
let currentMonitorFarm = "ALL";
let allFarms = [];
let partCatalog = [];
let farmSites = [];
let currentPart = "ALL";
let partSearchTerm = "";
let _stockLoaded = false;
let partCatalogList = [];