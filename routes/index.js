// routes/index.js
const express = require("express");
const router = express.Router();

const stockRoutes = require("./stock");
const assetRoutes = require("./asset");
const historyRoutes = require("./history");
const dashboardRoutes = require("./dashboard");
const adminRoutes = require("./admin");
const farmRoutes = require("./farm");
const labelRoutes = require("./label");
const partRoutes = require("./part");
const authRoutes = require("./auth");
const auditRoutes = require("./audit");

router.use("/", stockRoutes);
router.use("/", assetRoutes);
router.use("/", historyRoutes);
router.use("/", dashboardRoutes);
router.use("/", adminRoutes);
router.use("/", farmRoutes);
router.use("/", labelRoutes);
router.use("/", partRoutes);
router.use("/", authRoutes);
router.use("/", auditRoutes);
 
module.exports = router;