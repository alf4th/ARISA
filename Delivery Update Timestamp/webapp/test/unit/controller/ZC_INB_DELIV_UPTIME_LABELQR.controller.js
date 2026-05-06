/*global QUnit*/

sap.ui.define([
	"zdelivuptime/delivuptimelabelqr/controller/ZC_INB_DELIV_UPTIME_LABELQR.controller"
], function (Controller) {
	"use strict";

	QUnit.module("ZC_INB_DELIV_UPTIME_LABELQR Controller");

	QUnit.test("I should test the ZC_INB_DELIV_UPTIME_LABELQR controller", function (assert) {
		var oAppController = new Controller();
		oAppController.onInit();
		assert.ok(oAppController);
	});

});
