/*global QUnit*/

sap.ui.define([
	"uploadinbounddelivery/uploadinbounddelivery/controller/inbound_delivery.controller"
], function (Controller) {
	"use strict";

	QUnit.module("inbound_delivery Controller");

	QUnit.test("I should test the inbound_delivery controller", function (assert) {
		var oAppController = new Controller();
		oAppController.onInit();
		assert.ok(oAppController);
	});

});
