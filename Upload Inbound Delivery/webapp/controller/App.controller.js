sap.ui.define([
  "sap/ui/core/mvc/Controller"
], (BaseController) => {
  "use strict";

  return BaseController.extend("uploadinbounddelivery.uploadinbounddelivery.controller.App", {
onInit() {
    // Remove limited width class
    const shell = document.querySelector('.sapUShellApplicationContainer');
        if (shell) {
            shell.classList.remove('sapUShellApplicationContainerLimitedWidth');
        }
    }
  });
});