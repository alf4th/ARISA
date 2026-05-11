/* global XLSX */
sap.ui.define([
  "sap/ui/core/mvc/Controller",
  "sap/ui/model/json/JSONModel",
  "sap/m/MessageBox"
], function(Controller, JSONModel, MessageBox) {
  "use strict";

  return Controller.extend("uploadinbounddelivery.uploadinbounddelivery.controller.CustomPage", {

    onInit: function() {
      this.getView().setModel(new JSONModel({ excelData: [] }));
      this._file = null;
      
      // ✅ Attach route pattern matched untuk clear data saat navigasi
      var oRouter = this.getOwnerComponent().getRouter();
      oRouter.getRoute("CustomPage").attachPatternMatched(this._onRouteMatched, this);
    },

    // ✅ Clear data setiap kali route ke CustomPage dipanggil
    _onRouteMatched: function() {
      this._clearData();
    },

    // ✅ Fungsi untuk clear semua data
    _clearData: function() {
      this._file = null;
      this.getView().getModel().setProperty("/excelData", []);
      
      // Clear file uploader jika ada
      var oFileUploader = this.byId("fileUploader"); // sesuaikan dengan ID FileUploader Anda
      if (oFileUploader) {
        oFileUploader.clear();
        oFileUploader.setValue(""); // tambahan untuk memastikan value kosong
      }
    },

    onFileChange: function(oEvent) {
      this._file = oEvent.getParameter("files")[0];
    },

    onCancel: function() {
      this._clearData();
    },

    onStatusPress: function(oEvent) {
      var oContext = oEvent.getSource().getBindingContext();
      var sMessage = oContext.getProperty("StatusMessage");
      var iNo = oContext.getProperty("No");
      if (sMessage) {
        MessageBox.error(sMessage, { title: "Row " + iNo + " — Validation Errors" });
      }
    },

    onPreview: function() {
      if (!this._file) {
        sap.m.MessageToast.show("Please choose an Excel file first");
        return;
      }

      var that = this;
      var reader = new FileReader();
      reader.onload = function(e) {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: 'array' });

        var sheetName = "Inbound template";
        var sheet = workbook.Sheets[sheetName];

        if (!sheet) {
          sap.m.MessageToast.show('Sheet "Inbound template" not found');
          return;
        }

        // Read column headers from Excel row 8 (index 7) to avoid positional mapping issues
        var aHeaders = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 7, defval: "" })[0];
        console.log("Excel headers:", aHeaders);

        var excelData = XLSX.utils.sheet_to_json(sheet, {
          header: aHeaders,
          defval: "",
          range: 11  // Row 12 is the first data row (skip rows 1-11)
        });

        var requiredFields = [
          'Container Number',
          'Proforma Invoice',
          'Bill of Lading',          
          'Pallet Number',
          'Delivery Date',
          'Delivery Time',
          'Quantity',
          'Arisa Part Number',
          'Plant',
          'Purchase Order Number',
          'Purchase Order Item',
          'Commercial Invoice / Packing List'
        ];
        excelData = excelData.filter(function(row) {
          var cleanRow = Object.values(row).map(function(value) {
            return String(value).trim();
          });
          return cleanRow.some(function(val) { return val !== ""; });
        });

        var normalizedData = excelData.map(function(rowData, idx) {
          ['Delivery Date', 'Delivery Time'].forEach(function(field) {
            if (rowData[field] && !isNaN(rowData[field])) {
              rowData[field] = String(parseInt(rowData[field], 10));
            }
          });

          if (rowData['Batch Number']) {
            rowData['Batch Number'] = String(rowData['Batch Number']).replace(/\.\d+$/, '');
          }
          if (rowData['Purchase Order Item']) {
            rowData['Purchase Order Item'] = String(rowData['Purchase Order Item']).replace(/\.\d+$/, '');
          }
          if (rowData['Purchase Order Number']) {
            rowData['Purchase Order Number'] = String(rowData['Purchase Order Number']).replace(/\.\d+$/, '');
          }

          rowData["Cipl"] = String(rowData["Commercial Invoice / Packing List"] || "");
          rowData["No"] = idx + 1;

          var rowErrors = [];
          requiredFields.forEach(function(field) {
            if (!rowData[field] || rowData[field].toString().trim() === "") {
              rowErrors.push(field + " is required");
            }
          });

          if (rowErrors.length > 0) {
            rowData["Status"] = "Error";
            rowData["StatusState"] = "Error";
            rowData["StatusMessage"] = rowErrors.join("\n");
          } else {
            rowData["Status"] = "Valid";
            rowData["StatusState"] = "Success";
            rowData["StatusMessage"] = "";
          }

          return rowData;
        });

        that.getView().getModel().setProperty("/excelData", normalizedData);
      };
      reader.readAsArrayBuffer(this._file);
    },


    onSave: function() {
      var oView = this.getView();
      var oModel = oView.getModel();
      var aData = oModel.getProperty("/excelData");
      
      if (!aData || aData.length === 0) {
        sap.m.MessageToast.show("No data to save. Please preview data first.");
        return;
      }

      var aErrorRows = aData.filter(function(item) { return item["StatusState"] === "Error"; });
      if (aErrorRows.length > 0) {
        MessageBox.error(
          "Please check and fill all mandatory fields in the uploaded document.\n\n" +
          aErrorRows.length + " row(s) have incomplete or missing required fields.\n" +
          "Click on the 'Error' status in the preview table to see details for each row.",
          { title: "Validation Failed — Cannot Save" }
        );
        return;
      }
      
      // Show loading with message
      oView.setBusy(true);
      sap.ui.core.BusyIndicator.show(0);
      
      var oODataModel = this.getOwnerComponent().getModel();
      var that = this;
      
      // Prepare all payloads
      var aPayloads = aData.map(function(item) {
        return {
          Model: item["Model"],
          ModelDescription: item["Model Description"],
          SoNumberLineItem: item["SO Number / Line Item"],
          ContainerNumber: item["Container Number"],
          ProformaInvoiceNumber: item["Proforma Invoice"],
          BillOfPayment: item["Bill of Lading"],
          CommercialInvoicePl: item["Cipl"],
          PurchaseOrderNumber: item["Purchase Order Number"],
          PalletNumber: item["Pallet Number"],
          GrossWeight: item["Gross Weight"],
          PurchaseOrderItem: item["Purchase Order Item"],
          Material: item["Arisa Part Number"],
          VendorPartNumber: item["Vendor Part Number"],
          ActualDeliveryQuantity: item["Quantity"].toFixed(4),
          DeliveryDate: item["Delivery Date"],
          DeliveryTime: item["Delivery Time"],
          Plant: item["Plant"],
          BatchNumber: item["Batch Number"],
          FreeGoodsIndicator: item["Free Goods Indicator"],
          NoteOfFreeGoods: item["Note of Free Goods"],
          PoRefForFreeGoods: item["PO Ref for Free Goods"],
          PoLineItemRef: item["PO Line item Ref for Free Goods"],
          Response: "-"
        };
      });
      console.log(aPayloads);
      // return ;
      // Create batch group
      var mParameters = {
        groupId: "batchCreate",
        changeSetId: "changeSet1"
      };
      
      // Store created contexts to wait for completion
      var aContexts = [];
      var oListBinding = oODataModel.bindList("/ZC_TBINB_DLV");
      
      // Add all creates to the batch
      aPayloads.forEach(function(oPayload) {
        var oContext = oListBinding.create(oPayload, mParameters);
        aContexts.push(oContext);
      });
      
      // Wait for all contexts to be created (persisted to backend)
      Promise.all(
        aContexts.map(function(oContext) {
          return oContext.created();
        })
      )
      .then(function(aResults) {
        // All records successfully created
        sap.ui.core.BusyIndicator.hide();
        oView.setBusy(false);
        
        sap.m.MessageBox.success(
          "All " + aResults.length + " records saved successfully!",
          {
            onClose: function() {
              that.getOwnerComponent().getRouter().navTo("RouteListPage");
            }
          }
        );
      })
      .catch(function(oError) {
        // Handle errors
        sap.ui.core.BusyIndicator.hide();
        oView.setBusy(false);
        
        var sErrorMsg = "Failed to save records.\n\n";
        
        if (oError.message) {
          sErrorMsg += "Error: " + oError.message;
        } else if (oError.error && oError.error.message) {
          sErrorMsg += "Error: " + oError.error.message;
        }
        
        // Try to get more details from error response
        if (oError.error && oError.error.details) {
          sErrorMsg += "\n\nDetails:\n";
          oError.error.details.forEach(function(detail) {
            sErrorMsg += "- " + detail.message + "\n";
          });
        }
        
        sap.m.MessageBox.error(sErrorMsg);
      });
    }

  });
});