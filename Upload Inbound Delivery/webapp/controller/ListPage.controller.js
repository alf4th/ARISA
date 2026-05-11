sap.ui.define([
  "sap/ui/core/mvc/Controller",
  "sap/ui/model/Filter",
  "sap/ui/model/FilterOperator",
  "sap/ui/model/Sorter",  // Make sure this is included
  "sap/ui/export/Spreadsheet",
  "sap/m/MessageBox"
], function(Controller, Filter, FilterOperator, Sorter, Spreadsheet, MessageBox) {
  "use strict";

  return Controller.extend("uploadinbounddelivery.uploadinbounddelivery.controller.ListPage", {
    
    onInit: function() {
      var oRouter = this.getOwnerComponent().getRouter();
      oRouter.getRoute("RouteListPage").attachPatternMatched(this._onRouteMatched, this);
    },

    _onRouteMatched: function() {
      // ✅ Refresh model untuk memastikan data terbaru
      var oModel = this.getView().getModel();
      if (oModel && oModel.refresh) {
        oModel.refresh();
      }
      
      // ✅ Refresh binding table
      var oTable = this.byId("odataTable");
      if (oTable) {
        var oBinding = oTable.getBinding("rows");
        if (oBinding) {
          oBinding.refresh();
        }
      }
      
      // Add default sorting
      var oTable = this.byId("odataTable");
      if (oTable) {
        var oBinding = oTable.getBinding("rows");
        if (oBinding) {
          var oSorter = new Sorter("LastChangedAt", true); // true for descending order
          oBinding.sort(oSorter);
        }
      }
      
      // Apply filter setelah refresh
      this._applyFilterToNewlyUploadedTab();
    },

    _applyFilterToNewlyUploadedTab: function() {
      var oTable = this.byId("odataTable");
      if (oTable) {
        var oBinding = oTable.getBinding("rows");
        if (oBinding) {
          var aFilters = [
            new Filter("IsCleared", FilterOperator.EQ, 0)
          ];
          oBinding.filter(aFilters);
        }
      }
    },

    onTabSelect: function(oEvent) {
      var sKey = oEvent.getParameter("key");
      var oSorter = new Sorter("LastChangedAt", true);
      
      if (sKey === "newly") {
        var oTable = this.byId("odataTable");
        if (oTable) {
          var oBinding = oTable.getBinding("rows");
          if (oBinding) {
            var aFilters = [
              new Filter("IsCleared", FilterOperator.EQ, 0)
            ];
            oBinding.filter(aFilters);
            oBinding.sort(oSorter);
          }
        }
      } else if (sKey === "history") {
        var oHistoryTable = this.byId("historyTable");
        if (oHistoryTable) {
          var oHistoryBinding = oHistoryTable.getBinding("rows");
          if (oHistoryBinding) {
            var aHistoryFilters = [
              new Filter("IsCleared", FilterOperator.EQ, 1)
            ];
            oHistoryBinding.filter(aHistoryFilters);
            oHistoryBinding.sort(oSorter);
          }
        }
      }
    },

    _getCurrentTable: function() {
      var oIconTabBar = this.byId("idIconTabBar");
      var sSelectedKey = oIconTabBar.getSelectedKey();
      
      if (sSelectedKey === "newly") {
        return this.byId("odataTable");
      } else {
        return this.byId("historyTable");
      }
    },

    onClearData: function() {
      var oTable = this.byId("odataTable");
      var aSelectedIndices = oTable.getSelectedIndices();
      
      if (aSelectedIndices.length === 0) {
        MessageBox.warning("Please select at least one item to clear");
        return;
      }
      
      var that = this;
      MessageBox.confirm(
        "Are you sure you want to clear " + aSelectedIndices.length + " selected item(s)?",
        {
          title: "Confirm Clear",
          onClose: function(oAction) {
            if (oAction === MessageBox.Action.OK) {
              that._clearSelectedData(aSelectedIndices);
            }
          }
        }
      );
    },

    _clearSelectedData: function(aSelectedIndices) {
      var oTable = this.byId("odataTable");
      var oModel = oTable.getModel();
      var that = this;
      var iSuccessCount = 0;
      var iErrorCount = 0;
      var iTotalCount = aSelectedIndices.length;
      
      // Check if OData V2 or V4
      var bIsV4 = oModel.isA("sap.ui.model.odata.v4.ODataModel");
      
      if (bIsV4) {
        // OData V4 approach
        aSelectedIndices.forEach(function(iIndex) {
          var oContext = oTable.getContextByIndex(iIndex);
          if (oContext) {
            oContext.setProperty("IsCleared", 1);
          }
        });
        
        oModel.submitBatch("updateGroup").then(function() {
          MessageBox.success("Data cleared successfully!", {
            onClose: function() {
              oTable.clearSelection();
              oTable.getBinding("rows").refresh();
            }
          });
        }).catch(function(oError) {
          MessageBox.error("Failed to clear data: " + (oError.message || "Unknown error"));
          oTable.getBinding("rows").refresh();
        });
      } else {
        // OData V2 approach
        oModel.setUseBatch(true);
        oModel.setDeferredGroups(["changes"]);
        
        aSelectedIndices.forEach(function(iIndex) {
          var oContext = oTable.getContextByIndex(iIndex);
          if (oContext) {
            var sPath = oContext.getPath();
            oModel.setProperty(sPath + "/IsCleared", 1);
          }
        });
        
        oModel.submitChanges({
          groupId: "changes",
          success: function(oData) {
            MessageBox.success("Data cleared successfully!", {
              onClose: function() {
                oTable.clearSelection();
                oTable.getBinding("rows").refresh();
              }
            });
          },
          error: function(oError) {
            MessageBox.error("Failed to clear data: " + (oError.message || "Unknown error"));
            oTable.getBinding("rows").refresh();
          }
        });
      }
    },

    onExport: function() {
      var oTable = this._getCurrentTable();
      var oBinding = oTable.getBinding("rows");
      var oModel = oTable.getModel();
      
      // Ambil data dari binding
      var aData = [];
      var aContexts = oBinding.getContexts(0, oBinding.getLength());
      
      aContexts.forEach(function(oContext) {
        var oData = oContext.getObject();
        aData.push(oData);
      });
      
      if (aData.length === 0) {
        MessageBox.warning("No data to export");
        return;
      }
      
      // Konfigurasi kolom untuk export
      var aCols = [
        { label: "Model", property: "Model" },
        { label: "Model Description", property: "ModelDescription" },
        { label: "SO Number/Line Item", property: "SoNumberLineItem" },
        { label: "Container Number", property: "ContainerNumber" },
        { label: "Proforma Invoice", property: "ProformaInvoiceNumber" },
        { label: "Bill of Lading", property: "BillOfLading" },
        { label: "Commercial Invoice/Packing List", property: "CommercialInvoicePl" },
        { label: "Purchase Order Number", property: "PurchaseOrderNumber" },
        { label: "Pallet Number", property: "PalletNumber" },
        { label: "Gross Weight", property: "GrossWeight" },
        { label: "Purchase Order Item", property: "PurchaseOrderItem" },
        { label: "Arisa Part Number", property: "Material" },
        { label: "Vendor Part Number", property: "VendorPartNumber" },
        { label: "Quantity", property: "ActualDeliveryQuantity" },
        { label: "Delivery Date", property: "DeliveryDate" },
        { label: "Delivery Time", property: "DeliveryTime" },
        { label: "Plant", property: "Plant" },
        { label: "Batch Number", property: "BatchNumber" },
        { label: "Free Goods Indicator", property: "FreeGoodsIndicator" },
        { label: "Note of Free Goods", property: "NoteOfFreeGoods" },
        { label: "PO Ref for Free Goods", property: "PoRefForFreeGoods" },
        { label: "PO Line Item Ref", property: "PoLineItemRef" },
        { label: "Response", property: "Response" },
        { label: "Created By", property: "CreatedBy" },
        { label: "Changed By", property: "ChangedBy" },
        { label: "Local Last Changed At", property: "LocalLastChangedAt" },
        { label: "Last Changed At", property: "LastChangedAt" }
      ];
      
      // Export ke Excel
      var oSettings = {
        workbook: {
          columns: aCols
        },
        dataSource: aData,
        fileName: "InboundDelivery_" + new Date().getTime() + ".xlsx"
      };
      
      var oSpreadsheet = new Spreadsheet(oSettings);
      oSpreadsheet.build()
        .then(function() {
          MessageBox.success("Export completed successfully!");
        })
        .catch(function(oError) {
          MessageBox.error("Export failed: " + oError.message);
        });
    },

    onSearch: function(oEvent) {
      var sQuery = oEvent.getParameter("query");
      if (sQuery === undefined || sQuery === null) {
        sQuery = oEvent.getParameter("newValue") || "";
      }

      var oIconTabBar = this.byId("idIconTabBar");
      var sKey = oIconTabBar.getSelectedKey();
      var oTable = sKey === "newly" ? this.byId("odataTable") : this.byId("historyTable");
      var oBinding = oTable && oTable.getBinding("rows");
      if (!oBinding) return;

      var iIsCleared = sKey === "history" ? 1 : 0;
      var aFilters = [new Filter("IsCleared", FilterOperator.EQ, iIsCleared)];

      var sTrimmed = sQuery.trim();
      if (sTrimmed) {
        var aSearchFields = [
          "Model", "ModelDescription", "SoNumberLineItem", "ContainerNumber",
          "ProformaInvoiceNumber","BillOfLading", "CommercialInvoicePl", "PurchaseOrderNumber",
          "PalletNumber", "GrossWeight", "PurchaseOrderItem", "Material",
          "VendorPartNumber", "DeliveryDate", "DeliveryTime", "Plant",
          "BatchNumber", "FreeGoodsIndicator", "NoteOfFreeGoods",
          "PoRefForFreeGoods", "PoLineItemRef", "BillOfPayment"
        ];

        var aSearchFilters = aSearchFields.map(function(sField) {
          return new Filter(sField, FilterOperator.Contains, sTrimmed);
        });

        aFilters.push(new Filter({ filters: aSearchFilters, and: false }));
      }

      var oCombined = aFilters.length === 1
        ? aFilters[0]
        : new Filter({ filters: aFilters, and: true });

      oBinding.filter(oCombined);
      oBinding.sort(new Sorter("LastChangedAt", true));
    },

    onNavToCustomPage: function() {
      this.getOwnerComponent().getRouter().navTo("CustomPage");
    }
  });
});