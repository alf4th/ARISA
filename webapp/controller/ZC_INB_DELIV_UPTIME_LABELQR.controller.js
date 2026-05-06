sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageBox",
    "sap/m/MessageToast",
    "sap/ui/core/Fragment",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/ui/core/ListItem"
], function (Controller, JSONModel, MessageBox, MessageToast, Fragment, Filter, FilterOperator, ListItem) {
    "use strict";

    // ---------------------------------------------------------------------------
    // Initial view model state
    // ---------------------------------------------------------------------------
    var INITIAL_STATE = {
        extDeliveryId: "",
        deliveryFound: false,

        // Readonly fields from backend
        inboundDelivery: "",
        supplier: "",
        vendorName: "",
        plant: "",
        operatorName: "",
        shift: "",

        // Read-only fields auto-filled by backend
        unloadQueue: "",

        // Custom fields (editable) — order matches projection
        emkl: "",
        emklName: "",
        noPolisiTruck: "",
        namaDriver: "",
        gateOutPortDate: null,
        gateOutPortTime: null,
        dateCheckInArisa: null,
        checkInTime: null,
        dateUnload: null,
        unloadTime: null,
        unloadFinishedDate: null,
        unloadFinishedTime: null,
        checkOutDate: null,
        checkOutTime: null,

        // UI state
        customFieldsEnabled: false,
        printPreviewEnabled: false,
        isBusy: false,
        headerExpanded: true
    };

    return Controller.extend("zdelivuptime.delivuptimelabelqr.controller.ZC_INB_DELIV_UPTIME_LABELQR", {

        // =========================================================================
        // LIFECYCLE
        // =========================================================================

        onInit: function () {
            // Restore persisted state if user is returning from cross-app navigation
            var oRestoredState = this._restoreSessionState();
            var oInitialData   = oRestoredState || Object.assign({}, INITIAL_STATE);

            var oViewModel = new JSONModel(oInitialData);
            this.getView().setModel(oViewModel, "viewModel");

            // Fragment references
            this._oScanDialog       = null;
            this._oPreviewDialog    = null;
            this._oSupplierVHDialog = null;

            // Camera stream reference
            this._oCameraStream = null;

            // Debounce timer
            this._oSearchTimer = null;

            // OData V4 context references
            this._oDeliveryContext  = null;
            this._aDeliveryContexts = [];

            // If state was restored, re-fetch contexts silently so save still works
            if (oRestoredState && oRestoredState.extDeliveryId) {
                this._searchDelivery(oRestoredState.extDeliveryId);
            }
        },

        // =========================================================================
        // INPUT: External Delivery ID
        // =========================================================================

        /**
         * Debounced live change: triggers search 600ms after user stops typing.
         */
        onExtDeliveryIdChange: function (oEvent) {
            var sValue = oEvent.getParameter("value").trim();

            if (this._oSearchTimer) {
                clearTimeout(this._oSearchTimer);
                this._oSearchTimer = null;
            }

            if (!sValue) {
                this._resetToInitialState();
                return;
            }

            this._oSearchTimer = setTimeout(function () {
                this._searchDelivery(sValue);
            }.bind(this), 600);
        },

        /**
         * Immediate search on Enter key.
         */
        onExtDeliveryIdSubmit: function () {
            var sValue = this.getView().getModel("viewModel").getProperty("/extDeliveryId").trim();
            if (this._oSearchTimer) {
                clearTimeout(this._oSearchTimer);
                this._oSearchTimer = null;
            }
            if (sValue) {
                this._searchDelivery(sValue);
            }
        },

        // =========================================================================
        // SCAN BUTTON
        // =========================================================================

        onScanPress: function () {
            this._openScanDialog();
        },

        onToggleHeader: function () {
            var oViewModel = this.getView().getModel("viewModel");
            oViewModel.setProperty("/headerExpanded", !oViewModel.getProperty("/headerExpanded"));
        },

        onTabSelect: function (oEvent) {
            var sKey = oEvent.getParameter("key");
            var sTargetId = sKey === "general" ? "vboxGeneralInfoSection" : "vboxDetailInfoSection";
            var oTarget = this.byId(sTargetId);
            var oPage = this.byId("page");
            if (oTarget && oPage) {
                oPage.scrollToElement(oTarget, 500);
            }
        },

        onTimestampGateOutPress: function () {
            this._setTimestamp("/gateOutPortDate", "/gateOutPortTime");
        },

        onTimestampCheckInPress: function () {
            this._setTimestamp("/dateCheckInArisa", "/checkInTime");
        },

        onTimestampUnloadPress: function () {
            this._setTimestamp("/dateUnload", "/unloadTime");
        },

        onTimestampUnloadFinishedPress: function () {
            this._setTimestamp("/unloadFinishedDate", "/unloadFinishedTime");
        },

        onTimestampCheckOutPress: function () {
            this._setTimestamp("/checkOutDate", "/checkOutTime");
        },

        _setTimestamp: function (sDatePath, sTimePath) {
            var oNow = new Date();
            var sDate = oNow.getFullYear() + "-" +
                        String(oNow.getMonth() + 1).padStart(2, "0") + "-" +
                        String(oNow.getDate()).padStart(2, "0");
            var sTime = String(oNow.getHours()).padStart(2, "0") + ":" +
                        String(oNow.getMinutes()).padStart(2, "0") + ":" +
                        String(oNow.getSeconds()).padStart(2, "0");
            var oViewModel = this.getView().getModel("viewModel");
            oViewModel.setProperty(sDatePath, sDate);
            oViewModel.setProperty(sTimePath, sTime);
        },

        onNoPolisiTruckChange: function (oEvent) {
            var sUpper = oEvent.getSource().getValue().toUpperCase();
            oEvent.getSource().setValue(sUpper);
            this.getView().getModel("viewModel").setProperty("/noPolisiTruck", sUpper);
        },

        // =========================================================================
        // EMKL VALUE HELP
        // =========================================================================

        onEMKLValueHelp: function () {
            var oView = this.getView();
            if (!this._oSupplierVHDialog) {
                Fragment.load({
                    id: oView.getId(),
                    name: "zdelivuptime.delivuptimelabelqr.view.SupplierVH",
                    controller: this
                }).then(function (oDialog) {
                    this._oSupplierVHDialog = oDialog;
                    oView.addDependent(oDialog);
                    this._loadSupplierData("", oDialog);
                    oDialog.open();
                }.bind(this)).catch(function (oErr) {
                    MessageBox.error("Failed to open supplier value help: " + oErr.message);
                });
            } else {
                this._loadSupplierData("", this._oSupplierVHDialog);
                this._oSupplierVHDialog.open();
            }
        },

        _loadSupplierData: function (sQuery, oDialog) {
            var sBase = "/sap/opu/odata4/sap/zsb_inb_deliv_uptime_labelqr_4/srvd/sap/zsd_inb_deliv_uptime_labelqr/0001/I_Supplier_VH";
            var aParams = ["$select=Supplier,SupplierName", "$top=50", "$orderby=Supplier"];
            if (sQuery) {
                var sQ = sQuery.replace(/'/g, "''");
                aParams.push("$filter=contains(Supplier,'" + sQ + "') or contains(SupplierName,'" + sQ + "')");
            }
            var sUrl = sBase + "?" + aParams.join("&");

            oDialog.setBusy(true);
            fetch(sUrl, { credentials: "same-origin", headers: { Accept: "application/json" } })
                .then(function (oRes) {
                    if (!oRes.ok) { throw new Error("HTTP " + oRes.status); }
                    return oRes.json();
                })
                .then(function (oData) {
                    var aItems = (oData.value || []).map(function (o) {
                        return { Supplier: o.Supplier, SupplierName: o.SupplierName || "" };
                    });
                    oDialog.setModel(new JSONModel({ items: aItems }), "supplierVH");
                    oDialog.setBusy(false);
                })
                .catch(function (oErr) {
                    oDialog.setBusy(false);
                    MessageBox.error("Could not load supplier data: " + oErr.message);
                });
        },

        onSupplierVHSearch: function (oEvent) {
            var sQuery = oEvent.getParameter("value");
            this._loadSupplierData(sQuery || "", this._oSupplierVHDialog);
        },

        onEMKLSuggest: function (oEvent) {
            var sQuery = oEvent.getParameter("suggestValue") || "";
            if (!sQuery) { return; }
            var oInput = oEvent.getSource();
            var sBase = "/sap/opu/odata4/sap/zsb_inb_deliv_uptime_labelqr_4/srvd/sap/zsd_inb_deliv_uptime_labelqr/0001/I_Supplier_VH";
            var sQ = sQuery.replace(/'/g, "''");
            var sUrl = sBase + "?$select=Supplier,SupplierName&$top=10" +
                       "&$filter=contains(Supplier,'" + sQ + "') or contains(SupplierName,'" + sQ + "')";
            fetch(sUrl, { credentials: "same-origin", headers: { Accept: "application/json" } })
                .then(function (oRes) {
                    if (!oRes.ok) { throw new Error("HTTP " + oRes.status); }
                    return oRes.json();
                })
                .then(function (oData) {
                    oInput.destroySuggestionItems();
                    (oData.value || []).forEach(function (o) {
                        oInput.addSuggestionItem(new ListItem({
                            key: o.Supplier,
                            text: o.Supplier,
                            additionalText: o.SupplierName || ""
                        }));
                    });
                })
                .catch(function () {});
        },

        onEMKLSuggestionSelected: function (oEvent) {
            var oItem = oEvent.getParameter("selectedItem");
            if (!oItem) { return; }
            var oViewModel = this.getView().getModel("viewModel");
            oViewModel.setProperty("/emkl",     oItem.getKey());
            oViewModel.setProperty("/emklName", oItem.getAdditionalText());
        },

        onSupplierVHConfirm: function (oEvent) {
            var oSelected = oEvent.getParameter("selectedItem");
            if (oSelected) {
                var oCtx = oSelected.getBindingContext("supplierVH");
                var oViewModel = this.getView().getModel("viewModel");
                oViewModel.setProperty("/emkl",     oCtx.getProperty("Supplier"));
                oViewModel.setProperty("/emklName", oCtx.getProperty("SupplierName"));
            }
        },

        onSupplierVHClose: function () {
            if (this._oSupplierVHDialog) {
                this._oSupplierVHDialog.close();
            }
        },

        // =========================================================================
        // FOOTER BUTTONS
        // =========================================================================

        onSavePress: function () {
            var oViewModel = this.getView().getModel("viewModel");

            if (!oViewModel.getProperty("/deliveryFound") || !this._oDeliveryContext) {
                MessageBox.warning("Please enter a valid External Delivery ID first.");
                return;
            }

            if (!this._validateBeforeSave(oViewModel.getData())) {
                return;
            }

            this._saveToBackend();
        },

        onCancelPress: function () {
            var oViewModel = this.getView().getModel("viewModel");
            if (!oViewModel.getProperty("/deliveryFound")) {
                this._resetToInitialState();
                return;
            }

            MessageBox.confirm("Are you sure you want to cancel? All unsaved changes will be lost.", {
                onClose: function (sAction) {
                    if (sAction === MessageBox.Action.OK) {
                        this._resetToInitialState();
                    }
                }.bind(this)
            });
        },

        onPrintPreviewPress: function () {
            var oViewModel  = this.getView().getModel("viewModel");

            if (!oViewModel.getProperty("/printPreviewEnabled")) {
                MessageBox.warning("Print preview is only available after data has been saved.");
                return;
            }

            var sContainerNumber = oViewModel.getProperty("/extDeliveryId") || "";

            // Persist form state so it can be restored when the user navigates back
            this._saveSessionState(oViewModel.getData());

            // Cross-app navigation to InboundDeliveryLabel app
            var oCrossNav = sap.ushell &&
                            sap.ushell.Container &&
                            sap.ushell.Container.getService("CrossApplicationNavigation");

            if (oCrossNav) {
                oCrossNav.toExternal({
                    target: {
                        semanticObject: "InboundDeliveryLabel",
                        action: "display"
                    },
                    appSpecificRoute: "/detail/" + encodeURIComponent(sContainerNumber)
                });
            } else {
                // Fallback: direct hash navigation (local dev / non-FLP)
                window.location.hash = "InboundDeliveryLabel-display&/detail/" + encodeURIComponent(sContainerNumber);
            }
        },

        // =========================================================================
        // PRIVATE: Backend Search (OData V4)
        // =========================================================================

        _searchDelivery: function (sValue) {
            var oViewModel = this.getView().getModel("viewModel");
            var oModel = this.getOwnerComponent().getModel();

            oViewModel.setProperty("/isBusy", true);
            oViewModel.setProperty("/deliveryFound", false);
            this._oDeliveryContext = null;

            var oListBinding = oModel.bindList("/ZC_INB_DELIV_UPTIME_LABELQR", null, [],
                [new Filter("ExtDeliveryId", FilterOperator.EQ, sValue)]
            );

            // Fetch all records (multiple inbound deliveries per container number)
            oListBinding.requestContexts(0, 100).then(function (aContexts) {
                if (!aContexts || aContexts.length === 0) {
                    oViewModel.setProperty("/isBusy", false);
                    MessageToast.show("Container Number '" + sValue + "' not found.");
                    return Promise.resolve(null);
                }

                // Keep ALL contexts — each IBD will be PATCHed on save
                this._oDeliveryContext  = aContexts[0];
                this._aDeliveryContexts = aContexts;

                // Request all objects to concatenate InboundDelivery
                return Promise.all(aContexts.map(function (oCtx) {
                    return oCtx.requestObject();
                }));

            }.bind(this)).then(function (aDataArray) {
                if (!aDataArray) return;

                var oFirst = aDataArray[0];

                // Concatenate all InboundDelivery values separated by ", "
                var sInboundDeliveries = aDataArray.map(function (d) {
                    return d.InboundDelivery || "";
                }).filter(Boolean).join(", ");

                // Populate readonly display fields (from first record)
                oViewModel.setProperty("/inboundDelivery", sInboundDeliveries);
                oViewModel.setProperty("/supplier",        oFirst.supplier        || "");
                oViewModel.setProperty("/vendorName",      oFirst.Vendorname      || "");
                oViewModel.setProperty("/plant",           oFirst.Plant           || "");
                oViewModel.setProperty("/operatorName",    oFirst.OperatorName    || "");
                var shiftRaw = (oFirst.Shift !== null && oFirst.Shift !== undefined) ? String(oFirst.Shift).trim() : "";
                var shiftNum = parseFloat(shiftRaw);
                // Convert "1.00" → "1", "2.00" → "2", etc. to match Select key; zero → ""
                var shiftKey = (!isNaN(shiftNum) && shiftNum > 0) ? String(Math.round(shiftNum)) : "";
                oViewModel.setProperty("/shift", shiftKey);

                // Populate read-only auto fields
                oViewModel.setProperty("/unloadQueue",       oFirst.UnloadQueue        || "");

                // Populate custom fields (from first record)
                oViewModel.setProperty("/emkl",              oFirst.EMKL               || "");
                oViewModel.setProperty("/emklName",          oFirst.EMKLName            || "");
                oViewModel.setProperty("/noPolisiTruck",     oFirst.NoPolisiTruck      || "");
                oViewModel.setProperty("/namaDriver",        oFirst.NamaDriver         || "");
                oViewModel.setProperty("/gateOutPortDate",  oFirst.GateOutPortDate    || null);
                oViewModel.setProperty("/gateOutPortTime",  oFirst.GateOutPortTime    || null);
                oViewModel.setProperty("/dateCheckInArisa", oFirst.DateCheckInArisa   || null);
                oViewModel.setProperty("/checkInTime",      oFirst.CheckInTime        || null);
                oViewModel.setProperty("/dateUnload",       oFirst.DateUnload         || null);
                oViewModel.setProperty("/unloadTime",       oFirst.UnloadTime         || null);
                oViewModel.setProperty("/unloadFinishedDate", oFirst.UnloadFinishedDate || null);
                oViewModel.setProperty("/unloadFinishedTime", oFirst.UnloadFinishedTime || null);
                oViewModel.setProperty("/checkOutDate",     oFirst.CheckOutDate       || null);
                oViewModel.setProperty("/checkOutTime",     oFirst.CheckOutTime       || null);

                oViewModel.setProperty("/deliveryFound", true);
                oViewModel.setProperty("/isBusy", false);

                // Treat null, undefined, "", "00:00:00", zero numbers/strings as not filled
                var fnHasValue = function (v) {
                    if (v === null || v === undefined || v === "") return false;
                    var s = String(v).trim();
                    if (s === "" || s === "00:00:00") return false;
                    // Catch numeric zero variants: "0", "0.0", "0.00", etc.
                    var n = parseFloat(s);
                    if (!isNaN(n) && n === 0) return false;
                    return true;
                };
                var bCustomFieldsFilled = (
                    fnHasValue(oFirst.OperatorName) ||
                    fnHasValue(oFirst.Shift) ||
                    fnHasValue(oFirst.EMKL) ||
                    fnHasValue(oFirst.NoPolisiTruck) ||
                    fnHasValue(oFirst.NamaDriver) ||
                    fnHasValue(oFirst.GateOutPortDate) ||
                    fnHasValue(oFirst.GateOutPortTime) ||
                    fnHasValue(oFirst.DateCheckInArisa) ||
                    fnHasValue(oFirst.CheckInTime) ||
                    fnHasValue(oFirst.DateUnload) ||
                    fnHasValue(oFirst.UnloadTime) ||
                    fnHasValue(oFirst.UnloadFinishedDate) ||
                    fnHasValue(oFirst.UnloadFinishedTime) ||
                    fnHasValue(oFirst.CheckOutDate) ||
                    fnHasValue(oFirst.CheckOutTime)
                );

                if (bCustomFieldsFilled) {
                    oViewModel.setProperty("/printPreviewEnabled", true);
                    oViewModel.setProperty("/customFieldsEnabled", false);
                    this._showOverwriteConfirmation();
                } else {
                    oViewModel.setProperty("/printPreviewEnabled", false);
                    oViewModel.setProperty("/customFieldsEnabled", true);
                }

            }.bind(this)).catch(function (oError) {
                oViewModel.setProperty("/isBusy", false);
                var sMsg = (oError && oError.message) ? oError.message : "Error fetching delivery data.";
                MessageBox.error(sMsg);
            }.bind(this));
        },

        // =========================================================================
        // PRIVATE: Overwrite Confirmation
        // =========================================================================

        _showOverwriteConfirmation: function () {
            var oViewModel = this.getView().getModel("viewModel");

            MessageBox.confirm(
                "Custom fields already have data for this delivery.\n\nDo you want to overwrite the existing data?",
                {
                    title: "Overwrite Existing Data?",
                    actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                    emphasizedAction: MessageBox.Action.NO,
                    onClose: function (sAction) {
                        if (sAction === MessageBox.Action.YES) {
                            // Allow editing
                            oViewModel.setProperty("/customFieldsEnabled", true);
                        } else {
                            // Keep fields grey / readonly
                            oViewModel.setProperty("/customFieldsEnabled", false);
                        }
                    }
                }
            );
        },

        // =========================================================================
        // PRIVATE: Client-side validation before save
        // =========================================================================

        _validateBeforeSave: function (oData) {
            // --- Mandatory fields ---
            if (!oData.operatorName || !String(oData.operatorName).trim()) {
                MessageBox.error("Operator Name is mandatory.");
                return false;
            }
            if (!oData.shift || !String(oData.shift).trim()) {
                MessageBox.error("Shift is mandatory.");
                return false;
            }

            // --- Sequence validation ---
            // Maps to I_DELIVERYDOCUMENT extension fields:
            //   Step 1: YY1_NoPolisiTruck & YY1_NamaDriver
            //   Step 2: YY1_TglKeluarPelabuhan & YY1_JamKeluarPelabuhan & YY1_EMKL
            //   Step 3: YY1_TglMasukArisa & YY1_JamMasukArisa
            //   Step 4: YY1_TglBongkar & YY1_JamBongkar
            //   Step 5: YY1_SelesaiBongkar & YY1_JamSelesaiBongkar
            var fnFilled = function (v) {
                if (v === null || v === undefined) return false;
                var s = String(v).trim();
                return s !== "" && s !== "00:00:00";
            };

            // Step 1 is the mandatory entry point — cannot save without it
            if (!fnFilled(oData.noPolisiTruck) || !fnFilled(oData.namaDriver) || !fnFilled(oData.emkl)) {
                MessageBox.error("Please fill in the available columns in order.");
                return false;
            }

            var aSteps = [
                {
                    allFilled: fnFilled(oData.noPolisiTruck) && fnFilled(oData.namaDriver) && fnFilled(oData.emkl),
                    anyFilled: fnFilled(oData.noPolisiTruck) || fnFilled(oData.namaDriver) || fnFilled(oData.emkl)
                },
                {
                    allFilled: fnFilled(oData.gateOutPortDate) && fnFilled(oData.gateOutPortTime),
                    anyFilled: fnFilled(oData.gateOutPortDate) || fnFilled(oData.gateOutPortTime)
                },
                {
                    allFilled: fnFilled(oData.dateCheckInArisa) && fnFilled(oData.checkInTime),
                    anyFilled: fnFilled(oData.dateCheckInArisa) || fnFilled(oData.checkInTime)
                },
                {
                    allFilled: fnFilled(oData.dateUnload) && fnFilled(oData.unloadTime),
                    anyFilled: fnFilled(oData.dateUnload) || fnFilled(oData.unloadTime)
                },
                {
                    allFilled: fnFilled(oData.unloadFinishedDate) && fnFilled(oData.unloadFinishedTime),
                    anyFilled: fnFilled(oData.unloadFinishedDate) || fnFilled(oData.unloadFinishedTime)
                },
                {
                    allFilled: fnFilled(oData.checkOutDate) && fnFilled(oData.checkOutTime),
                    anyFilled: fnFilled(oData.checkOutDate) || fnFilled(oData.checkOutTime)
                }
            ];

            var bFoundIncomplete = false;
            for (var i = 0; i < aSteps.length; i++) {
                var oStep = aSteps[i];
                if (oStep.anyFilled && !oStep.allFilled) {
                    // Partial fill within a step
                    MessageBox.error("Please fill in the available columns in order.");
                    return false;
                }
                if (bFoundIncomplete && oStep.anyFilled) {
                    // A later step has data but a prior step was empty
                    MessageBox.error("Please fill in the available columns in order.");
                    return false;
                }
                if (!oStep.allFilled) {
                    bFoundIncomplete = true;
                }
            }

            return true;
        },

        // =========================================================================
        // PRIVATE: Re-fetch backend-computed fields after save
        // =========================================================================

        _refreshComputedFields: function () {
            var oViewModel = this.getView().getModel("viewModel");
            var oModel     = this.getOwnerComponent().getModel();
            var sId        = oViewModel.getProperty("/extDeliveryId");

            var oListBinding = oModel.bindList("/ZC_INB_DELIV_UPTIME_LABELQR", null, [],
                [new Filter("ExtDeliveryId", FilterOperator.EQ, sId)]
            );

            return oListBinding.requestContexts(0, 1).then(function (aContexts) {
                if (!aContexts || aContexts.length === 0) return;
                return aContexts[0].requestObject();
            }).then(function (oData) {
                if (oData) {
                    oViewModel.setProperty("/unloadQueue", oData.UnloadQueue || "");
                }
            }).catch(function () { /* non-critical — ignore refresh errors */ });
        },

        // =========================================================================
        // PRIVATE: Save to Backend (OData V4 PATCH)
        // =========================================================================

        _saveToBackend: function () {
            var oViewModel = this.getView().getModel("viewModel");
            var oModel     = this.getOwnerComponent().getModel();
            var aContexts  = this._aDeliveryContexts;
            var oData      = oViewModel.getData();

            if (!aContexts || aContexts.length === 0) return;

            oViewModel.setProperty("/isBusy", true);

            // Value helpers
            var fnHasStr  = function (v) { return v !== null && v !== undefined && String(v).trim() !== ""; };
            var fnHasDate = function (v) { return v !== null && v !== undefined && String(v).trim() !== ""; };
            var fnHasTime = function (v) {
                if (!v) return false;
                var s = String(v).trim();
                return s !== "" && s !== "00:00:00";
            };

            // Build a helper that patches ONE context and waits for its response.
            // Each IBD uses a unique batch group so they are sent sequentially,
            // not merged — required because backend BO interface handles 1 update per event.
            var fnPatchOne = function (oCtx, iIdx) {
                var sBatchGroup = "updateDelivery_" + iIdx;
                var aPromises   = [];

                aPromises.push(oCtx.setProperty("OperatorName", fnHasStr(oData.operatorName) ? String(oData.operatorName) : "", sBatchGroup));
                aPromises.push(oCtx.setProperty("Shift",        fnHasStr(oData.shift)        ? String(oData.shift)        : "", sBatchGroup));
                aPromises.push(oCtx.setProperty("EMKL",         fnHasStr(oData.emkl)         ? String(oData.emkl)         : "", sBatchGroup));
                aPromises.push(oCtx.setProperty("NoPolisiTruck", fnHasStr(oData.noPolisiTruck) ? String(oData.noPolisiTruck) : "", sBatchGroup));
                aPromises.push(oCtx.setProperty("NamaDriver",    fnHasStr(oData.namaDriver)    ? String(oData.namaDriver)    : "", sBatchGroup));

                if (fnHasDate(oData.gateOutPortDate))    { aPromises.push(oCtx.setProperty("GateOutPortDate",    oData.gateOutPortDate,    sBatchGroup)); }
                if (fnHasTime(oData.gateOutPortTime))    { aPromises.push(oCtx.setProperty("GateOutPortTime",    oData.gateOutPortTime,    sBatchGroup)); }
                if (fnHasDate(oData.dateCheckInArisa))   { aPromises.push(oCtx.setProperty("DateCheckInArisa",   oData.dateCheckInArisa,   sBatchGroup)); }
                if (fnHasTime(oData.checkInTime))        { aPromises.push(oCtx.setProperty("CheckInTime",        oData.checkInTime,        sBatchGroup)); }
                if (fnHasDate(oData.dateUnload))         { aPromises.push(oCtx.setProperty("DateUnload",         oData.dateUnload,         sBatchGroup)); }
                if (fnHasTime(oData.unloadTime))         { aPromises.push(oCtx.setProperty("UnloadTime",         oData.unloadTime,         sBatchGroup)); }
                if (fnHasDate(oData.unloadFinishedDate)) { aPromises.push(oCtx.setProperty("UnloadFinishedDate", oData.unloadFinishedDate, sBatchGroup)); }
                if (fnHasTime(oData.unloadFinishedTime)) { aPromises.push(oCtx.setProperty("UnloadFinishedTime", oData.unloadFinishedTime, sBatchGroup)); }
                if (fnHasDate(oData.checkOutDate))       { aPromises.push(oCtx.setProperty("CheckOutDate",       oData.checkOutDate,       sBatchGroup)); }
                if (fnHasTime(oData.checkOutTime))       { aPromises.push(oCtx.setProperty("CheckOutTime",       oData.checkOutTime,       sBatchGroup)); }

                // Trigger PATCH for this IBD, then wait for all its promises to resolve
                oModel.submitBatch(sBatchGroup);
                return Promise.all(aPromises);
            };

            // Sequential chain: IBD[0] → wait → IBD[1] → wait → IBD[2] → ...
            var pChain = Promise.resolve();
            aContexts.forEach(function (oCtx, iIdx) {
                pChain = pChain.then(function () {
                    return fnPatchOne(oCtx, iIdx);
                });
            });

            pChain.then(function () {
                // Re-fetch to get backend-computed fields (e.g. UnloadQueue set by before-modify)
                return this._refreshComputedFields();
            }.bind(this)).then(function () {
                oViewModel.setProperty("/isBusy", false);
                oViewModel.setProperty("/printPreviewEnabled", true);
                oViewModel.setProperty("/customFieldsEnabled", false);
                MessageBox.success("Data saved successfully.", { title: "Success" });
            }).catch(function (oError) {
                oViewModel.setProperty("/isBusy", false);
                var sMsg = (oError && oError.message) ? oError.message : "Save failed. Please try again.";
                MessageBox.error(sMsg);
            });
        },

        // =========================================================================
        // SCAN DIALOG
        // =========================================================================

        onScanDialogClose: function () {
            this._stopCameraStream();
        },

        onScanCancel: function () {
            this._stopCameraStream();
            if (this._oScanDialog) {
                this._oScanDialog.close();
            }
        },

        onShowManualInput: function () {
            var oView = this.getView();
            oView.byId("cameraPreviewBox").setVisible(false);
            oView.byId("manualInputBox").setVisible(true);
            oView.byId("scanStatusText").setText("Enter barcode manually and press Enter.");
            this._stopCameraStream();
        },

        onManualBarcodeSubmit: function (oEvent) {
            var sValue = oEvent.getSource().getValue().trim();
            if (sValue) {
                this._applyScannedValue(sValue);
            }
        },

        onCaptureScan: function () {
            var result = this._decodeFrameWithJsQr();
            if (result) {
                this._applyScannedValue(result);
            } else {
                this.getView().byId("scanStatusText").setText("No barcode detected. Try again or use manual input.");
            }
        },

        _openScanDialog: function () {
            var oView = this.getView();

            // Try native SAP BarcodeScanner first (Fiori Client / mobile)
            if (sap.ndc && sap.ndc.BarcodeScanner) {
                sap.ndc.BarcodeScanner.scan(
                    function (mResult) {
                        if (!mResult.cancelled && mResult.text) {
                            this._applyScannedValue(mResult.text);
                        }
                    }.bind(this),
                    function (sError) {
                        MessageBox.error("Barcode scan failed: " + sError);
                    }
                );
                return;
            }

            if (!this._oScanDialog) {
                Fragment.load({
                    id: oView.getId(),
                    name: "zdelivuptime.delivuptimelabelqr.view.ScanDialog",
                    controller: this
                }).then(function (oDialog) {
                    this._oScanDialog = oDialog;
                    oView.addDependent(oDialog);
                    oDialog.open();
                    this._startCameraStream();
                }.bind(this)).catch(function (oErr) {
                    MessageBox.error("Failed to open Scan dialog: " + oErr.message);
                });
            } else {
                oView.byId("cameraPreviewBox").setVisible(true);
                oView.byId("manualInputBox").setVisible(false);
                oView.byId("manualBarcodeInput").setValue("");
                var oContainer = document.getElementById("scanCameraContainer");
                if (oContainer) { oContainer.innerHTML = ""; }
                this._oScanDialog.open();
                this._startCameraStream();
            }
        },

        _startCameraStream: function () {
            var oView = this.getView();
            var oStatusText = oView.byId("scanStatusText");
            var oCameraBox   = oView.byId("cameraPreviewBox");
            var oManualBox   = oView.byId("manualInputBox");

            if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
                oStatusText.setText("Camera not supported. Use manual input.");
                oCameraBox.setVisible(false);
                oManualBox.setVisible(true);
                return;
            }

            var oContainer = document.getElementById("scanCameraContainer");
            if (oContainer && !oContainer.querySelector("video")) {
                oContainer.innerHTML =
                    "<video id='scanVideo' style='width:320px;height:240px;border:1px solid #ccc;' autoplay playsinline></video>" +
                    "<canvas id='scanCanvas' style='display:none'></canvas>";
            }

            oCameraBox.setVisible(true);
            oManualBox.setVisible(false);
            oStatusText.setText("Camera starting... point at barcode to scan automatically.");

            navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } })
                .then(function (stream) {
                    this._oCameraStream = stream;
                    var oVideo = document.getElementById("scanVideo");
                    if (oVideo) {
                        oVideo.srcObject = stream;
                        this._startQrPolling();
                    }
                }.bind(this))
                .catch(function (err) {
                    oStatusText.setText("Camera error: " + err.message + ". Use manual input.");
                    oCameraBox.setVisible(false);
                    oManualBox.setVisible(true);
                }.bind(this));
        },

        _startQrPolling: function () {
            var that = this;
            this._loadJsQr(function () {
                that._oQrPollInterval = setInterval(function () {
                    var result = that._decodeFrameWithJsQr();
                    if (result) {
                        clearInterval(that._oQrPollInterval);
                        that._applyScannedValue(result);
                    }
                }, 500);
            });
        },

        _decodeFrameWithJsQr: function () {
            var oVideo  = document.getElementById("scanVideo");
            var oCanvas = document.getElementById("scanCanvas");
            if (!oVideo || !oCanvas || !window.jsQR) return null;

            oCanvas.width  = oVideo.videoWidth;
            oCanvas.height = oVideo.videoHeight;
            var oCtx = oCanvas.getContext("2d");
            oCtx.drawImage(oVideo, 0, 0, oCanvas.width, oCanvas.height);
            var imageData = oCtx.getImageData(0, 0, oCanvas.width, oCanvas.height);
            var code = window.jsQR(imageData.data, imageData.width, imageData.height);
            return code ? code.data : null;
        },

        _loadJsQr: function (fnCallback) {
            if (window.jsQR) { fnCallback(); return; }
            var oScript = document.createElement("script");
            oScript.src = "https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.js";
            oScript.onload = fnCallback;
            document.head.appendChild(oScript);
        },

        _stopCameraStream: function () {
            if (this._oQrPollInterval) {
                clearInterval(this._oQrPollInterval);
                this._oQrPollInterval = null;
            }
            if (this._oCameraStream) {
                this._oCameraStream.getTracks().forEach(function (track) { track.stop(); });
                this._oCameraStream = null;
            }
        },

        _applyScannedValue: function (sValue) {
            this._stopCameraStream();
            if (this._oScanDialog) {
                this._oScanDialog.close();
            }
            var oViewModel = this.getView().getModel("viewModel");
            oViewModel.setProperty("/extDeliveryId", sValue);
            this._searchDelivery(sValue);
        },

        // =========================================================================
        // PRINT PREVIEW DIALOG
        // =========================================================================

        onPreviewClose: function () {
            if (this._oPreviewDialog) {
                this._oPreviewDialog.close();
            }
        },

        onPrintPdf: function () {
            var oViewModel = this.getView().getModel("viewModel");
            this._generateAndDownloadPdf(oViewModel.getData());
        },

        _openPreviewDialog: function () {
            var oView = this.getView();
            var oData = this.getView().getModel("viewModel").getData();

            if (!this._oPreviewDialog) {
                Fragment.load({
                    id: oView.getId(),
                    name: "zdelivuptime.delivuptimelabelqr.view.PrintPreviewDialog",
                    controller: this
                }).then(function (oDialog) {
                    this._oPreviewDialog = oDialog;
                    oView.addDependent(oDialog);
                    oDialog.open();
                    this._renderLabelPreview(oData);
                }.bind(this)).catch(function (oErr) {
                    MessageBox.error("Failed to open Preview dialog: " + oErr.message);
                });
            } else {
                this._oPreviewDialog.open();
                this._renderLabelPreview(oData);
            }
        },

        /**
         * Renders the delivery label HTML preview.
         * Layout: 2x2 table — ExtDeliveryId | QR Code / Plant | Remarks(empty)
         * Remark field is intentionally left empty per spec.
         */
        _renderLabelPreview: function (oData) {
            var that = this;
            var sExtDelivId = oData.extDeliveryId || "-";
            var sPlant      = oData.plant         || "-";

            var sBorderOuter = "2px solid #222";
            var sBorderInner = "2px solid #444";
            var sTdBase      = "border:" + sBorderInner + ";padding:14px 16px;vertical-align:top;";

            var sHtml =
                "<div style='border:" + sBorderOuter + ";font-family:Arial,sans-serif;background:#fff;max-width:540px;'>" +
                    "<table style='width:100%;border-collapse:collapse;'>" +
                        "<tr>" +
                            "<td style='" + sTdBase + "width:58%;'>" +
                                "<div style='font-size:10px;color:#555;margin-bottom:8px;'>Container ID</div>" +
                                "<div style='font-size:42px;font-weight:bold;letter-spacing:1px;word-break:break-all;line-height:1.1;'>" + sExtDelivId + "</div>" +
                            "</td>" +
                            "<td style='" + sTdBase + "width:42%;text-align:right;'>" +
                                "<div style='font-size:10px;color:#555;margin-bottom:8px;'>Barcode:</div>" +
                                "<div id='delivLabelQr'><em style='color:#aaa;font-size:11px;'>Loading QR...</em></div>" +
                            "</td>" +
                        "</tr>" +
                        "<tr>" +
                            "<td style='" + sTdBase + "'>" +
                                "<div style='font-size:10px;color:#555;margin-bottom:8px;'>Plant:</div>" +
                                "<div style='font-size:42px;font-weight:bold;line-height:1.1;'>" + sPlant + "</div>" +
                            "</td>" +
                            "<td style='" + sTdBase + "'>" +
                                "<div style='font-size:10px;color:#555;margin-bottom:8px;'>Remarks:</div>" +
                            "</td>" +
                        "</tr>" +
                    "</table>" +
                "</div>";

            var oHtmlControl = this.getView().byId("labelPreviewHtml");
            oHtmlControl.setContent(sHtml);

            // Render QR code after DOM update
            setTimeout(function () {
                that._loadQrCodeJs(function () {
                    var oQrTarget = document.getElementById("delivLabelQr");
                    if (oQrTarget && window.QRCode) {
                        oQrTarget.innerHTML = "";
                        new window.QRCode(oQrTarget, {
                            text: sExtDelivId,
                            width: 110,
                            height: 110,
                            correctLevel: window.QRCode.CorrectLevel.M
                        });
                    }
                });
            }, 200);
        },

        /**
         * Generates and downloads a PDF of the delivery label.
         * Remark field is intentionally empty per spec.
         */
        _generateAndDownloadPdf: function (oData) {
            this._loadPdfMake(function () {
                var sExtDelivId = oData.extDeliveryId || "-";
                var sPlant      = oData.plant         || "-";

                var oDocDef = {
                    pageSize: "A5",
                    pageOrientation: "landscape",
                    content: [
                        {
                            table: {
                                widths: ["58%", "42%"],
                                body: [
                                    [
                                        {
                                            stack: [
                                                { text: "Container ID", fontSize: 8, color: "#555", margin: [0, 0, 0, 6] },
                                                { text: sExtDelivId, fontSize: 36, bold: true }
                                            ],
                                            margin: [10, 10, 10, 10]
                                        },
                                        {
                                            stack: [
                                                { text: "Barcode:", fontSize: 8, color: "#555", margin: [0, 0, 0, 6], alignment: "right" },
                                                { qr: sExtDelivId, fit: 110, alignment: "right" }
                                            ],
                                            margin: [10, 10, 10, 10]
                                        }
                                    ],
                                    [
                                        {
                                            stack: [
                                                { text: "Plant:", fontSize: 8, color: "#555", margin: [0, 0, 0, 6] },
                                                { text: sPlant, fontSize: 36, bold: true }
                                            ],
                                            margin: [10, 10, 10, 10]
                                        },
                                        {
                                            stack: [
                                                { text: "Remarks:", fontSize: 8, color: "#555", margin: [0, 0, 0, 6] }
                                            ],
                                            margin: [10, 10, 10, 10]
                                        }
                                    ]
                                ]
                            },
                            layout: {
                                hLineWidth: function () { return 1.5; },
                                vLineWidth: function () { return 1.5; },
                                hLineColor: function () { return "#222"; },
                                vLineColor: function () { return "#222"; }
                            }
                        }
                    ],
                    defaultStyle: { font: "Roboto" }
                };

                window.pdfMake.createPdf(oDocDef).download("DeliveryLabel_" + sExtDelivId + ".pdf");
            });
        },

        _loadQrCodeJs: function (fnCallback) {
            if (window.QRCode) { fnCallback(); return; }
            var oScript = document.createElement("script");
            oScript.src = "https://cdn.jsdelivr.net/npm/qrcodejs@1.0.0/qrcode.min.js";
            oScript.onload = fnCallback;
            document.head.appendChild(oScript);
        },

        _loadPdfMake: function (fnCallback) {
            if (window.pdfMake && window.pdfMake.vfs) { fnCallback(); return; }
            var oScript1 = document.createElement("script");
            oScript1.src = "https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js";
            oScript1.onload = function () {
                var oScript2 = document.createElement("script");
                oScript2.src = "https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js";
                oScript2.onload = fnCallback;
                document.head.appendChild(oScript2);
            };
            document.head.appendChild(oScript1);
        },

        // =========================================================================
        // PRIVATE: Session State Persistence (cross-app navigation round-trip)
        // =========================================================================

        _saveSessionState: function (oData) {
            try {
                sessionStorage.setItem("delivuptime_state", JSON.stringify(oData));
            } catch (e) { /* quota or private-mode — ignore */ }
        },

        _restoreSessionState: function () {
            try {
                var sState = sessionStorage.getItem("delivuptime_state");
                if (sState) {
                    sessionStorage.removeItem("delivuptime_state");
                    return JSON.parse(sState);
                }
            } catch (e) { /* ignore */ }
            return null;
        },

        // =========================================================================
        // PRIVATE: Reset
        // =========================================================================

        _resetToInitialState: function () {
            var oViewModel = this.getView().getModel("viewModel");
            oViewModel.setData(Object.assign({}, INITIAL_STATE));
            this._oDeliveryContext  = null;
            this._aDeliveryContexts = [];
        }

    });
});