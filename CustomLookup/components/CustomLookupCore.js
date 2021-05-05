"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    Object.defineProperty(o, k2, { enumerable: true, get: function() { return m[k]; } });
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var React = __importStar(require("react"));
var lib_1 = require("office-ui-fabric-react/lib");
var react_1 = require("@fluentui/react");
var office_ui_fabric_react_1 = require("office-ui-fabric-react");
var icons_1 = require("@uifabric/icons");
function CustomLookupCore(props) {
    var _this = this;
    var _a;
    var _b = React.useState({}), lookupState = _b[0], setLookupState = _b[1];
    var _c = React.useState([]), cachedLookupData = _c[0], setCachedLookupData = _c[1];
    var _d = React.useState(true), isBusy = _d[0], setIsBusy = _d[1];
    var _e = React.useState({
        entityColumns: [],
        nameAttribute: "name",
        entityToFetch: '',
        lookupAttributeOnPage: ''
    }), jsonContext = _e[0], setJsonContext = _e[1];
    icons_1.initializeIcons();
    var theme = office_ui_fabric_react_1.getTheme();
    var lookupData = {};
    var contentStyles = office_ui_fabric_react_1.mergeStyleSets({
        container: {
            display: 'flex',
            flexFlow: 'column nowrap',
            alignItems: 'stretch',
            width: '80%',
            height: '60%'
        },
        header: [
            theme.fonts.xLargePlus,
            {
                flex: '1 1 auto',
                borderTop: "4px solid " + theme.palette.themePrimary,
                color: theme.palette.neutralPrimary,
                display: 'flex',
                alignItems: 'center',
                fontWeight: react_1.FontWeights.semibold,
                padding: '12px 12px 14px 24px',
            },
        ],
        body: {
            flex: '4 4 auto',
            padding: '0 24px 24px 24px',
            overflowY: 'hidden',
            selectors: {
                p: { margin: '14px 0' },
                'p:first-child': { marginTop: 0 },
                'p:last-child': { marginBottom: 0 },
            },
        },
    });
    var stackItemTokens = {
        margin: "0,0,0,50",
        padding: 10,
    };
    function getWebResource(url) {
        return __awaiter(this, void 0, void 0, function () {
            var response;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, fetch(url)];
                    case 1:
                        response = _a.sent();
                        return [2 /*return*/, response.text()];
                }
            });
        });
    }
    var getFormattedColumn = function (schema) {
        switch (schema.type) {
            case 'lookup':
                return '_' + schema.schemaName + '_value';
                break;
            default:
                return schema.schemaName;
        }
        return '';
    };
    var getDynamicsColumn = function (jsonCol) {
        var obj = [];
        jsonCol.map(function (col) { return obj.push({
            displayName: col.displayName,
            schemaName: col.schemaName,
            type: col.type,
            formattedSchemaName: getFormattedColumn(col)
        }); });
        return obj;
    };
    React.useEffect(function () {
        getWebResource(props.webResourceURL).then(function (json) {
            var jobject = JSON.parse(json);
            var dataContext = {
                entityToFetch: jobject.entityToFetch,
                nameAttribute: jobject.nameAttribute,
                entityColumns: getDynamicsColumn(jobject.columns),
                lookupAttributeOnPage: jobject.lookupAttributeOnPage
            };
            setJsonContext(dataContext);
            var pageAttribute = Xrm.Page.data.entity.attributes.get(dataContext.lookupAttributeOnPage).getValue();
            var state = {
                currentItemText: pageAttribute ? pageAttribute[0].name : '',
                currentItemId: pageAttribute ? pageAttribute[0].id : '',
                isModelOpen: false,
                lookupColumns: Array.from(dataContext.entityColumns, function (col) { return col.displayName; })
                //..lookupData: []
            };
            //Xrm.Utility.alertDialog(state.lookupColumns!.toString(), () => { })
            setLookupState(state);
        }).catch(function (err) {
            localAlert(err);
        });
    }, [props]);
    function localAlert(msg) {
        Xrm.Utility.alertDialog(msg, function () { });
    }
    function timeout(ms) {
        return new Promise(function (resolve) { return setTimeout(resolve, ms); });
    }
    function buildQuery() {
        return __awaiter(this, void 0, void 0, function () {
            var query;
            return __generator(this, function (_a) {
                query = "?$select=";
                jsonContext.entityColumns.map(function (item) {
                    query += item.formattedSchemaName + ',';
                });
                if (query.endsWith(',')) {
                    query = query.substr(0, query.lastIndexOf(',')); //+ '&$top=2'
                }
                //localAlert(query);
                return [2 /*return*/, query];
            });
        });
    }
    //id: records.entities[i][jsonContext.entityToFetch + 'id'],
    //text: records.entities[i][jsonContext.entityColumns[j].formattedSchemaName + "@OData.Community.Display.V1.FormattedValue"]
    var getDataFromCE = function (query) { return __awaiter(_this, void 0, void 0, function () {
        var data, tempRow, records, i, j;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    data = [];
                    tempRow = {};
                    return [4 /*yield*/, Xrm.WebApi.online.retrieveMultipleRecords(jsonContext.entityToFetch, query)
                        //localAlert(records.entities.length.toString());
                    ];
                case 1:
                    records = _a.sent();
                    //localAlert(records.entities.length.toString());
                    for (i = 0; i < records.entities.length; i++) {
                        tempRow = {};
                        tempRow.id = records.entities[i][jsonContext.entityToFetch + 'id'];
                        for (j = 0; j < jsonContext.entityColumns.length; j++) {
                            if (jsonContext.entityColumns[j].type === 'lookup') {
                                Object.defineProperty(tempRow, jsonContext.entityColumns[j].schemaName, {
                                    value: records.entities[i][jsonContext.entityColumns[j].formattedSchemaName + "@OData.Community.Display.V1.FormattedValue"],
                                    writable: true,
                                    configurable: true,
                                    enumerable: true
                                });
                            }
                            else {
                                Object.defineProperty(tempRow, jsonContext.entityColumns[j].schemaName, {
                                    value: Object.values(tempRow)[j] = records.entities[i][jsonContext.entityColumns[j].formattedSchemaName],
                                    writable: true,
                                    configurable: true,
                                    enumerable: true
                                });
                            }
                        }
                        data.push(tempRow);
                    }
                    console.log(data);
                    return [2 /*return*/, data];
            }
        });
    }); };
    var onDialogOpen = function () { return __awaiter(_this, void 0, void 0, function () {
        var lookupData, _a, state;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    setIsBusy(true);
                    setLookupState({
                        currentItemText: lookupState.currentItemText,
                        isModelOpen: true,
                        lookupData: [],
                        lookupColumns: lookupState.lookupColumns
                    });
                    _a = getDataFromCE;
                    return [4 /*yield*/, buildQuery()];
                case 1: return [4 /*yield*/, _a.apply(void 0, [_b.sent()])];
                case 2:
                    lookupData = _b.sent();
                    state = {
                        isModelOpen: true,
                        lookupData: lookupData,
                        lookupColumns: lookupState.lookupColumns,
                        currentItemText: lookupState.currentItemText
                    };
                    setLookupState(state);
                    setCachedLookupData(lookupData || []);
                    setIsBusy(false);
                    return [2 /*return*/];
            }
        });
    }); };
    function getListColumns() {
        var cols = [];
        cols.push({
            key: "id",
            minWidth: 80,
            maxWidth: 180,
            name: "Select",
            isResizable: true,
            isCollapsible: true,
            data: 'string'
        });
        if (!lookupState.lookupColumns)
            return [];
        for (var i = 0; i < jsonContext.entityColumns.length; i++) {
            cols.push({
                key: jsonContext.entityColumns[i].schemaName,
                minWidth: 80,
                maxWidth: 180,
                name: lookupState.lookupColumns[i],
                isResizable: true,
                isCollapsible: true,
                data: 'string'
            });
        }
        return cols;
    }
    //const getData = (): ILookupData[] => {
    //    var data: ILookupData[] = []
    //    for (var i = 0; i < 50; i++) {
    //        data.push({
    //            id: i.toString(),
    //            text: "text" + i
    //        });
    //    }
    //    return data;
    //}
    var onTextFilter = function (ev, tx) {
        var element = ev;
        try {
            var filtered = tx ? cachedLookupData.filter(function (item) { return item[element.target.name] &&
                item[element.target.name].toLowerCase().indexOf(tx.toLowerCase()) > -1; }) : cachedLookupData;
            //localAlert(filtered.length.toString())
            setLookupState({
                currentItemText: lookupState.currentItemText,
                lookupData: filtered,
                isModelOpen: true,
                lookupColumns: lookupState.lookupColumns
            });
        }
        catch (error) {
            localAlert(error);
        }
    };
    function _renderItemColumn(item, index, column) {
        if ((column === null || column === void 0 ? void 0 : column.key) === "id") {
            return (React.createElement(lib_1.Link, { onClick: function () {
                    var lookupValue = new Array();
                    lookupValue[0] = new Object();
                    lookupValue[0].id = item === null || item === void 0 ? void 0 : item.id; // GUID of the lookup id
                    lookupValue[0].name = item[jsonContext.nameAttribute]; // Name of the lookup
                    lookupValue[0].entityType = jsonContext.entityToFetch; //Entity Type of the lookup entity
                    Xrm.Page.getAttribute(jsonContext.lookupAttributeOnPage).setValue(lookupValue); // You need to replace the lookup field 
                    setLookupState({
                        currentItemText: item[jsonContext.nameAttribute],
                        isModelOpen: false,
                        lookupColumns: lookupState.lookupColumns,
                        currentItemId: item === null || item === void 0 ? void 0 : item.id
                    });
                } }, "Select"));
        }
        return (React.createElement(lib_1.Text, null, item[column.key]));
    }
    var onDialogDismiss = function () {
        var state = {
            isModelOpen: false,
            currentItemText: lookupState.currentItemText,
            lookupColumns: lookupState.lookupColumns
        };
        setLookupState(state);
    };
    return (React.createElement(lib_1.Fabric, null,
        React.createElement(react_1.Stack, { horizontal: true, style: { border: 2 } },
            React.createElement(lib_1.StackItem, null,
                React.createElement(lib_1.Link, { onClick: function () {
                        Xrm.Utility.openEntityForm(jsonContext.entityToFetch, lookupState.currentItemId, undefined, { openInNewWindow: true });
                    } },
                    lookupState.currentItemText,
                    "  ")),
            React.createElement(lib_1.StackItem, { style: { marginLeft: 30 } },
                React.createElement(react_1.IconButton, { iconProps: { iconName: "Search" }, onClick: onDialogOpen }))),
        React.createElement(lib_1.Modal, { isOpen: lookupState.isModelOpen, containerClassName: contentStyles.container, onDismiss: onDialogDismiss },
            React.createElement(react_1.Stack, { horizontal: true, tokens: { childrenGap: 5 } }, (_a = jsonContext.entityColumns) === null || _a === void 0 ? void 0 :
                _a.map(function (columnName) {
                    return (React.createElement(react_1.Stack.Item, null,
                        React.createElement(lib_1.TextField, { name: columnName.schemaName, disabled: isBusy, label: columnName.displayName + ' Filter', onChange: onTextFilter })));
                }),
                React.createElement(react_1.Stack.Item, { align: "end", tokens: { margin: "0,0,0,50" } },
                    React.createElement(lib_1.PrimaryButton, { onClick: function () {
                            Xrm.Utility.openEntityForm(jsonContext.entityToFetch, undefined, undefined, { openInNewWindow: true });
                        }, marginWidth: 25, text: "Create new " + jsonContext.entityToFetch }))),
            React.createElement("div", null,
                React.createElement(lib_1.ShimmeredDetailsList, { selectionMode: lib_1.SelectionMode.none, compact: true, columns: getListColumns(), items: lookupState.lookupData ? lookupState.lookupData : [], onRenderItemColumn: _renderItemColumn, enableShimmer: isBusy })))));
}
exports.default = CustomLookupCore;
