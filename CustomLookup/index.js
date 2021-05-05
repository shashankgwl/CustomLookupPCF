"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.CustomLookup = void 0;
var react_1 = __importDefault(require("react"));
var react_dom_1 = __importDefault(require("react-dom"));
var CustomLookupCore_1 = __importDefault(require("./components/CustomLookupCore"));
var CustomLookup = /** @class */ (function () {
    function CustomLookup() {
    }
    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    CustomLookup.prototype.init = function (context, notifyOutputChanged, state, container) {
        var _a;
        // Add control initialization code
        this.root = container;
        if (context.parameters.configWRUrl.raw) {
            if (((_a = context.parameters.configWRUrl.raw) === null || _a === void 0 ? void 0 : _a.length) <= 0) {
                Xrm.Utility.alertDialog("Configuration URL is missing, please contact administrator", function () { });
            }
            else {
                var pageContext = {
                    webResourceURL: context.parameters.configWRUrl.raw
                };
                react_dom_1.default.render(react_1.default.createElement(CustomLookupCore_1.default, pageContext), this.root);
            }
        }
    };
    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    CustomLookup.prototype.updateView = function (context) {
        // Add code to update control view
        //ReactDOM.render(React.cre)
    };
    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    CustomLookup.prototype.getOutputs = function () {
        return {};
    };
    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    CustomLookup.prototype.destroy = function () {
        // Add code to cleanup control if necessary
    };
    return CustomLookup;
}());
exports.CustomLookup = CustomLookup;
