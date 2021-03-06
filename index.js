"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
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
var logging_1 = require("@pnp/logging");
var MSGraphHelper = /** @class */ (function () {
    function MSGraphHelper() {
    }
    MSGraphHelper.Init = function (msGraphClientFactory) {
        return __awaiter(this, void 0, void 0, function () {
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _a = this;
                        return [4 /*yield*/, msGraphClientFactory.getClient()];
                    case 1:
                        _a._graphClient = _b.sent();
                        logging_1.Logger.subscribe(new logging_1.ConsoleListener());
                        logging_1.Logger.activeLogLevel = 1 /* Info */;
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get
     *
     * @param {string} apiUrl API url
     * @param {string} version Version (default to v1.0)
     * @param {Array<string>} selectProperties Select properties
     * @param {string} filter Filter
     * @param {number} top Number of items to retrieve
     * @param {Array<string>} expand Expand
     */
    MSGraphHelper.Get = function (apiUrl, version, selectProperties, filter, top, expand) {
        if (version === void 0) { version = "v1.0"; }
        return __awaiter(this, void 0, void 0, function () {
            var values, query, response, nextLink, error_1, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 8, , 9]);
                        values = [];
                        query = this._graphClient.api(apiUrl).version(version);
                        if (selectProperties && selectProperties.length > 0) {
                            query = query.select(selectProperties);
                        }
                        if (filter && filter.length > 0) {
                            query = query.filter(filter);
                        }
                        if (top) {
                            query = query.top(top);
                        }
                        if (expand) {
                            query = query.expand(expand);
                        }
                        logging_1.Logger.log({ message: "(MSGraphHelper) Get", data: { urlComponents: query.urlComponents }, level: 1 /* Info */ });
                        return [4 /*yield*/, query.get()];
                    case 1:
                        response = _a.sent();
                        if (response.value && response.value.length > 0) {
                            values.push.apply(values, response.value);
                        }
                        nextLink = response["@odata.nextLink"];
                        if (!nextLink) return [3 /*break*/, 7];
                        _a.label = 2;
                    case 2:
                        if (!true) return [3 /*break*/, 7];
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, 5, , 6]);
                        query.parsePath(nextLink);
                        return [4 /*yield*/, query.get()];
                    case 4:
                        response = _a.sent();
                        nextLink = response["@odata.nextLink"];
                        if (response.value && response.value.length > 0) {
                            values.push.apply(values, response.value);
                        }
                        if (!nextLink) {
                            return [3 /*break*/, 7];
                        }
                        return [3 /*break*/, 6];
                    case 5:
                        error_1 = _a.sent();
                        throw error_1;
                    case 6: return [3 /*break*/, 2];
                    case 7: return [2 /*return*/, values];
                    case 8:
                        error_2 = _a.sent();
                        throw error_2;
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Patch
     *
     * @param {string} apiUrl API url
     * @param {string} version Version (default to v1.0)
     * @param {any} content Content
     */
    MSGraphHelper.Patch = function (apiUrl, version, content) {
        if (version === void 0) { version = "v1.0"; }
        return __awaiter(this, void 0, void 0, function () {
            var p;
            var _this = this;
            return __generator(this, function (_a) {
                p = new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
                    var query, callback;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0:
                                if (typeof (content) === "object") {
                                    content = JSON.stringify(content);
                                }
                                query = this._graphClient.api(apiUrl).version(version);
                                callback = function (error, _response, rawResponse) {
                                    if (error) {
                                        reject(error);
                                    }
                                    else {
                                        resolve();
                                    }
                                };
                                return [4 /*yield*/, query.update(content, callback)];
                            case 1:
                                _a.sent();
                                return [2 /*return*/];
                        }
                    });
                }); });
                return [2 /*return*/, p];
            });
        });
    };
    /**
     * Post
     *
     * @param {string} apiUrl API url
     * @param {string} version Version (default to v1.0)
     * @param {any} content Content
     */
    MSGraphHelper.Post = function (apiUrl, version, content) {
        if (version === void 0) { version = "v1.0"; }
        return __awaiter(this, void 0, void 0, function () {
            var p;
            var _this = this;
            return __generator(this, function (_a) {
                p = new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
                    var query, callback;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0:
                                if (typeof (content) === "object") {
                                    content = JSON.stringify(content);
                                }
                                query = this._graphClient.api(apiUrl).version(version);
                                callback = function (error, response, rawResponse) {
                                    if (error) {
                                        reject(error);
                                    }
                                    else {
                                        resolve(response);
                                    }
                                };
                                return [4 /*yield*/, query.post(content, callback)];
                            case 1:
                                _a.sent();
                                return [2 /*return*/];
                        }
                    });
                }); });
                return [2 /*return*/, p];
            });
        });
    };
    /**
     * Delete
     *
     * @param {string} apiUrl API url
     * @param {string} version Version (default to v1.0)
     */
    MSGraphHelper.Delete = function (apiUrl, version) {
        if (version === void 0) { version = "v1.0"; }
        return __awaiter(this, void 0, void 0, function () {
            var p;
            var _this = this;
            return __generator(this, function (_a) {
                p = new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
                    var query, callback;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0:
                                query = this._graphClient.api(apiUrl).version(version);
                                callback = function (error, response, rawResponse) {
                                    if (error) {
                                        reject(error);
                                    }
                                    else {
                                        resolve(response);
                                    }
                                };
                                return [4 /*yield*/, query.delete(callback)];
                            case 1:
                                _a.sent();
                                return [2 /*return*/];
                        }
                    });
                }); });
                return [2 /*return*/, p];
            });
        });
    };
    return MSGraphHelper;
}());
exports.default = MSGraphHelper;
//# sourceMappingURL=index.js.map