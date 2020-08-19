/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/functions/functions.ts");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/functions/functions.ts":
/*!************************************!*\
  !*** ./src/functions/functions.ts ***!
  \************************************/
/*! no static exports found */
/***/ (function(module, exports) {

var __awaiter = this && this.__awaiter || function (thisArg, _arguments, P, generator) {
  return new (P || (P = Promise))(function (resolve, reject) {
    function fulfilled(value) {
      try {
        step(generator.next(value));
      } catch (e) {
        reject(e);
      }
    }

    function rejected(value) {
      try {
        step(generator["throw"](value));
      } catch (e) {
        reject(e);
      }
    }

    function step(result) {
      result.done ? resolve(result.value) : new P(function (resolve) {
        resolve(result.value);
      }).then(fulfilled, rejected);
    }

    step((generator = generator.apply(thisArg, _arguments || [])).next());
  });
};

var __generator = this && this.__generator || function (thisArg, body) {
  var _ = {
    label: 0,
    sent: function () {
      if (t[0] & 1) throw t[1];
      return t[1];
    },
    trys: [],
    ops: []
  },
      f,
      y,
      t,
      g;
  return g = {
    next: verb(0),
    "throw": verb(1),
    "return": verb(2)
  }, typeof Symbol === "function" && (g[Symbol.iterator] = function () {
    return this;
  }), g;

  function verb(n) {
    return function (v) {
      return step([n, v]);
    };
  }

  function step(op) {
    if (f) throw new TypeError("Generator is already executing.");

    while (_) try {
      if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
      if (y = 0, t) op = [op[0] & 2, t.value];

      switch (op[0]) {
        case 0:
        case 1:
          t = op;
          break;

        case 4:
          _.label++;
          return {
            value: op[1],
            done: false
          };

        case 5:
          _.label++;
          y = op[1];
          op = [0];
          continue;

        case 7:
          op = _.ops.pop();

          _.trys.pop();

          continue;

        default:
          if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) {
            _ = 0;
            continue;
          }

          if (op[0] === 3 && (!t || op[1] > t[0] && op[1] < t[3])) {
            _.label = op[1];
            break;
          }

          if (op[0] === 6 && _.label < t[1]) {
            _.label = t[1];
            t = op;
            break;
          }

          if (t && _.label < t[2]) {
            _.label = t[2];

            _.ops.push(op);

            break;
          }

          if (t[2]) _.ops.pop();

          _.trys.pop();

          continue;
      }

      op = body.call(thisArg, _);
    } catch (e) {
      op = [6, e];
      y = 0;
    } finally {
      f = t = 0;
    }

    if (op[0] & 5) throw op[1];
    return {
      value: op[0] ? op[1] : void 0,
      done: true
    };
  }
};
/**
 * Adds two numbers without using batching
 * @CustomFunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */


function addNoBatch(first, second) {
  return first + second;
}
/**
 * Divides two numbers using batching
 * @CustomFunction
 * @param dividend The number being divided
 * @param divisor The number to divide the dividend with
 * @returns The sum of the two numbers
 */


function div2(dividend, divisor) {
  return _pushOperation("div2", [dividend, divisor]);
}
/**
 * Multiplies two numbers together using batching
 * @CustomFunction
 * @param first First number to multiply
 * @param second Second number to multiply
 * @returns The product of the two numbers
 */


function mul2(first, second) {
  return _pushOperation("mul2", [first, second]);
}
/**
 * Defines the implementation of the custom functions
 * for the function id defined in the metadata file (functions.json).
 */


CustomFunctions.associate("ADDNOBATCH", addNoBatch);
CustomFunctions.associate("DIV2", div2);
CustomFunctions.associate("MUL2", mul2);
var _batch = [];
var _isBatchedRequestScheduled = false; // This function encloses your custom functions as individual entries,
// which have some additional properties so you can keep track of whether or not
// a request has been resolved or rejected.

function _pushOperation(op, args) {
  // Create an entry for your custom function.
  var invocationEntry = {
    operation: op,
    args: args,
    resolve: undefined,
    reject: undefined
  }; // Create a unique promise for this invocation,
  // and save its resolve and reject functions into the invocation entry.

  var promise = new Promise(function (resolve, reject) {
    invocationEntry.resolve = resolve;
    invocationEntry.reject = reject;
  }); // Push the invocation entry into the next batch.

  _batch.push(invocationEntry); // If a remote request hasn't been scheduled yet,
  // schedule it after a certain timeout, e.g. 100 ms.


  if (!_isBatchedRequestScheduled) {
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  } // Return the promise for this invocation.


  return promise;
} // This is a private helper function, used only within your custom function add-in.
// You wouldn't call _makeRemoteRequest in Excel, for example.
// This function makes a request for remote processing of the whole batch,
// and matches the response batch to the request batch.


function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  var batchCopy = _batch.splice(0, _batch.length);

  _isBatchedRequestScheduled = false; // Build a simpler request batch that only contains the arguments for each invocation.

  var requestBatch = batchCopy.map(function (item) {
    return {
      operation: item.operation,
      args: item.args
    };
  }); // Make the remote request.

  _fetchFromRemoteService(requestBatch).then(function (responseBatch) {
    // Match each value from the response batch to its corresponding invocation entry from the request batch,
    // and resolve the invocation promise with its corresponding response value.
    responseBatch.forEach(function (response, index) {
      if (response.error) {
        batchCopy[index].reject(new Error(response.error));
      } else {
        console.log(response);
        batchCopy[index].resolve(response.result);
      }
    });
  });
} // --------------------- A public API ------------------------------
// This function simulates the work of a remote service. Because each service
// differs, you will need to modify this function appropriately to work with the service you are using. 
// This function takes a batch of argument sets and returns a [promise of] batch of values.
// NOTE: When implementing this function on a server, also apply an appropriate authentication mechanism
//       to ensure only the correct callers can access it.


function _fetchFromRemoteService(requestBatch) {
  return __awaiter(this, void 0, void 0, function () {
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          // Simulate a slow network request to the server;
          return [4
          /*yield*/
          , pause(1000)];

        case 1:
          // Simulate a slow network request to the server;
          _a.sent();

          return [2
          /*return*/
          , requestBatch.map(function (request) {
            var operation = request.operation,
                args = request.args;

            try {
              if (operation === "div2") {
                // Divide the first argument by the second argument.
                return {
                  result: args[0] / args[1]
                };
              } else if (operation === "mul2") {
                // Multiply the arguments for the given entry.
                var myresult = args[0] * args[1];
                console.log(myresult);
                return {
                  result: myresult
                };
              } else {
                return {
                  error: "Operation not supported: " + operation
                };
              }
            } catch (error) {
              return {
                error: "Operation failed: " + operation
              };
            }
          })];
      }
    });
  });
}

function pause(ms) {
  return new Promise(function (resolve) {
    return setTimeout(resolve, ms);
  });
}

/***/ })

/******/ });
//# sourceMappingURL=functions.js.map